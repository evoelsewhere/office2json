import { sanitizeHyperlink } from '../../../common/text/html';
import type { ResolvedXlsxResourceLimits } from '../internal/resource-limits';
import { parseXlsxCellReference } from '../internal/cell-reference';
import type { XlsxCellValue, XlsxHyperlinkTarget } from '../types';
import { canonicalXlsxJson } from './canonical-json';
import { XlsxWriteError } from './errors';
import type { ResolvedXlsxWriteLimits, XlsxEditOperation } from './types';
import { validateXlsxOperationStyle } from './style-validation';
import { writeLimitFailure } from './write-limits';

export type XlsxCellEditOperation = Extract<
  XlsxEditOperation,
  {
    kind:
      | 'clear-cell'
      | 'delete-columns'
      | 'delete-rows'
      | 'insert-columns'
      | 'insert-rows'
      | 'set-cell'
      | 'set-cell-style'
      | 'set-column'
      | 'set-hyperlink'
      | 'set-row';
  }
>;

const ERROR_CODES = new Set([
  '#BLOCKED!',
  '#BUSY!',
  '#CALC!',
  '#CONNECT!',
  '#DIV/0!',
  '#FIELD!',
  '#GETTING_DATA',
  '#N/A',
  '#NAME?',
  '#NULL!',
  '#NUM!',
  '#REF!',
  '#SPILL!',
  '#UNKNOWN!',
  '#VALUE!',
]);
const KNOWN_OPERATIONS = new Set([
  'add-worksheet',
  'delete-worksheet',
  'rename-worksheet',
]);
const OPERATION_ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/u;
const SHA256_PATTERN = /^[0-9a-f]{64}$/u;
const SHEET_KEY_PATTERN = /^xlsx:sheet:[0-9a-f]{32}$/u;

function invalid(message: string, operationId?: string): never {
  throw new XlsxWriteError('invalid-roundtrip-json', message, {
    ...(operationId === undefined ? {} : { operationId }),
  });
}

function unsupported(
  message: string,
  operationId: string | undefined,
  featureClass?: string,
): never {
  throw new XlsxWriteError('unsupported-edit-operation', message, {
    ...(featureClass === undefined ? {} : { featureClass }),
    ...(operationId === undefined ? {} : { operationId }),
  });
}

function exactKeys(
  value: Record<string, unknown>,
  required: readonly string[],
  optional?: readonly string[],
): boolean {
  const names = Object.keys(value);
  if (optional === undefined) {
    return (
      names.length === required.length &&
      required.every((key) => Object.hasOwn(value, key))
    );
  }
  return (
    required.every((key) => Object.hasOwn(value, key)) &&
    names.every((key) => required.includes(key) || optional.includes(key))
  );
}

function plainRecord(value: unknown): value is Record<string, unknown> {
  if (value === null || value === undefined) return false;
  return Object.getPrototypeOf(value) === Object.prototype;
}

function operationId(value: unknown): string {
  if (typeof value !== 'string' || !OPERATION_ID_PATTERN.test(value)) {
    invalid('XLSX operation ID is invalid');
  }
  return value;
}

function validateSheetCommon(
  operation: Record<string, unknown>,
  id: string,
): { ifMatch?: string; operationId: string; sheetKey: string } {
  if (
    typeof operation.sheetKey !== 'string' ||
    !SHEET_KEY_PATTERN.test(operation.sheetKey)
  ) {
    invalid('XLSX operation sheet key is invalid', id);
  }
  if (
    operation.ifMatch !== undefined &&
    (typeof operation.ifMatch !== 'string' ||
      !SHA256_PATTERN.test(operation.ifMatch))
  ) {
    invalid('XLSX operation precondition hash is invalid', id);
  }
  return {
    ...(operation.ifMatch === undefined ? {} : { ifMatch: operation.ifMatch }),
    operationId: id,
    sheetKey: operation.sheetKey,
  };
}

function validateCommon(
  operation: Record<string, unknown>,
  id: string,
): { cell: string; ifMatch?: string; operationId: string; sheetKey: string } {
  const common = validateSheetCommon(operation, id);
  const parsed = parseXlsxCellReference(operation.cell);
  if (
    !parsed ||
    parsed.absoluteColumn ||
    parsed.absoluteRow ||
    operation.cell !== parsed.address
  ) {
    invalid('XLSX operation cell reference is invalid', id);
  }
  return { ...common, cell: parsed.address };
}

function boundedIndex(
  value: unknown,
  limit: number,
  message: string,
  id: string,
): number {
  if (
    !Number.isSafeInteger(value) ||
    (value as number) < 1 ||
    (value as number) > limit
  ) {
    invalid(message, id);
  }
  return value as number;
}

function boundedDimension(
  value: unknown,
  maximum: number,
  message: string,
  id: string,
): number {
  if (
    typeof value !== 'number' ||
    !Number.isFinite(value) ||
    value < 0 ||
    value > maximum
  ) {
    invalid(message, id);
  }
  return value === 0 ? 0 : value;
}

function validateTextValue(
  value: Record<string, unknown>,
  id: string,
): XlsxCellValue {
  if (!exactKeys(value, ['kind', 'text'], ['runs'])) {
    invalid('XLSX text cell value shape is invalid', id);
  }
  if (typeof value.text !== 'string') {
    invalid('XLSX text cell value is invalid', id);
  }
  if (value.runs !== undefined) {
    unsupported(
      'XLSX cell editing does not yet support rich text runs',
      id,
      'rich-text',
    );
  }
  return { kind: 'text', text: value.text };
}

function validateCellValue(value: unknown, id: string): XlsxCellValue {
  if (!plainRecord(value)) {
    invalid('XLSX cell value shape is invalid', id);
  }
  const record = value;
  if (record.kind === 'text') return validateTextValue(record, id);
  if (record.kind === 'number') {
    if (
      !exactKeys(record, ['kind', 'value']) ||
      typeof record.value !== 'number' ||
      !Number.isFinite(record.value)
    ) {
      invalid('XLSX number cell value is invalid', id);
    }
    return { kind: 'number', value: record.value === 0 ? 0 : record.value };
  }
  if (record.kind === 'boolean') {
    if (
      !exactKeys(record, ['kind', 'value']) ||
      typeof record.value !== 'boolean'
    ) {
      invalid('XLSX boolean cell value is invalid', id);
    }
    return { kind: 'boolean', value: record.value };
  }
  if (record.kind === 'error') {
    if (
      !exactKeys(record, ['code', 'kind']) ||
      typeof record.code !== 'string' ||
      !ERROR_CODES.has(record.code)
    ) {
      invalid('XLSX error cell value is invalid', id);
    }
    return { code: record.code, kind: 'error' };
  }
  if (record.kind === 'date') {
    unsupported(
      'XLSX cell editing does not yet support date values',
      id,
      'date-value',
    );
  }
  invalid('XLSX cell value kind is invalid', id);
}

function validateContent(
  value: unknown,
  id: string,
  readerLimits: ResolvedXlsxResourceLimits,
): Extract<XlsxEditOperation, { kind: 'set-cell' }>['content'] {
  if (!plainRecord(value)) {
    invalid('XLSX set-cell content shape is invalid', id);
  }
  const content = value;
  if (content.kind === 'formula') {
    if (
      !exactKeys(content, ['expression', 'kind']) ||
      typeof content.expression !== 'string' ||
      content.expression.length === 0 ||
      content.expression.startsWith('=')
    ) {
      invalid('XLSX set-cell formula is invalid', id);
    }
    if (content.expression.length > readerLimits.maxFormulaCharacters) {
      throw new XlsxWriteError(
        'resource-limit-exceeded',
        'XLSX operation formula exceeds its character limit',
        {
          actual: content.expression.length,
          limit: readerLimits.maxFormulaCharacters,
          limitName: 'maxFormulaCharacters',
          operationId: id,
        },
      );
    }
    return { expression: content.expression, kind: 'formula' };
  }
  if (content.kind === 'value') {
    if (!exactKeys(content, ['kind', 'value'])) {
      invalid('XLSX set-cell value content shape is invalid', id);
    }
    return { kind: 'value', value: validateCellValue(content.value, id) };
  }
  invalid('XLSX set-cell content kind is invalid', id);
}

function validateHyperlinkTarget(
  value: unknown,
  id: string,
  readerLimits: ResolvedXlsxResourceLimits,
): XlsxHyperlinkTarget | null {
  if (value === null) return null;
  if (!plainRecord(value)) {
    invalid('XLSX set-hyperlink target shape is invalid', id);
  }
  if (value.kind === 'internal') {
    if (
      !exactKeys(value, ['kind', 'location']) ||
      typeof value.location !== 'string' ||
      value.location.length === 0
    ) {
      invalid('XLSX internal hyperlink target is invalid', id);
    }
    if (value.location.length > readerLimits.maxTextCharacters) {
      throw new XlsxWriteError(
        'resource-limit-exceeded',
        'XLSX hyperlink location exceeds its text limit',
        {
          actual: value.location.length,
          limit: readerLimits.maxTextCharacters,
          limitName: 'maxTextCharacters',
          operationId: id,
        },
      );
    }
    return { kind: 'internal', location: value.location };
  }
  if (value.kind === 'external') {
    if (
      !exactKeys(value, ['kind', 'url'], ['location']) ||
      typeof value.url !== 'string' ||
      (value.location !== undefined && typeof value.location !== 'string')
    ) {
      invalid('XLSX external hyperlink target is invalid', id);
    }
    const safe = sanitizeHyperlink(value.url);
    if (!safe || safe !== value.url) {
      throw new XlsxWriteError(
        'preservation-conflict',
        'XLSX external hyperlink protocol or lexical form is unsafe',
        { featureClass: 'hyperlink-protocol', operationId: id },
      );
    }
    const url = new URL(safe);
    if (url.username !== '' || url.password !== '') {
      throw new XlsxWriteError(
        'preservation-conflict',
        'XLSX external hyperlink credentials are not allowed',
        { featureClass: 'hyperlink-credentials', operationId: id },
      );
    }
    if (url.toString() !== value.url) {
      invalid('XLSX external hyperlink URL must be canonical', id);
    }
    const textCharacters =
      value.url.length +
      (value.location === undefined ? 0 : value.location.length);
    if (textCharacters > readerLimits.maxTextCharacters) {
      throw new XlsxWriteError(
        'resource-limit-exceeded',
        'XLSX hyperlink target exceeds its text limit',
        {
          actual: textCharacters,
          limit: readerLimits.maxTextCharacters,
          limitName: 'maxTextCharacters',
          operationId: id,
        },
      );
    }
    return {
      kind: 'external',
      ...(value.location === undefined ? {} : { location: value.location }),
      url: value.url,
    };
  }
  invalid('XLSX hyperlink target kind is invalid', id);
}

function operationBytes(operation: unknown): number {
  return new TextEncoder().encode(canonicalXlsxJson(operation)).byteLength;
}

export function validateXlsxCellOperations(
  value: unknown,
  writeLimits: ResolvedXlsxWriteLimits,
  readerLimits: ResolvedXlsxResourceLimits,
): XlsxCellEditOperation[] {
  if (!Array.isArray(value)) {
    invalid('XLSX round-trip operations must be an array');
  }
  if (value.length > writeLimits.maxOperations) {
    writeLimitFailure('maxOperations', value.length, writeLimits.maxOperations);
  }
  const ids = new Set<string>();
  const operations: XlsxCellEditOperation[] = [];
  for (const candidate of value) {
    if (!plainRecord(candidate)) {
      invalid('XLSX operation shape is invalid');
    }
    const operation = candidate;
    const id = operationId(operation.operationId);
    if (ids.has(id)) {
      invalid('XLSX operation IDs must be unique', id);
    }
    ids.add(id);
    if (operation.kind === 'clear-cell') {
      if (
        !exactKeys(
          operation,
          ['cell', 'kind', 'operationId', 'sheetKey'],
          ['ifMatch'],
        )
      ) {
        invalid('XLSX clear-cell operation shape is invalid', id);
      }
      const common = validateCommon(operation, id);
      operations.push({
        ...common,
        kind: 'clear-cell',
      });
      continue;
    }
    if (operation.kind === 'set-cell') {
      if (
        !exactKeys(
          operation,
          ['cell', 'content', 'kind', 'operationId', 'sheetKey'],
          ['ifMatch'],
        )
      ) {
        invalid('XLSX set-cell operation shape is invalid', id);
      }
      const common = validateCommon(operation, id);
      const content = validateContent(operation.content, id, readerLimits);
      operations.push({
        ...common,
        content,
        kind: 'set-cell',
      });
      continue;
    }
    if (operation.kind === 'set-cell-style') {
      if (
        !exactKeys(
          operation,
          ['cell', 'kind', 'operationId', 'sheetKey', 'style'],
          ['ifMatch'],
        )
      ) {
        invalid('XLSX set-cell-style operation shape is invalid', id);
      }
      const common = validateCommon(operation, id);
      operations.push({
        ...common,
        kind: 'set-cell-style',
        style: validateXlsxOperationStyle(operation.style, id),
      });
      continue;
    }
    if (operation.kind === 'set-hyperlink') {
      if (
        !exactKeys(
          operation,
          ['cell', 'kind', 'operationId', 'sheetKey', 'target'],
          ['ifMatch'],
        )
      ) {
        invalid('XLSX set-hyperlink operation shape is invalid', id);
      }
      const common = validateCommon(operation, id);
      operations.push({
        ...common,
        kind: 'set-hyperlink',
        target: validateHyperlinkTarget(operation.target, id, readerLimits),
      });
      continue;
    }
    if (operation.kind === 'set-row') {
      if (
        !exactKeys(
          operation,
          ['kind', 'operationId', 'row', 'sheetKey'],
          ['height', 'hidden', 'ifMatch'],
        ) ||
        (operation.height === undefined && operation.hidden === undefined)
      ) {
        invalid('XLSX set-row operation shape is invalid', id);
      }
      if (
        operation.hidden !== undefined &&
        typeof operation.hidden !== 'boolean'
      ) {
        invalid('XLSX set-row hidden value is invalid', id);
      }
      operations.push({
        ...validateSheetCommon(operation, id),
        ...(operation.height === undefined
          ? {}
          : {
              height: boundedDimension(
                operation.height,
                409,
                'XLSX set-row height is invalid',
                id,
              ),
            }),
        ...(operation.hidden === undefined ? {} : { hidden: operation.hidden }),
        kind: 'set-row',
        row: boundedIndex(
          operation.row,
          readerLimits.maxRowsPerWorksheet,
          'XLSX set-row index is invalid',
          id,
        ),
      });
      continue;
    }
    if (operation.kind === 'set-column') {
      if (
        !exactKeys(
          operation,
          ['end', 'kind', 'operationId', 'sheetKey', 'start'],
          ['hidden', 'ifMatch', 'width'],
        ) ||
        (operation.hidden === undefined && operation.width === undefined)
      ) {
        invalid('XLSX set-column operation shape is invalid', id);
      }
      if (
        operation.hidden !== undefined &&
        typeof operation.hidden !== 'boolean'
      ) {
        invalid('XLSX set-column hidden value is invalid', id);
      }
      const start = boundedIndex(
        operation.start,
        readerLimits.maxColumnsPerWorksheet,
        'XLSX set-column start is invalid',
        id,
      );
      const end = boundedIndex(
        operation.end,
        readerLimits.maxColumnsPerWorksheet,
        'XLSX set-column end is invalid',
        id,
      );
      if (start > end) invalid('XLSX set-column range is invalid', id);
      operations.push({
        ...validateSheetCommon(operation, id),
        end,
        ...(operation.hidden === undefined ? {} : { hidden: operation.hidden }),
        kind: 'set-column',
        start,
        ...(operation.width === undefined
          ? {}
          : {
              width: boundedDimension(
                operation.width,
                255,
                'XLSX set-column width is invalid',
                id,
              ),
            }),
      });
      continue;
    }
    if (
      operation.kind === 'delete-columns' ||
      operation.kind === 'delete-rows' ||
      operation.kind === 'insert-columns' ||
      operation.kind === 'insert-rows'
    ) {
      if (
        !exactKeys(
          operation,
          ['count', 'index', 'kind', 'operationId', 'sheetKey'],
          ['ifMatch'],
        )
      ) {
        invalid('XLSX structural operation shape is invalid', id);
      }
      const rowOperation =
        operation.kind === 'delete-rows' || operation.kind === 'insert-rows';
      const limit = rowOperation
        ? readerLimits.maxRowsPerWorksheet
        : readerLimits.maxColumnsPerWorksheet;
      const index = boundedIndex(
        operation.index,
        limit,
        'XLSX structural operation index is invalid',
        id,
      );
      const count = boundedIndex(
        operation.count,
        limit,
        'XLSX structural operation count is invalid',
        id,
      );
      if (count > limit - index + 1) {
        invalid('XLSX structural operation range is invalid', id);
      }
      operations.push({
        ...validateSheetCommon(operation, id),
        count,
        index,
        kind: operation.kind,
      });
      continue;
    }
    if (
      typeof operation.kind === 'string' &&
      KNOWN_OPERATIONS.has(operation.kind)
    ) {
      unsupported(
        `XLSX operation ${operation.kind} is not supported by this profile`,
        id,
        operation.kind,
      );
    }
    invalid('XLSX operation kind is invalid', id);
  }
  let totalBytes = 0;
  let totalFormulaCharacters = 0;
  let totalTextCharacters = 0;
  for (const operation of operations) {
    const bytes = operationBytes(operation);
    if (bytes > writeLimits.maxOperationBytes) {
      throw new XlsxWriteError(
        'resource-limit-exceeded',
        'XLSX operation exceeds its byte limit',
        {
          actual: bytes,
          limit: writeLimits.maxOperationBytes,
          limitName: 'maxOperationBytes',
          operationId: operation.operationId,
        },
      );
    }
    totalBytes += bytes;
    if (totalBytes > writeLimits.maxTotalOperationBytes) {
      throw new XlsxWriteError(
        'resource-limit-exceeded',
        'XLSX operations exceed their total byte limit',
        {
          actual: totalBytes,
          limit: writeLimits.maxTotalOperationBytes,
          limitName: 'maxTotalOperationBytes',
          operationId: operation.operationId,
        },
      );
    }
    if (operation.kind === 'set-cell' && operation.content.kind === 'formula') {
      totalFormulaCharacters += operation.content.expression.length;
      if (totalFormulaCharacters > readerLimits.maxTotalFormulaCharacters) {
        throw new XlsxWriteError(
          'resource-limit-exceeded',
          'XLSX operations exceed their total formula character limit',
          {
            actual: totalFormulaCharacters,
            limit: readerLimits.maxTotalFormulaCharacters,
            limitName: 'maxTotalFormulaCharacters',
            operationId: operation.operationId,
          },
        );
      }
    }
    if (
      operation.kind === 'set-cell' &&
      operation.content.kind === 'value' &&
      operation.content.value.kind === 'text'
    ) {
      totalTextCharacters += operation.content.value.text.length;
      if (totalTextCharacters > readerLimits.maxTextCharacters) {
        throw new XlsxWriteError(
          'resource-limit-exceeded',
          'XLSX operations exceed their text character limit',
          {
            actual: totalTextCharacters,
            limit: readerLimits.maxTextCharacters,
            limitName: 'maxTextCharacters',
            operationId: operation.operationId,
          },
        );
      }
    }
    if (operation.kind === 'set-hyperlink' && operation.target !== null) {
      totalTextCharacters +=
        operation.target.kind === 'internal'
          ? operation.target.location.length
          : operation.target.url.length +
            (operation.target.location?.length ?? 0);
      if (totalTextCharacters > readerLimits.maxTextCharacters) {
        throw new XlsxWriteError(
          'resource-limit-exceeded',
          'XLSX operations exceed their text character limit',
          {
            actual: totalTextCharacters,
            limit: readerLimits.maxTextCharacters,
            limitName: 'maxTextCharacters',
            operationId: operation.operationId,
          },
        );
      }
    }
  }
  return operations;
}
