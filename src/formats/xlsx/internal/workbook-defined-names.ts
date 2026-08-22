import { XlsxParseError } from '../errors';
import type { XlsxDefinedName } from '../types';
import { parseXlsxCellReference } from './cell-reference';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';

type XmlRecord = Record<string, unknown>;

export interface ParsedXlsxDefinedNames {
  definedNames: XlsxDefinedName[];
  formulaCharacters: number;
  textCharacters: number;
}

const METADATA_ATTRIBUTES = [
  ['comment', 'comment'],
  ['customMenu', 'customMenu'],
  ['description', 'description'],
  ['help', 'help'],
  ['shortcutKey', 'shortcutKey'],
  ['statusBar', 'statusBar'],
] as const;

const BOOLEAN_ATTRIBUTES = [
  ['function', 'function'],
  ['publishToServer', 'publishToServer'],
  ['vbProcedure', 'vbProcedure'],
  ['workbookParameter', 'workbookParameter'],
  ['xlm', 'xlm'],
] as const;

const XML_REFERENCE_PATTERN =
  /&(?:amp|apos|gt|lt|quot|#(?:x[0-9A-Fa-f]+|[0-9]+));/gu;

function fail(part: string, message: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function structureFailure(part: string, message: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-structure',
    message,
    part,
    severity: 'error',
  });
}

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function records(value: unknown): XmlRecord[] | undefined {
  const values = Array.isArray(value) ? value : [value];
  const output: XmlRecord[] = [];
  for (const item of values) {
    const parsed = record(item);
    if (!parsed) return undefined;
    output.push(parsed);
  }
  return output;
}

function attributes(value: XmlRecord): XmlRecord {
  const source = record(value.attrs) ?? {};
  return Object.fromEntries(
    Object.entries(source).map(([name, attributeValue]) => [
      name,
      typeof attributeValue === 'string'
        ? decodeXmlEntities(attributeValue)
        : attributeValue,
    ]),
  );
}

function decodeXmlEntities(value: string): string {
  return value.replace(XML_REFERENCE_PATTERN, (reference) => {
    if (reference === '&amp;') return '&';
    if (reference === '&apos;') return "'";
    if (reference === '&gt;') return '>';
    if (reference === '&lt;') return '<';
    if (reference === '&quot;') return '"';
    const hexadecimal = reference[2] === 'x';
    const digits = reference.slice(hexadecimal ? 3 : 2, -1);
    return String.fromCodePoint(Number.parseInt(digits, hexadecimal ? 16 : 10));
  });
}

function booleanAttribute(
  value: unknown,
  part: string,
  message: string,
): boolean | undefined {
  if (value === undefined) return undefined;
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  fail(part, message);
}

function unsignedInteger(
  value: unknown,
  part: string,
  message: string,
): number | undefined {
  if (value === undefined) return undefined;
  const parsed = Number(value);
  if (
    String(parsed) !== value ||
    !Number.isSafeInteger(parsed) ||
    parsed < 0 ||
    parsed > 0xffff_ffff
  ) {
    fail(part, message);
  }
  return parsed;
}

export function isValidXlsxDefinedName(value: unknown): value is string {
  if (typeof value !== 'string' || value.length > 255) {
    return false;
  }
  if (!/^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(value)) return false;
  if (parseXlsxCellReference(value)) return false;
  const folded = value.toUpperCase();
  if (folded === 'R' || folded === 'C') return false;
  return !/^R[1-9]\d*C[1-9]\d*$/u.test(folded);
}

function textValue(value: XmlRecord): string | undefined {
  return typeof value.value === 'string'
    ? decodeXmlEntities(value.value)
    : undefined;
}

function consume(
  actual: number,
  amount: number,
  limitName: 'maxTextCharacters' | 'maxTotalFormulaCharacters',
  limits: ResolvedXlsxResourceLimits,
  part: string,
): number {
  const next = actual + amount;
  if (!Number.isSafeInteger(next) || next > limits[limitName]) {
    throw new XlsxResourceLimitError(limitName, next, limits[limitName], part);
  }
  return next;
}

export function parseXlsxDefinedNames(
  value: unknown,
  prefix: string,
  part: string,
  sheetCount: number,
  limits: ResolvedXlsxResourceLimits,
): ParsedXlsxDefinedNames {
  if (value === undefined) {
    return { definedNames: [], formulaCharacters: 0, textCharacters: 0 };
  }
  const container = record(value);
  if (!container) {
    structureFailure(part, 'Workbook defined-names collection is invalid');
  }
  const nodes = records(
    container[prefix ? `${prefix}:definedName` : 'definedName'],
  );
  if (!nodes || nodes.length === 0) {
    structureFailure(part, 'Workbook defined-names collection is empty');
  }
  if (nodes.length > limits.maxDefinedNames) {
    throw new XlsxResourceLimitError(
      'maxDefinedNames',
      nodes.length,
      limits.maxDefinedNames,
      part,
    );
  }

  const scopes = new Map<number | undefined, Set<string>>();
  const definedNames: XlsxDefinedName[] = [];
  let formulaCharacters = 0;
  let textCharacters = 0;
  for (const node of nodes) {
    const attrs = attributes(node);
    if (!isValidXlsxDefinedName(attrs.name)) {
      fail(part, 'Workbook defined name is invalid');
    }
    const expression = textValue(node);
    if (expression === undefined || expression.length === 0) {
      fail(part, 'Workbook defined-name expression is missing');
    }
    if (expression.length > limits.maxFormulaCharacters) {
      throw new XlsxResourceLimitError(
        'maxFormulaCharacters',
        expression.length,
        limits.maxFormulaCharacters,
        part,
      );
    }
    formulaCharacters = consume(
      formulaCharacters,
      expression.length,
      'maxTotalFormulaCharacters',
      limits,
      part,
    );
    const sheetIndex = unsignedInteger(
      attrs.localSheetId,
      part,
      'Workbook defined-name sheet scope is invalid',
    );
    if (sheetIndex !== undefined && sheetIndex >= sheetCount) {
      fail(part, 'Workbook defined-name sheet scope is invalid');
    }
    const foldedName = attrs.name.toUpperCase();
    const scope = scopes.get(sheetIndex);
    if (scope?.has(foldedName)) {
      fail(part, 'Workbook contains duplicate defined names in one scope');
    }
    if (scope) scope.add(foldedName);
    else scopes.set(sheetIndex, new Set([foldedName]));

    const metadata: Partial<XlsxDefinedName> = {};
    for (const [attributeName, propertyName] of METADATA_ATTRIBUTES) {
      const metadataValue = attrs[attributeName];
      if (metadataValue === undefined) continue;
      if (typeof metadataValue !== 'string') {
        fail(part, 'Workbook defined-name metadata is invalid');
      }
      textCharacters = consume(
        textCharacters,
        metadataValue.length,
        'maxTextCharacters',
        limits,
        part,
      );
      metadata[propertyName] = metadataValue;
    }
    for (const [attributeName, propertyName] of BOOLEAN_ATTRIBUTES) {
      const flag = booleanAttribute(
        attrs[attributeName],
        part,
        'Workbook defined-name flag is invalid',
      );
      if (flag !== undefined) metadata[propertyName] = flag;
    }
    const functionGroupId = unsignedInteger(
      attrs.functionGroupId,
      part,
      'Workbook defined-name function group is invalid',
    );
    const hidden =
      booleanAttribute(
        attrs.hidden,
        part,
        'Workbook defined-name hidden flag is invalid',
      ) ?? false;
    textCharacters = consume(
      textCharacters,
      attrs.name.length,
      'maxTextCharacters',
      limits,
      part,
    );
    definedNames.push({
      ...metadata,
      expression,
      ...(functionGroupId === undefined ? {} : { functionGroupId }),
      hidden,
      name: attrs.name,
      ...(sheetIndex === undefined ? {} : { sheetIndex }),
    });
  }
  return { definedNames, formulaCharacters, textCharacters };
}

export function xlsxDefinedNameFormulaCharacters(
  values: readonly XlsxDefinedName[],
): number {
  return values.reduce((total, value) => total + value.expression.length, 0);
}

export function xlsxDefinedNameTextCharacters(
  values: readonly XlsxDefinedName[],
): number {
  return values.reduce(
    (total, value) =>
      total +
      value.name.length +
      METADATA_ATTRIBUTES.reduce(
        (metadataTotal, [, propertyName]) =>
          metadataTotal + (value[propertyName]?.length ?? 0),
        0,
      ),
    0,
  );
}
