import { XlsxParseError } from '../errors';
import type {
  XlsxDataValidation,
  XlsxDataValidationImeMode,
  XlsxDataValidationOperator,
  XlsxDataValidationSettings,
  XlsxDataValidationType,
  XlsxRange,
} from '../types';
import { parseXlsxRangeReference } from './cell-reference';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import type { XlsxResolvedSheetSelection } from './selection';
import type { XlsxXmlElement, XlsxXmlEventSink } from './streaming-xml';
import {
  consumeXlsxWorksheetBudget,
  consumeXlsxWorksheetFormulaCharacters,
  type XlsxWorksheetBudget,
} from './worksheet';

type XmlRecord = Record<string, unknown>;

const XML_REFERENCE_PATTERN =
  /&(?:amp|apos|gt|lt|quot|#(?:x[0-9A-Fa-f]+|[0-9]+));/gu;

const TYPE_MAP = new Map<string, XlsxDataValidationType>([
  ['custom', 'custom'],
  ['date', 'date'],
  ['decimal', 'decimal'],
  ['list', 'list'],
  ['none', 'none'],
  ['textLength', 'text-length'],
  ['time', 'time'],
  ['whole', 'whole'],
]);

const OPERATOR_MAP = new Map<string, XlsxDataValidationOperator>([
  ['between', 'between'],
  ['equal', 'equal'],
  ['greaterThan', 'greater-than'],
  ['greaterThanOrEqual', 'greater-than-or-equal'],
  ['lessThan', 'less-than'],
  ['lessThanOrEqual', 'less-than-or-equal'],
  ['notBetween', 'not-between'],
  ['notEqual', 'not-equal'],
]);

const IME_MODE_MAP = new Map<string, XlsxDataValidationImeMode>([
  ['disabled', 'disabled'],
  ['fullAlpha', 'full-alpha'],
  ['fullHangul', 'full-hangul'],
  ['fullKatakana', 'full-katakana'],
  ['halfAlpha', 'half-alpha'],
  ['halfHangul', 'half-hangul'],
  ['halfKatakana', 'half-katakana'],
  ['hiragana', 'hiragana'],
  ['noControl', 'no-control'],
  ['off', 'off'],
  ['on', 'on'],
]);

function structureFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-structure',
    message,
    part,
    severity: 'error',
  });
}

function valueFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-value',
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
  if (value === undefined) return [];
  const values = Array.isArray(value) ? value : [value];
  const output: XmlRecord[] = [];
  for (const item of values) {
    const parsed = record(item);
    if (!parsed) return undefined;
    output.push(parsed);
  }
  return output;
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

function optionalBoolean(
  value: unknown,
  part: string,
  message: string,
): boolean | undefined {
  if (value === undefined) return undefined;
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  valueFailure(message, part);
}

function unsignedInteger(
  value: unknown,
  part: string,
  message: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > 0xffff_ffff) {
    valueFailure(message, part);
  }
  return parsed;
}

function textAttribute(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined;
}

function consumeText(
  value: string,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): void {
  consumeXlsxWorksheetBudget(
    budget,
    'textCharacters',
    value.length,
    'maxTextCharacters',
    limits,
    part,
  );
}

function formula(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
  message: string,
): string | undefined {
  if (value === undefined) return undefined;
  if (Array.isArray(value)) structureFailure(message, part);
  const node = record(value);
  const expression =
    typeof value === 'string'
      ? decodeXmlEntities(value)
      : typeof node?.value === 'string'
        ? decodeXmlEntities(node.value)
        : undefined;
  if (expression === undefined || expression.length === 0) {
    valueFailure(message, part);
  }
  consumeXlsxWorksheetFormulaCharacters(budget, expression, limits, part);
  return expression;
}

function validationRanges(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxRange[] {
  if (typeof value !== 'string') {
    valueFailure('Data-validation range list is invalid', part);
  }
  const references = value.trim().split(/[\t\n\r ]+/u);
  if (references[0] === '') {
    valueFailure('Data-validation range list is invalid', part);
  }
  const ranges = references.map((reference) => {
    const parsed = parseXlsxRangeReference(reference);
    if (!parsed || reference.includes('$')) {
      valueFailure('Data-validation range is invalid', part);
    }
    if (parsed.end.row > limits.maxRowsPerWorksheet) {
      throw new XlsxResourceLimitError(
        'maxRowsPerWorksheet',
        parsed.end.row,
        limits.maxRowsPerWorksheet,
        part,
      );
    }
    if (parsed.end.column > limits.maxColumnsPerWorksheet) {
      throw new XlsxResourceLimitError(
        'maxColumnsPerWorksheet',
        parsed.end.column,
        limits.maxColumnsPerWorksheet,
        part,
      );
    }
    consumeXlsxWorksheetBudget(
      budget,
      'rangeAreas',
      1,
      'maxRangeAreas',
      limits,
      part,
    );
    return parsed;
  });
  if (new Set(ranges.map((range) => range.reference)).size !== ranges.length) {
    valueFailure('Data-validation range list contains duplicates', part);
  }
  return ranges;
}

function rangesIntersect(left: XlsxRange, right: XlsxRange): boolean {
  return (
    left.start.row <= right.end.row &&
    left.end.row >= right.start.row &&
    left.start.column <= right.end.column &&
    left.end.column >= right.start.column
  );
}

function selectionRelation(
  selection: XlsxResolvedSheetSelection,
  ranges: readonly XlsxRange[],
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxDataValidation['selectionRelation'] | null {
  if (selection.kind !== 'selected-ranges') {
    return selection.kind === 'full-sheet' ? 'full-sheet' : null;
  }
  for (const range of ranges) {
    for (const selected of selection.ranges) {
      consumeXlsxWorksheetBudget(
        budget,
        'scannedCells',
        1,
        'maxScannedCells',
        limits,
        part,
      );
      if (rangesIntersect(range, selected)) return 'intersects-selection';
    }
  }
  return null;
}

function validationRule(
  node: XmlRecord,
  selection: XlsxResolvedSheetSelection,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxDataValidation | null {
  consumeXlsxWorksheetBudget(
    budget,
    'validationRules',
    1,
    'maxValidationRules',
    limits,
    part,
  );
  const attrs = attributes(node);
  const sourceType = attrs.type ?? 'none';
  const type =
    typeof sourceType === 'string' ? TYPE_MAP.get(sourceType) : undefined;
  if (!type) valueFailure('Data-validation type is invalid', part);
  const sourceOperator = attrs.operator ?? 'between';
  const operator =
    typeof sourceOperator === 'string'
      ? OPERATOR_MAP.get(sourceOperator)
      : undefined;
  if (!operator) valueFailure('Data-validation operator is invalid', part);
  const sourceErrorStyle = attrs.errorStyle ?? 'stop';
  const errorStyle =
    sourceErrorStyle === 'stop' ||
    sourceErrorStyle === 'warning' ||
    sourceErrorStyle === 'information'
      ? sourceErrorStyle
      : undefined;
  if (!errorStyle) valueFailure('Data-validation error style is invalid', part);
  const sourceImeMode = attrs.imeMode ?? 'noControl';
  const imeMode =
    typeof sourceImeMode === 'string'
      ? IME_MODE_MAP.get(sourceImeMode)
      : undefined;
  if (!imeMode) valueFailure('Data-validation IME mode is invalid', part);
  const ranges = validationRanges(attrs.sqref, budget, limits, part);
  const formula1 = formula(
    node.formula1,
    budget,
    limits,
    part,
    'Data-validation first formula is invalid',
  );
  const formula2 = formula(
    node.formula2,
    budget,
    limits,
    part,
    'Data-validation second formula is invalid',
  );
  if (type !== 'none' && formula1 === undefined) {
    valueFailure('Data-validation first formula is missing', part);
  }
  if (
    type !== 'none' &&
    type !== 'list' &&
    type !== 'custom' &&
    (operator === 'between' || operator === 'not-between') &&
    formula2 === undefined
  ) {
    valueFailure('Data-validation second formula is missing', part);
  }
  const error = textAttribute(attrs.error);
  const errorTitle = textAttribute(attrs.errorTitle);
  const prompt = textAttribute(attrs.prompt);
  const promptTitle = textAttribute(attrs.promptTitle);
  for (const text of [error, errorTitle, prompt, promptTitle]) {
    if (text !== undefined) consumeText(text, budget, limits, part);
  }
  const relation = selectionRelation(selection, ranges, budget, limits, part);
  if (relation === null) return null;
  return {
    allowBlank:
      optionalBoolean(
        attrs.allowBlank,
        part,
        'Data-validation allow-blank flag is invalid',
      ) ?? false,
    ...(error === undefined ? {} : { error }),
    errorStyle,
    ...(errorTitle === undefined ? {} : { errorTitle }),
    ...(formula1 === undefined ? {} : { formula1 }),
    ...(formula2 === undefined ? {} : { formula2 }),
    imeMode,
    operator,
    ...(prompt === undefined ? {} : { prompt }),
    ...(promptTitle === undefined ? {} : { promptTitle }),
    ranges,
    selectionRelation: relation,
    showDropDown:
      optionalBoolean(
        attrs.showDropDown,
        part,
        'Data-validation drop-down flag is invalid',
      ) ?? false,
    showErrorMessage:
      optionalBoolean(
        attrs.showErrorMessage,
        part,
        'Data-validation error-message flag is invalid',
      ) ?? false,
    showInputMessage:
      optionalBoolean(
        attrs.showInputMessage,
        part,
        'Data-validation input-message flag is invalid',
      ) ?? false,
    type,
  };
}

export interface XlsxParsedDataValidations {
  rules: XlsxDataValidation[];
  settings: XlsxDataValidationSettings;
}

export function parseXlsxDataValidations(
  value: unknown,
  selection: XlsxResolvedSheetSelection,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxParsedDataValidations {
  const node = record(value);
  if (!node || Array.isArray(value)) {
    structureFailure('Data-validations collection is invalid', part);
  }
  const attrs = attributes(node);
  const validationNodes = records(node.dataValidation);
  if (!validationNodes) {
    structureFailure('Data-validation rule collection is invalid', part);
  }
  const declaredCount = unsignedInteger(
    attrs.count,
    part,
    'Data-validation count is invalid',
  );
  if (declaredCount !== undefined && declaredCount !== validationNodes.length) {
    structureFailure('Data-validation count does not match', part);
  }
  const rules = validationNodes.flatMap((validation) => {
    const parsed = validationRule(validation, selection, budget, limits, part);
    return parsed ? [parsed] : [];
  });
  const xWindow = unsignedInteger(
    attrs.xWindow,
    part,
    'Data-validation prompt X position is invalid',
  );
  const yWindow = unsignedInteger(
    attrs.yWindow,
    part,
    'Data-validation prompt Y position is invalid',
  );
  return {
    rules,
    settings: {
      disablePrompts:
        optionalBoolean(
          attrs.disablePrompts,
          part,
          'Data-validation disable-prompts flag is invalid',
        ) ?? false,
      ...(xWindow === undefined ? {} : { xWindow }),
      ...(yWindow === undefined ? {} : { yWindow }),
    },
  };
}

interface CapturedNode {
  name: string;
  node: XmlRecord;
}

function capturedAttributes(element: XlsxXmlElement): XmlRecord {
  return Object.fromEntries(
    [...element.attributes].map(([name, value]) => [
      name.startsWith('{}') ? name.slice(2) : name,
      value,
    ]),
  );
}

export class XlsxDataValidationsCapture implements XlsxXmlEventSink {
  private readonly stack: CapturedNode[] = [];
  private root: XmlRecord | undefined;

  constructor(
    private readonly selection: XlsxResolvedSheetSelection,
    private readonly budget: XlsxWorksheetBudget,
    private readonly limits: ResolvedXlsxResourceLimits,
    private readonly part: string,
  ) {}

  openElement(element: XlsxXmlElement): void {
    const node: XmlRecord = { attrs: capturedAttributes(element) };
    const parent = this.stack.at(-1)?.node;
    if (parent) {
      const current = parent[element.localName];
      if (current === undefined) parent[element.localName] = node;
      else if (Array.isArray(current)) current.push(node);
      else parent[element.localName] = [current, node];
    } else {
      if (element.localName !== 'dataValidations' || this.root !== undefined) {
        structureFailure('Data-validations capture root is invalid', this.part);
      }
      this.root = node;
    }
    this.stack.push({ name: element.localName, node });
  }

  closeElement(element: XlsxXmlElement): void {
    if (this.stack.pop()?.name !== element.localName) {
      structureFailure(
        'Data-validations capture nesting is invalid',
        this.part,
      );
    }
  }

  text(value: string): void {
    const current = this.stack.at(-1);
    if (current?.name === 'formula1' || current?.name === 'formula2') {
      const existing =
        typeof current.node.value === 'string' ? current.node.value : '';
      current.node.value = `${existing}${value}`;
      return;
    }
    if (value.trim().length !== 0) {
      structureFailure('Data-validations text content is invalid', this.part);
    }
  }

  result(): XlsxParsedDataValidations {
    if (!this.root || this.stack.length !== 0) {
      structureFailure('Data-validations capture is incomplete', this.part);
    }
    return parseXlsxDataValidations(
      this.root,
      this.selection,
      this.budget,
      this.limits,
      this.part,
    );
  }
}
