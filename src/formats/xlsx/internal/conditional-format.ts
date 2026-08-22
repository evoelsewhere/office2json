import { XlsxParseError } from '../errors';
import type {
  XlsxColor,
  XlsxConditionalDataBar,
  XlsxConditionalFormatting,
  XlsxConditionalFormattingOperator,
  XlsxConditionalFormattingRule,
  XlsxConditionalFormattingRuleType,
  XlsxConditionalIconSet,
  XlsxConditionalValueObject,
  XlsxIconSet,
  XlsxRange,
} from '../types';
import { parseXlsxRangeReference } from './cell-reference';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import type { XlsxResolvedSheetSelection } from './selection';
import { parseXlsxStyleColor } from './style-color';
import type { XlsxXmlElement, XlsxXmlEventSink } from './streaming-xml';
import {
  consumeXlsxWorksheetBudget,
  consumeXlsxWorksheetFormulaCharacters,
  type XlsxWorksheetBudget,
} from './worksheet';

type XmlRecord = Record<string, unknown>;

const XML_REFERENCE_PATTERN =
  /&(?:amp|apos|gt|lt|quot|#(?:x[0-9A-Fa-f]+|[0-9]+));/gu;

const RULE_TYPE_MAP = new Map<string, XlsxConditionalFormattingRuleType>([
  ['aboveAverage', 'above-average'],
  ['beginsWith', 'begins-with'],
  ['cellIs', 'cell-is'],
  ['colorScale', 'color-scale'],
  ['containsBlanks', 'contains-blanks'],
  ['containsErrors', 'contains-errors'],
  ['containsText', 'contains-text'],
  ['dataBar', 'data-bar'],
  ['duplicateValues', 'duplicate-values'],
  ['endsWith', 'ends-with'],
  ['expression', 'expression'],
  ['iconSet', 'icon-set'],
  ['notContainsBlanks', 'not-contains-blanks'],
  ['notContainsErrors', 'not-contains-errors'],
  ['notContainsText', 'not-contains-text'],
  ['timePeriod', 'time-period'],
  ['top10', 'top'],
  ['uniqueValues', 'unique-values'],
]);

const OPERATOR_MAP = new Map<string, XlsxConditionalFormattingOperator>([
  ['between', 'between'],
  ['equal', 'equal'],
  ['greaterThan', 'greater-than'],
  ['greaterThanOrEqual', 'greater-than-or-equal'],
  ['lessThan', 'less-than'],
  ['lessThanOrEqual', 'less-than-or-equal'],
  ['notBetween', 'not-between'],
  ['notEqual', 'not-equal'],
]);

const TIME_PERIOD_MAP = new Map<
  string,
  NonNullable<XlsxConditionalFormattingRule['timePeriod']>
>([
  ['last7Days', 'last-7-days'],
  ['lastMonth', 'last-month'],
  ['lastWeek', 'last-week'],
  ['nextMonth', 'next-month'],
  ['nextWeek', 'next-week'],
  ['thisMonth', 'this-month'],
  ['thisWeek', 'this-week'],
  ['today', 'today'],
  ['tomorrow', 'tomorrow'],
  ['yesterday', 'yesterday'],
]);

const ICON_SETS = new Set<XlsxIconSet>([
  '3Arrows',
  '3ArrowsGray',
  '3Flags',
  '3Signs',
  '3Symbols',
  '3Symbols2',
  '3TrafficLights1',
  '3TrafficLights2',
  '4Arrows',
  '4ArrowsGray',
  '4Rating',
  '4RedToBlack',
  '4TrafficLights',
  '5Arrows',
  '5ArrowsGray',
  '5Quarters',
  '5Rating',
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
  maximum: number,
  part: string,
  message: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > maximum) {
    valueFailure(message, part);
  }
  return parsed;
}

function requiredPositiveInteger(
  value: unknown,
  part: string,
  message: string,
): number {
  const parsed = unsignedInteger(value, 0xffff_ffff, part, message);
  if (parsed === undefined || parsed === 0) valueFailure(message, part);
  return parsed;
}

function finiteNumber(value: unknown, part: string, message: string): number {
  if (
    typeof value !== 'string' ||
    !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/u.test(value)
  ) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) valueFailure(message, part);
  return parsed === 0 ? 0 : parsed;
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

function formulaText(
  value: XmlRecord,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
  message: string,
): string {
  const expression =
    typeof value.value === 'string'
      ? decodeXmlEntities(value.value)
      : undefined;
  if (expression === undefined || expression.length === 0) {
    valueFailure(message, part);
  }
  consumeXlsxWorksheetFormulaCharacters(budget, expression, limits, part);
  return expression;
}

function conditionalRanges(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxRange[] {
  if (typeof value !== 'string') {
    valueFailure('Conditional-format range list is invalid', part);
  }
  const references = value.trim().split(/[\t\n\r ]+/u);
  if (references[0] === '') {
    valueFailure('Conditional-format range list is invalid', part);
  }
  const ranges = references.map((reference) => {
    const parsed = parseXlsxRangeReference(reference);
    if (!parsed || reference.includes('$')) {
      valueFailure('Conditional-format range is invalid', part);
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
    valueFailure('Conditional-format range list contains duplicates', part);
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
): XlsxConditionalFormatting['selectionRelation'] | null {
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

function threshold(
  node: XmlRecord,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxConditionalValueObject {
  const attrs = attributes(node);
  const greaterThanOrEqual =
    optionalBoolean(
      attrs.gte,
      part,
      'Conditional-format threshold inclusive flag is invalid',
    ) ?? true;
  const type = attrs.type;
  if (type === 'min' || type === 'max') {
    return {
      greaterThanOrEqual,
      kind: type === 'min' ? 'minimum' : 'maximum',
    };
  }
  if (type === 'formula') {
    if (typeof attrs.val !== 'string' || attrs.val.length === 0) {
      valueFailure('Conditional-format threshold formula is invalid', part);
    }
    consumeXlsxWorksheetFormulaCharacters(budget, attrs.val, limits, part);
    return {
      expression: attrs.val,
      greaterThanOrEqual,
      kind: 'formula',
    };
  }
  if (type === 'num' || type === 'percent' || type === 'percentile') {
    const value = finiteNumber(
      attrs.val,
      part,
      'Conditional-format threshold value is invalid',
    );
    if (
      (type === 'percent' || type === 'percentile') &&
      (value < 0 || value > 100)
    ) {
      valueFailure('Conditional-format threshold value is invalid', part);
    }
    return {
      greaterThanOrEqual,
      kind:
        type === 'num'
          ? 'number'
          : type === 'percent'
            ? 'percent'
            : 'percentile',
      value,
    };
  }
  valueFailure('Conditional-format threshold type is invalid', part);
}

function colors(value: unknown, part: string, context: string): XlsxColor[] {
  const nodes = records(value);
  if (!nodes) structureFailure(`${context} color collection is invalid`, part);
  return nodes.map((node) => {
    const color = parseXlsxStyleColor(node, part, context);
    if (!color) valueFailure(`${context} color is missing`, part);
    return color;
  });
}

function thresholds(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
  context: string,
): XlsxConditionalValueObject[] {
  const nodes = records(value);
  if (!nodes) {
    structureFailure(`${context} threshold collection is invalid`, part);
  }
  return nodes.map((node) => threshold(node, budget, limits, part));
}

function colorScale(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): NonNullable<XlsxConditionalFormattingRule['colorScale']> {
  const node = record(value);
  if (!node || Array.isArray(value)) {
    structureFailure('Conditional color scale is invalid', part);
  }
  const stops = thresholds(
    node.cfvo,
    budget,
    limits,
    part,
    'Conditional color scale',
  );
  const stopColors = colors(node.color, part, 'Conditional color scale');
  if (
    (stops.length !== 2 && stops.length !== 3) ||
    stopColors.length !== stops.length
  ) {
    structureFailure('Conditional color scale stop count is invalid', part);
  }
  return {
    stops: stops.map((threshold, index) => ({
      color: stopColors[index]!,
      threshold,
    })),
  };
}

function dataBar(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxConditionalDataBar {
  const node = record(value);
  if (!node || Array.isArray(value)) {
    structureFailure('Conditional data bar is invalid', part);
  }
  const barThresholds = thresholds(
    node.cfvo,
    budget,
    limits,
    part,
    'Conditional data bar',
  );
  if (barThresholds.length !== 2) {
    structureFailure('Conditional data-bar threshold count is invalid', part);
  }
  const barColors = colors(node.color, part, 'Conditional data bar');
  if (barColors.length !== 1) {
    structureFailure('Conditional data-bar color count is invalid', part);
  }
  const attrs = attributes(node);
  const minimumLength =
    unsignedInteger(
      attrs.minLength,
      100,
      part,
      'Conditional data-bar minimum length is invalid',
    ) ?? 10;
  const maximumLength =
    unsignedInteger(
      attrs.maxLength,
      100,
      part,
      'Conditional data-bar maximum length is invalid',
    ) ?? 90;
  if (minimumLength > maximumLength) {
    valueFailure('Conditional data-bar lengths are inconsistent', part);
  }
  return {
    color: barColors[0]!,
    maximumLength,
    minimumLength,
    showValue:
      optionalBoolean(
        attrs.showValue,
        part,
        'Conditional data-bar show-value flag is invalid',
      ) ?? true,
    thresholds: [barThresholds[0]!, barThresholds[1]!],
  };
}

function iconSetCardinality(value: XlsxIconSet): number {
  return value.codePointAt(0)! - 0x30;
}

function iconSet(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxConditionalIconSet {
  const node = record(value);
  if (!node || Array.isArray(value)) {
    structureFailure('Conditional icon set is invalid', part);
  }
  const attrs = attributes(node);
  const sourceIconSet = attrs.iconSet ?? '3TrafficLights1';
  if (!ICON_SETS.has(sourceIconSet as XlsxIconSet)) {
    valueFailure('Conditional icon set kind is invalid', part);
  }
  const normalizedIconSet = sourceIconSet as XlsxIconSet;
  const iconThresholds = thresholds(
    node.cfvo,
    budget,
    limits,
    part,
    'Conditional icon set',
  );
  if (iconThresholds.length !== iconSetCardinality(normalizedIconSet)) {
    structureFailure('Conditional icon-set threshold count is invalid', part);
  }
  return {
    iconSet: normalizedIconSet,
    percent:
      optionalBoolean(
        attrs.percent,
        part,
        'Conditional icon-set percent flag is invalid',
      ) ?? true,
    reverse:
      optionalBoolean(
        attrs.reverse,
        part,
        'Conditional icon-set reverse flag is invalid',
      ) ?? false,
    showValue:
      optionalBoolean(
        attrs.showValue,
        part,
        'Conditional icon-set show-value flag is invalid',
      ) ?? true,
    thresholds: iconThresholds,
  };
}

function differentialStyle(
  value: unknown,
  count: number,
  part: string,
): number | undefined {
  const index = unsignedInteger(
    value,
    0xffff_ffff,
    part,
    'Conditional-format differential-style reference is invalid',
  );
  if (index !== undefined && index >= count) {
    valueFailure(
      'Conditional-format differential-style reference is invalid',
      part,
    );
  }
  return index;
}

function rule(
  node: XmlRecord,
  differentialStyleCount: number,
  priorities: Set<number>,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxConditionalFormattingRule {
  consumeXlsxWorksheetBudget(
    budget,
    'conditionalFormattingRules',
    1,
    'maxConditionalFormattingRules',
    limits,
    part,
  );
  const attrs = attributes(node);
  const type =
    typeof attrs.type === 'string' ? RULE_TYPE_MAP.get(attrs.type) : undefined;
  if (!type) valueFailure('Conditional-format rule type is invalid', part);
  const priority = requiredPositiveInteger(
    attrs.priority,
    part,
    'Conditional-format priority is invalid',
  );
  if (priorities.has(priority)) {
    valueFailure(
      'Worksheet contains duplicate conditional-format priorities',
      part,
    );
  }
  priorities.add(priority);
  const formulaNodes = records(node.formula);
  if (!formulaNodes || formulaNodes.length > 3) {
    structureFailure('Conditional-format formula collection is invalid', part);
  }
  const formulas = formulaNodes.map((formula) =>
    formulaText(
      formula,
      budget,
      limits,
      part,
      'Conditional-format formula is invalid',
    ),
  );
  const sourceOperator = attrs.operator;
  const operator =
    typeof sourceOperator === 'string'
      ? OPERATOR_MAP.get(sourceOperator)
      : undefined;
  if (sourceOperator !== undefined && !operator) {
    valueFailure('Conditional-format operator is invalid', part);
  }
  if (type === 'cell-is') {
    if (!operator) valueFailure('Cell conditional operator is missing', part);
    const expected =
      operator === 'between' || operator === 'not-between' ? 2 : 1;
    if (formulas.length !== expected) {
      valueFailure('Cell conditional formula count is invalid', part);
    }
  }
  if (type === 'expression' && formulas.length === 0) {
    valueFailure('Expression conditional formula is missing', part);
  }
  const text = typeof attrs.text === 'string' ? attrs.text : undefined;
  if (
    type === 'begins-with' ||
    type === 'contains-text' ||
    type === 'ends-with' ||
    type === 'not-contains-text'
  ) {
    if (text === undefined || text.length === 0) {
      valueFailure('Text conditional comparison text is missing', part);
    }
  }
  if (text !== undefined) consumeText(text, budget, limits, part);
  const timePeriod =
    typeof attrs.timePeriod === 'string'
      ? TIME_PERIOD_MAP.get(attrs.timePeriod)
      : undefined;
  if (type === 'time-period' && !timePeriod) {
    valueFailure('Time-period conditional value is invalid', part);
  }
  const rank = unsignedInteger(
    attrs.rank,
    0xffff_ffff,
    part,
    'Top conditional rank is invalid',
  );
  if (type === 'top' && (rank === undefined || rank === 0)) {
    valueFailure('Top conditional rank is invalid', part);
  }
  const standardDeviations = unsignedInteger(
    attrs.stdDev,
    0xffff_ffff,
    part,
    'Average conditional standard deviations are invalid',
  );
  const visualSources = [node.colorScale, node.dataBar, node.iconSet].filter(
    (source) => source !== undefined,
  );
  if (visualSources.length > 1) {
    structureFailure(
      'Conditional-format rule has multiple visual definitions',
      part,
    );
  }
  const parsedColorScale =
    type === 'color-scale'
      ? colorScale(node.colorScale, budget, limits, part)
      : undefined;
  const parsedDataBar =
    type === 'data-bar'
      ? dataBar(node.dataBar, budget, limits, part)
      : undefined;
  const parsedIconSet =
    type === 'icon-set'
      ? iconSet(node.iconSet, budget, limits, part)
      : undefined;
  if (
    (type !== 'color-scale' && node.colorScale !== undefined) ||
    (type !== 'data-bar' && node.dataBar !== undefined) ||
    (type !== 'icon-set' && node.iconSet !== undefined)
  ) {
    structureFailure(
      'Conditional-format visual definition mismatches its rule',
      part,
    );
  }
  const style = differentialStyle(attrs.dxfId, differentialStyleCount, part);
  return {
    ...(type === 'above-average'
      ? {
          aboveAverage:
            optionalBoolean(
              attrs.aboveAverage,
              part,
              'Average conditional direction flag is invalid',
            ) ?? true,
          equalAverage:
            optionalBoolean(
              attrs.equalAverage,
              part,
              'Average conditional equality flag is invalid',
            ) ?? false,
        }
      : {}),
    ...(type === 'top'
      ? {
          bottom:
            optionalBoolean(
              attrs.bottom,
              part,
              'Top conditional bottom flag is invalid',
            ) ?? false,
        }
      : {}),
    ...(parsedColorScale === undefined ? {} : { colorScale: parsedColorScale }),
    ...(parsedDataBar === undefined ? {} : { dataBar: parsedDataBar }),
    ...(style === undefined ? {} : { differentialStyle: style }),
    formulas,
    ...(parsedIconSet === undefined ? {} : { iconSet: parsedIconSet }),
    ...(operator === undefined ? {} : { operator }),
    ...(type === 'top'
      ? {
          percent:
            optionalBoolean(
              attrs.percent,
              part,
              'Top conditional percent flag is invalid',
            ) ?? false,
        }
      : {}),
    priority,
    ...(rank === undefined ? {} : { rank }),
    ...(standardDeviations === undefined ? {} : { standardDeviations }),
    stopIfTrue:
      optionalBoolean(
        attrs.stopIfTrue,
        part,
        'Conditional-format stop flag is invalid',
      ) ?? false,
    ...(text === undefined ? {} : { text }),
    ...(timePeriod === undefined ? {} : { timePeriod }),
    type,
  };
}

export function parseXlsxConditionalFormatting(
  value: unknown,
  differentialStyleCount: number,
  priorities: Set<number>,
  selection: XlsxResolvedSheetSelection,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxConditionalFormatting | null {
  const node = record(value);
  if (!node || Array.isArray(value)) {
    structureFailure('Conditional-format collection is invalid', part);
  }
  const attrs = attributes(node);
  const ranges = conditionalRanges(attrs.sqref, budget, limits, part);
  const ruleNodes = records(node.cfRule);
  if (!ruleNodes || ruleNodes.length === 0) {
    structureFailure('Conditional-format rule collection is invalid', part);
  }
  const rules = ruleNodes.map((ruleNode) =>
    rule(ruleNode, differentialStyleCount, priorities, budget, limits, part),
  );
  const relation = selectionRelation(selection, ranges, budget, limits, part);
  if (relation === null) return null;
  return {
    pivot:
      optionalBoolean(
        attrs.pivot,
        part,
        'Conditional-format pivot flag is invalid',
      ) ?? false,
    ranges,
    rules,
    selectionRelation: relation,
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

export class XlsxConditionalFormattingCapture implements XlsxXmlEventSink {
  private readonly stack: CapturedNode[] = [];
  private root: XmlRecord | undefined;

  constructor(
    private readonly differentialStyleCount: number,
    private readonly priorities: Set<number>,
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
      if (
        element.localName !== 'conditionalFormatting' ||
        this.root !== undefined
      ) {
        structureFailure(
          'Conditional-format capture root is invalid',
          this.part,
        );
      }
      this.root = node;
    }
    this.stack.push({ name: element.localName, node });
  }

  closeElement(element: XlsxXmlElement): void {
    if (this.stack.pop()?.name !== element.localName) {
      structureFailure(
        'Conditional-format capture nesting is invalid',
        this.part,
      );
    }
  }

  text(value: string): void {
    const current = this.stack.at(-1);
    if (current?.name === 'formula') {
      const existing =
        typeof current.node.value === 'string' ? current.node.value : '';
      current.node.value = `${existing}${value}`;
      return;
    }
    if (value.trim().length !== 0) {
      structureFailure('Conditional-format text content is invalid', this.part);
    }
  }

  result(): XlsxConditionalFormatting | null {
    if (!this.root || this.stack.length !== 0) {
      structureFailure('Conditional-format capture is incomplete', this.part);
    }
    return parseXlsxConditionalFormatting(
      this.root,
      this.differentialStyleCount,
      this.priorities,
      this.selection,
      this.budget,
      this.limits,
      this.part,
    );
  }
}
