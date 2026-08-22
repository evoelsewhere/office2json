import { XlsxParseError } from '../errors';
import type {
  XlsxAutoFilter,
  XlsxCalendarType,
  XlsxDateGroupFilter,
  XlsxDynamicFilterType,
  XlsxFilterColumn,
  XlsxFilterRule,
  XlsxIconSet,
  XlsxRange,
  XlsxSortCondition,
  XlsxSortState,
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
  type XlsxWorksheetBudget,
} from './worksheet';

type XmlRecord = Record<string, unknown>;

const CALENDAR_TYPES = new Set([
  'gregorian',
  'gregorianArabic',
  'gregorianMeFrench',
  'gregorianUs',
  'gregorianXlitEnglish',
  'gregorianXlitFrench',
  'hebrew',
  'hijri',
  'japan',
  'korea',
  'none',
  'saka',
  'taiwan',
  'thai',
]);
const DYNAMIC_TYPES = new Set([
  'aboveAverage',
  'belowAverage',
  'lastMonth',
  'lastQuarter',
  'lastWeek',
  'lastYear',
  'M1',
  'M10',
  'M11',
  'M12',
  'M2',
  'M3',
  'M4',
  'M5',
  'M6',
  'M7',
  'M8',
  'M9',
  'nextMonth',
  'nextQuarter',
  'nextWeek',
  'nextYear',
  'null',
  'Q1',
  'Q2',
  'Q3',
  'Q4',
  'thisMonth',
  'thisQuarter',
  'thisWeek',
  'thisYear',
  'today',
  'tomorrow',
  'yearToDate',
  'yesterday',
]);
const ICON_SETS = new Set([
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
const XML_REFERENCE_PATTERN =
  /&(?:amp|apos|gt|lt|quot|#(?:x[0-9A-Fa-f]+|[0-9]+));/gu;

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

function child(node: XmlRecord, prefix: string, name: string): unknown {
  return node[prefix ? `${prefix}:${name}` : name];
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
): number {
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > maximum) {
    valueFailure(message, part);
  }
  return parsed;
}

function optionalNumber(
  value: unknown,
  part: string,
  message: string,
): number | undefined {
  if (value === undefined) return undefined;
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

function range(value: unknown, part: string, message: string): XlsxRange {
  const parsed = parseXlsxRangeReference(value);
  if (!parsed || (typeof value === 'string' && value.includes('$'))) {
    valueFailure(message, part);
  }
  return parsed;
}

function boundedRange(
  value: unknown,
  limits: ResolvedXlsxResourceLimits,
  part: string,
  message: string,
): XlsxRange {
  const parsed = range(value, part, message);
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
  return parsed;
}

function consumeRangeArea(
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): void {
  consumeXlsxWorksheetBudget(
    budget,
    'rangeAreas',
    1,
    'maxRangeAreas',
    limits,
    part,
  );
}

function containsRange(container: XlsxRange, value: XlsxRange): boolean {
  return (
    value.start.row >= container.start.row &&
    value.end.row <= container.end.row &&
    value.start.column >= container.start.column &&
    value.end.column <= container.end.column
  );
}

function iconSetCardinality(value: XlsxIconSet): number {
  return value.codePointAt(0)! - 0x30;
}

function differentialStyle(
  value: unknown,
  count: number,
  part: string,
  message: string,
): number | undefined {
  if (value === undefined) return undefined;
  const index = unsignedInteger(value, 0xffff_ffff, part, message);
  if (index >= count) valueFailure(message, part);
  return index;
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

function selectionRelation(
  selection: XlsxResolvedSheetSelection,
  value: XlsxRange,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxAutoFilter['selectionRelation'] | null {
  if (selection.kind !== 'selected-ranges') {
    return selection.kind === 'full-sheet' ? 'full-sheet' : null;
  }
  for (const selected of selection.ranges) {
    consumeXlsxWorksheetBudget(
      budget,
      'scannedCells',
      1,
      'maxScannedCells',
      limits,
      part,
    );
    if (
      selected.start.row <= value.end.row &&
      selected.end.row >= value.start.row &&
      selected.start.column <= value.end.column &&
      selected.end.column >= value.start.column
    ) {
      return 'intersects-selection';
    }
  }
  return null;
}

function dateGroup(value: XmlRecord, part: string): XlsxDateGroupFilter {
  const attrs = attributes(value);
  const grouping = attrs.dateTimeGrouping;
  if (
    grouping !== 'year' &&
    grouping !== 'month' &&
    grouping !== 'day' &&
    grouping !== 'hour' &&
    grouping !== 'minute' &&
    grouping !== 'second'
  ) {
    valueFailure('Date-group filter grouping is invalid', part);
  }
  const year = unsignedInteger(
    attrs.year,
    9999,
    part,
    'Date-group filter year is invalid',
  );
  const needsMonth = grouping !== 'year';
  const needsDay = grouping !== 'year' && grouping !== 'month';
  const needsHour =
    grouping === 'hour' || grouping === 'minute' || grouping === 'second';
  const needsMinute = grouping === 'minute' || grouping === 'second';
  const needsSecond = grouping === 'second';
  const month =
    attrs.month === undefined
      ? undefined
      : unsignedInteger(
          attrs.month,
          12,
          part,
          'Date-group filter month is invalid',
        );
  const day =
    attrs.day === undefined
      ? undefined
      : unsignedInteger(
          attrs.day,
          31,
          part,
          'Date-group filter day is invalid',
        );
  const hour =
    attrs.hour === undefined
      ? undefined
      : unsignedInteger(
          attrs.hour,
          23,
          part,
          'Date-group filter hour is invalid',
        );
  const minute =
    attrs.minute === undefined
      ? undefined
      : unsignedInteger(
          attrs.minute,
          59,
          part,
          'Date-group filter minute is invalid',
        );
  const second =
    attrs.second === undefined
      ? undefined
      : unsignedInteger(
          attrs.second,
          59,
          part,
          'Date-group filter second is invalid',
        );
  if (
    (needsMonth && (month === undefined || month === 0)) ||
    (needsDay && (day === undefined || day === 0)) ||
    (needsHour && hour === undefined) ||
    (needsMinute && minute === undefined) ||
    (needsSecond && second === undefined)
  ) {
    valueFailure('Date-group filter fields are incomplete', part);
  }
  return {
    ...(day === undefined ? {} : { day }),
    grouping,
    ...(hour === undefined ? {} : { hour }),
    ...(minute === undefined ? {} : { minute }),
    ...(month === undefined ? {} : { month }),
    ...(second === undefined ? {} : { second }),
    year,
  };
}

function valuesRule(
  node: XmlRecord,
  prefix: string,
  part: string,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
): XlsxFilterRule {
  const attrs = attributes(node);
  const filterNodes = records(child(node, prefix, 'filter'));
  const dateNodes = records(child(node, prefix, 'dateGroupItem'));
  if (!filterNodes || !dateNodes) {
    structureFailure('Value-filter collection is invalid', part);
  }
  const values = filterNodes.map((filter) => {
    const value = attributes(filter).val;
    if (typeof value !== 'string')
      valueFailure('Filter value is invalid', part);
    consumeText(value, budget, limits, part);
    return value;
  });
  const dates = dateNodes.map((date) => dateGroup(date, part));
  const blank =
    optionalBoolean(attrs.blank, part, 'Filter blank flag is invalid') ?? false;
  if (values.length === 0 && dates.length === 0 && !blank) {
    valueFailure('Value filter is empty', part);
  }
  const calendarType = attrs.calendarType;
  if (
    calendarType !== undefined &&
    (typeof calendarType !== 'string' || !CALENDAR_TYPES.has(calendarType))
  ) {
    valueFailure('Filter calendar type is invalid', part);
  }
  return {
    blank,
    ...(calendarType === undefined
      ? {}
      : { calendarType: calendarType as XlsxCalendarType }),
    dates,
    kind: 'values',
    values,
  };
}

function customRule(
  node: XmlRecord,
  prefix: string,
  part: string,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
): XlsxFilterRule {
  const conditionNodes = records(child(node, prefix, 'customFilter'));
  if (
    !conditionNodes ||
    conditionNodes.length < 1 ||
    conditionNodes.length > 2
  ) {
    structureFailure('Custom-filter condition collection is invalid', part);
  }
  const operators: Record<
    string,
    Extract<
      XlsxFilterRule,
      { kind: 'custom' }
    >['conditions'][number]['operator']
  > = {
    equal: 'equal',
    greaterThan: 'greater-than',
    greaterThanOrEqual: 'greater-than-or-equal',
    lessThan: 'less-than',
    lessThanOrEqual: 'less-than-or-equal',
    notEqual: 'not-equal',
  };
  const conditions = conditionNodes.map((condition) => {
    const attrs = attributes(condition);
    const operator =
      attrs.operator === undefined
        ? 'equal'
        : typeof attrs.operator === 'string'
          ? operators[attrs.operator]
          : undefined;
    if (!operator) valueFailure('Custom-filter operator is invalid', part);
    if (typeof attrs.val !== 'string') {
      valueFailure('Custom-filter value is invalid', part);
    }
    consumeText(attrs.val, budget, limits, part);
    return { operator, value: attrs.val };
  });
  return {
    and:
      optionalBoolean(
        attributes(node).and,
        part,
        'Custom-filter conjunction is invalid',
      ) ?? false,
    conditions,
    kind: 'custom',
  };
}

function dynamicRule(node: XmlRecord, part: string): XlsxFilterRule {
  const attrs = attributes(node);
  if (typeof attrs.type !== 'string' || !DYNAMIC_TYPES.has(attrs.type)) {
    valueFailure('Dynamic-filter type is invalid', part);
  }
  const value = optionalNumber(
    attrs.val,
    part,
    'Dynamic-filter value is invalid',
  );
  const maxValue = optionalNumber(
    attrs.maxVal,
    part,
    'Dynamic-filter maximum is invalid',
  );
  return {
    kind: 'dynamic',
    ...(maxValue === undefined ? {} : { maxValue }),
    type: attrs.type as XlsxDynamicFilterType,
    ...(value === undefined ? {} : { value }),
  };
}

function topRule(node: XmlRecord, part: string): XlsxFilterRule {
  const attrs = attributes(node);
  const value = optionalNumber(attrs.val, part, 'Top-filter value is invalid');
  if (value === undefined || value < 0) {
    valueFailure('Top-filter value is invalid', part);
  }
  const filterValue = optionalNumber(
    attrs.filterVal,
    part,
    'Top-filter threshold is invalid',
  );
  return {
    ...(filterValue === undefined ? {} : { filterValue }),
    kind: 'top',
    percent:
      optionalBoolean(
        attrs.percent,
        part,
        'Top-filter percent flag is invalid',
      ) ?? false,
    top:
      optionalBoolean(attrs.top, part, 'Top-filter direction is invalid') ??
      true,
    value,
  };
}

function colorRule(
  node: XmlRecord,
  differentialStyleCount: number,
  part: string,
): XlsxFilterRule {
  const attrs = attributes(node);
  const style = differentialStyle(
    attrs.dxfId,
    differentialStyleCount,
    part,
    'Color-filter differential-style reference is invalid',
  );
  return {
    cellColor:
      optionalBoolean(
        attrs.cellColor,
        part,
        'Color-filter cell-color flag is invalid',
      ) ?? true,
    ...(style === undefined ? {} : { differentialStyle: style }),
    kind: 'color',
  };
}

function iconRule(node: XmlRecord, part: string): XlsxFilterRule {
  const attrs = attributes(node);
  if (typeof attrs.iconSet !== 'string' || !ICON_SETS.has(attrs.iconSet)) {
    valueFailure('Icon-filter set is invalid', part);
  }
  const iconSet = attrs.iconSet as XlsxIconSet;
  return {
    iconId: unsignedInteger(
      attrs.iconId,
      iconSetCardinality(iconSet) - 1,
      part,
      'Icon-filter ID is invalid',
    ),
    iconSet,
    kind: 'icon',
  };
}

function filterColumn(
  node: XmlRecord,
  prefix: string,
  rangeWidth: number,
  differentialStyleCount: number,
  part: string,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
): XlsxFilterColumn {
  const attrs = attributes(node);
  const columnId = unsignedInteger(
    attrs.colId,
    0xffff_ffff,
    part,
    'Filter column ID is invalid',
  );
  if (columnId >= rangeWidth) valueFailure('Filter column ID is invalid', part);
  const ruleSources = [
    ['filters', child(node, prefix, 'filters')],
    ['custom', child(node, prefix, 'customFilters')],
    ['dynamic', child(node, prefix, 'dynamicFilter')],
    ['top', child(node, prefix, 'top10')],
    ['color', child(node, prefix, 'colorFilter')],
    ['icon', child(node, prefix, 'iconFilter')],
  ] as const;
  const present = ruleSources.filter(([, value]) => value !== undefined);
  if (present.length > 1) {
    structureFailure('Filter column contains multiple rule kinds', part);
  }
  let rule: XlsxFilterRule = { kind: 'none' };
  if (present.length === 1) {
    const [kind, source] = present[0]!;
    const ruleNode = record(source);
    if (!ruleNode || Array.isArray(source)) {
      structureFailure('Filter rule is invalid', part);
    }
    rule =
      kind === 'filters'
        ? valuesRule(ruleNode, prefix, part, budget, limits)
        : kind === 'custom'
          ? customRule(ruleNode, prefix, part, budget, limits)
          : kind === 'dynamic'
            ? dynamicRule(ruleNode, part)
            : kind === 'top'
              ? topRule(ruleNode, part)
              : kind === 'color'
                ? colorRule(ruleNode, differentialStyleCount, part)
                : iconRule(ruleNode, part);
  }
  return {
    columnId,
    hiddenButton:
      optionalBoolean(
        attrs.hiddenButton,
        part,
        'Filter hidden-button flag is invalid',
      ) ?? false,
    rule,
    showButton:
      optionalBoolean(
        attrs.showButton,
        part,
        'Filter show-button flag is invalid',
      ) ?? true,
  };
}

function sortCondition(
  node: XmlRecord,
  differentialStyleCount: number,
  sortRange: XlsxRange,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxSortCondition {
  const attrs = attributes(node);
  const sortBy =
    attrs.sortBy === undefined || attrs.sortBy === 'value'
      ? 'value'
      : attrs.sortBy === 'cellColor'
        ? 'cell-color'
        : attrs.sortBy === 'fontColor'
          ? 'font-color'
          : attrs.sortBy === 'icon'
            ? 'icon'
            : undefined;
  if (!sortBy) valueFailure('Sort condition kind is invalid', part);
  const sourceIconSet = attrs.iconSet;
  if (
    sourceIconSet !== undefined &&
    (typeof sourceIconSet !== 'string' || !ICON_SETS.has(sourceIconSet))
  ) {
    valueFailure('Sort condition icon set is invalid', part);
  }
  const iconSet = sourceIconSet as XlsxIconSet | undefined;
  const iconId =
    attrs.iconId === undefined
      ? undefined
      : unsignedInteger(
          attrs.iconId,
          iconSet === undefined ? 4 : iconSetCardinality(iconSet) - 1,
          part,
          'Sort condition icon ID is invalid',
        );
  if (sortBy === 'icon' && (iconSet === undefined || iconId === undefined)) {
    valueFailure('Icon sort condition metadata is missing', part);
  }
  const style = differentialStyle(
    attrs.dxfId,
    differentialStyleCount,
    part,
    'Sort condition differential-style reference is invalid',
  );
  const customList = attrs.customList;
  if (customList !== undefined && typeof customList !== 'string') {
    valueFailure('Sort condition custom list is invalid', part);
  }
  if (typeof customList === 'string') {
    consumeText(customList, budget, limits, part);
  }
  const conditionRange = boundedRange(
    attrs.ref,
    limits,
    part,
    'Sort condition range is invalid',
  );
  consumeRangeArea(budget, limits, part);
  if (!containsRange(sortRange, conditionRange)) {
    valueFailure('Sort condition range is outside the sort state', part);
  }
  return {
    ...(customList === undefined ? {} : { customList }),
    descending:
      optionalBoolean(
        attrs.descending,
        part,
        'Sort condition direction is invalid',
      ) ?? false,
    ...(style === undefined ? {} : { differentialStyle: style }),
    ...(iconId === undefined ? {} : { iconId }),
    ...(iconSet === undefined ? {} : { iconSet }),
    range: conditionRange,
    sortBy,
  };
}

function sortState(
  value: unknown,
  prefix: string,
  differentialStyleCount: number,
  filterRange: XlsxRange,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxSortState | undefined {
  if (value === undefined) return undefined;
  const node = record(value);
  if (!node || Array.isArray(value))
    structureFailure('Sort state is invalid', part);
  const conditionNodes = records(child(node, prefix, 'sortCondition'));
  if (!conditionNodes || conditionNodes.length === 0) {
    structureFailure('Sort condition collection is invalid', part);
  }
  if (conditionNodes.length > 64) {
    valueFailure('Sort condition count exceeds the SpreadsheetML bound', part);
  }
  const attrs = attributes(node);
  const stateRange = boundedRange(
    attrs.ref,
    limits,
    part,
    'Sort-state range is invalid',
  );
  consumeRangeArea(budget, limits, part);
  if (!containsRange(filterRange, stateRange)) {
    valueFailure('Sort-state range is outside the auto-filter', part);
  }
  const conditions = conditionNodes.map((condition) =>
    sortCondition(
      condition,
      differentialStyleCount,
      stateRange,
      budget,
      limits,
      part,
    ),
  );
  const method = attrs.sortMethod;
  const sortMethod =
    method === undefined || method === 'none'
      ? 'none'
      : method === 'pinYin'
        ? 'pin-yin'
        : method === 'stroke'
          ? 'stroke'
          : undefined;
  if (!sortMethod) valueFailure('Sort method is invalid', part);
  return {
    caseSensitive:
      optionalBoolean(
        attrs.caseSensitive,
        part,
        'Sort case-sensitive flag is invalid',
      ) ?? false,
    columnSort:
      optionalBoolean(
        attrs.columnSort,
        part,
        'Sort column direction flag is invalid',
      ) ?? false,
    conditions,
    range: stateRange,
    sortMethod,
  };
}

export function parseXlsxAutoFilter(
  value: unknown,
  prefix: string,
  differentialStyleCount: number,
  selection: XlsxResolvedSheetSelection,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxAutoFilter | undefined {
  if (value === undefined) return undefined;
  const node = record(value);
  if (!node || Array.isArray(value)) {
    structureFailure('Auto-filter structure is invalid', part);
  }
  const attrs = attributes(node);
  const filterRange = boundedRange(
    attrs.ref,
    limits,
    part,
    'Auto-filter range is invalid',
  );
  consumeRangeArea(budget, limits, part);
  const columnNodes = records(child(node, prefix, 'filterColumn'));
  if (!columnNodes)
    structureFailure('Filter column collection is invalid', part);
  if (columnNodes.length > limits.maxColumnsPerWorksheet) {
    valueFailure('Filter column count exceeds the worksheet bound', part);
  }
  const width = filterRange.end.column - filterRange.start.column + 1;
  const columns = columnNodes.map((column) =>
    filterColumn(
      column,
      prefix,
      width,
      differentialStyleCount,
      part,
      budget,
      limits,
    ),
  );
  if (
    new Set(columns.map((column) => column.columnId)).size !== columns.length
  ) {
    valueFailure('Auto-filter contains duplicate column IDs', part);
  }
  const sort = sortState(
    child(node, prefix, 'sortState'),
    prefix,
    differentialStyleCount,
    filterRange,
    budget,
    limits,
    part,
  );
  const relation = selectionRelation(
    selection,
    filterRange,
    budget,
    limits,
    part,
  );
  if (relation === null) return undefined;
  return {
    columns,
    range: filterRange,
    selectionRelation: relation,
    ...(sort === undefined ? {} : { sort }),
  };
}

interface CapturedFilterNode {
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

export class XlsxAutoFilterCapture implements XlsxXmlEventSink {
  private readonly stack: CapturedFilterNode[] = [];
  private root: XmlRecord | undefined;

  constructor(
    private readonly differentialStyleCount: number,
    private readonly selection: XlsxResolvedSheetSelection,
    private readonly budget: XlsxWorksheetBudget,
    private readonly limits: ResolvedXlsxResourceLimits,
    private readonly part: string,
  ) {}

  openElement(element: XlsxXmlElement): void {
    const attrs = capturedAttributes(element);
    const node: XmlRecord = { attrs };
    const parent = this.stack.at(-1)?.node;
    if (parent) {
      const current = parent[element.localName];
      if (current === undefined) parent[element.localName] = node;
      else if (Array.isArray(current)) current.push(node);
      else parent[element.localName] = [current, node];
    } else {
      if (element.localName !== 'autoFilter' || this.root !== undefined) {
        structureFailure('Auto-filter capture root is invalid', this.part);
      }
      this.root = node;
    }
    this.stack.push({ name: element.localName, node });
  }

  closeElement(element: XlsxXmlElement): void {
    if (this.stack.pop()?.name !== element.localName) {
      structureFailure('Auto-filter capture nesting is invalid', this.part);
    }
  }

  text(value: string): void {
    if (value.trim().length !== 0) {
      structureFailure('Auto-filter text content is invalid', this.part);
    }
  }

  result(): XlsxAutoFilter | undefined {
    if (!this.root || this.stack.length !== 0) {
      structureFailure('Auto-filter capture is incomplete', this.part);
    }
    return parseXlsxAutoFilter(
      this.root,
      '',
      this.differentialStyleCount,
      this.selection,
      this.budget,
      this.limits,
      this.part,
    );
  }
}
