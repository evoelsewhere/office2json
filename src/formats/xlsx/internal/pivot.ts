import type { XmlLookupValue } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type {
  XlsxPivotCache,
  XlsxPivotCacheField,
  XlsxPivotCacheItem,
  XlsxPivotCacheMode,
  XlsxPivotCacheRecordValue,
  XlsxPivotCacheSource,
  XlsxPivotDataField,
  XlsxPivotField,
  XlsxPivotFieldItem,
  XlsxPivotFilter,
  XlsxPivotPageField,
  XlsxPivotTable,
} from '../types';
import { parseXlsxRangeReference } from './cell-reference';
import { getXlsxRelationshipPartName } from './package-identity';
import { XlsxPartReader } from './part-reader';
import { parseXlsxRelationships, type XlsxRelationship } from './relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import type { XlsxXmlElement, XlsxXmlEventSink } from './streaming-xml';
import type { XlsxResolvedSheetSelection } from './selection';
import type { XlsxPivotCacheDeclaration } from './workbook-manifest';
import type { XlsxWorkbookDiscovery } from './workbook-discovery';
import {
  consumeXlsxWorksheetBudget,
  type XlsxWorksheetBudget,
} from './worksheet';

type XmlRecord = Record<string, unknown>;

function relationshipBase(dialect: XlsxWorkbookDiscovery['dialect']): string {
  return dialect === 'strict'
    ? 'http://purl.oclc.org/ooxml/officeDocument/relationships'
    : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
}

export function parseXlsxPivotDataDisplayMode(
  value: string | undefined,
  part: string,
): XlsxPivotDataField['showDataAs'] {
  const modes = {
    difference: 'difference',
    index: 'index',
    normal: 'normal',
    percent: 'percent',
    percentDiff: 'percentDifference',
    percentOfCol: 'percentOfColumn',
    percentOfRow: 'percentOfRow',
    percentOfTotal: 'percentOfTotal',
    runTotal: 'runningTotal',
  } as const satisfies Record<string, XlsxPivotDataField['showDataAs']>;
  const parsed = modes[(value ?? 'normal') as keyof typeof modes];
  if (parsed === undefined) {
    fail('invalid-document-value', 'Pivot display mode is invalid', part);
  }
  return parsed;
}

export function parseXlsxPivotSubtotal(
  value: string | undefined,
  part: string,
): XlsxPivotDataField['subtotal'] {
  const subtotals = {
    average: 'average',
    count: 'count',
    countNums: 'countNumbers',
    max: 'maximum',
    min: 'minimum',
    product: 'product',
    stdDev: 'standardDeviation',
    stdDevp: 'standardDeviationPopulation',
    sum: 'sum',
    var: 'variance',
    varp: 'variancePopulation',
  } as const satisfies Record<string, XlsxPivotDataField['subtotal']>;
  const parsed = subtotals[(value ?? 'sum') as keyof typeof subtotals];
  if (parsed === undefined) {
    fail('invalid-document-value', 'Pivot subtotal is invalid', part);
  }
  return parsed;
}

export function parseXlsxPivotFilterType(
  value: string | undefined,
  part: string,
): XlsxPivotFilter['type'] {
  const types = new Set<XlsxPivotFilter['type']>([
    'captionBeginsWith',
    'captionBetween',
    'captionContains',
    'captionEndsWith',
    'captionEqual',
    'captionGreaterThan',
    'captionGreaterThanOrEqual',
    'captionLessThan',
    'captionLessThanOrEqual',
    'captionNotBeginsWith',
    'captionNotBetween',
    'captionNotContains',
    'captionNotEndsWith',
    'captionNotEqual',
    'count',
    'dateBetween',
    'dateEqual',
    'dateNewerThan',
    'dateNewerThanOrEqual',
    'dateNotBetween',
    'dateNotEqual',
    'dateOlderThan',
    'dateOlderThanOrEqual',
    'lastMonth',
    'lastQuarter',
    'lastWeek',
    'lastYear',
    'month1',
    'month10',
    'month11',
    'month12',
    'month2',
    'month3',
    'month4',
    'month5',
    'month6',
    'month7',
    'month8',
    'month9',
    'nextMonth',
    'nextQuarter',
    'nextWeek',
    'nextYear',
    'percent',
    'quarter1',
    'quarter2',
    'quarter3',
    'quarter4',
    'sum',
    'thisMonth',
    'thisQuarter',
    'thisWeek',
    'thisYear',
    'today',
    'tomorrow',
    'unknown',
    'valueBetween',
    'valueEqual',
    'valueGreaterThan',
    'valueGreaterThanOrEqual',
    'valueLessThan',
    'valueLessThanOrEqual',
    'valueNotBetween',
    'valueNotEqual',
    'yearToDate',
    'yesterday',
  ]);
  if (!types.has(value as XlsxPivotFilter['type'])) {
    fail('invalid-document-value', 'Pivot filter type is invalid', part);
  }
  return value as XlsxPivotFilter['type'];
}

export interface XlsxPivotBudget {
  records: number;
}

export interface XlsxPivotCacheRegistryEntry {
  cacheId: number;
  fieldCount?: number;
  fields?: readonly XlsxPivotCacheField[];
  index: number;
  target: string;
}

export interface XlsxPivotCacheLoadResult {
  caches: XlsxPivotCache[];
  registry: ReadonlyMap<number, XlsxPivotCacheRegistryEntry>;
}

function fail(
  code: 'invalid-document-structure' | 'invalid-document-value',
  message: string,
  part: string,
): never {
  throw new XlsxParseError({ code, message, part, severity: 'error' });
}

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function records(value: unknown): XmlRecord[] | undefined {
  if (value === undefined) return [];
  const values: unknown[] = Array.isArray(value)
    ? [...(value as unknown[])]
    : [value];
  const output: XmlRecord[] = [];
  for (const value of values) {
    const node = record(value);
    if (!node) return undefined;
    output.push(node);
  }
  return output;
}

function attributes(node: XmlRecord): Record<string, string> {
  return (record(node.attrs) ?? {}) as Record<string, string>;
}

function localName(name: string): string {
  return name.slice(name.lastIndexOf(':') + 1);
}

function child(node: XmlRecord, name: string): unknown {
  return Object.entries(node).find(([key]) => localName(key) === name)?.[1];
}

function root(
  value: XmlLookupValue,
  local: string,
  namespace: string,
  part: string,
): XmlRecord {
  const entry = Object.entries(value).find(
    ([name]) => localName(name) === local,
  );
  const node = record(entry?.[1]);
  if (!entry || !node)
    fail('invalid-document-structure', `Pivot ${local} root is missing`, part);
  const pieces = entry[0].split(':');
  const prefix = pieces.length === 1 ? '' : pieces[0]!;
  const sourcePrefix = prefix.startsWith('ns_') ? prefix.slice(3) : prefix;
  if (
    attributes(node)[sourcePrefix ? `xmlns:${sourcePrefix}` : 'xmlns'] !==
    namespace
  ) {
    fail(
      'invalid-document-structure',
      `Pivot ${local} root has the wrong namespace`,
      part,
    );
  }
  return node;
}

function booleanValue(
  value: string | undefined,
  fallback: boolean,
  message: string,
  part: string,
): boolean {
  if (value === undefined) return fallback;
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  fail('invalid-document-value', message, part);
}

function requiredBooleanValue(
  value: string,
  message: string,
  part: string,
): boolean {
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  fail('invalid-document-value', message, part);
}

function daysInMonth(year: string, month: number): number {
  if (month === 2) {
    const tail = Number(year.slice(-4));
    const leap = tail % 400 === 0 || (tail % 4 === 0 && tail % 100 !== 0);
    return leap ? 29 : 28;
  }
  return month === 4 || month === 6 || month === 9 || month === 11 ? 30 : 31;
}

export function parseXlsxPivotDateTime(
  value: unknown,
  message: string,
  part: string,
): string {
  if (typeof value !== 'string') fail('invalid-document-value', message, part);
  const match =
    /^(\d{4,})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?(?:Z|[+-](?:(?:0\d|1[0-3]):[0-5]\d|14:00))?$/u.exec(
      value,
    );
  if (!match || /^0+$/u.test(match[1]!)) {
    fail('invalid-document-value', message, part);
  }
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);
  if (
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > daysInMonth(match[1]!, month) ||
    hour > 23 ||
    minute > 59 ||
    second > 59
  ) {
    fail('invalid-document-value', message, part);
  }
  return value;
}

export function parseXlsxPivotUnsignedInteger(
  value: unknown,
  message: string,
  part: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed))
    fail('invalid-document-value', message, part);
  return parsed;
}

export function parseXlsxPivotFiniteNumber(
  value: unknown,
  message: string,
  part: string,
): number {
  if (
    typeof value !== 'string' ||
    !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/u.test(value)
  ) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) fail('invalid-document-value', message, part);
  return Object.is(parsed, -0) ? 0 : parsed;
}

const unsignedInteger = parseXlsxPivotUnsignedInteger;
const finiteNumber = parseXlsxPivotFiniteNumber;

function text(
  value: string,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): string {
  consumeXlsxWorksheetBudget(
    budget,
    'textCharacters',
    value.length,
    'maxTextCharacters',
    limits,
    part,
  );
  return value;
}

function cacheItem(
  name: string,
  node: XmlRecord,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxPivotCacheItem {
  const value = attributes(node).v;
  if (name === 'm') return { kind: 'blank' };
  if (value === undefined) {
    fail('invalid-document-value', 'Pivot cache item value is missing', part);
  }
  if (name === 'b') {
    return {
      kind: 'boolean',
      value: requiredBooleanValue(
        value,
        'Pivot cache boolean is invalid',
        part,
      ),
    };
  }
  if (name === 'd') {
    return {
      kind: 'date',
      value: parseXlsxPivotDateTime(value, 'Pivot cache date is invalid', part),
    };
  }
  if (name === 'e')
    return { kind: 'error', value: text(value, budget, limits, part) };
  if (name === 'n') {
    return {
      kind: 'number',
      value: finiteNumber(value, 'Pivot cache number is invalid', part),
    };
  }
  if (name === 's')
    return { kind: 'text', value: text(value, budget, limits, part) };
  fail('invalid-document-structure', 'Pivot cache item type is invalid', part);
}

function sharedItems(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
):
  | {
      items: XlsxPivotCacheItem[];
      metadata: NonNullable<XlsxPivotCacheField['sharedItems']>;
    }
  | undefined {
  if (value === undefined) return undefined;
  const node = record(value);
  if (!node)
    fail('invalid-document-structure', 'Pivot shared items are invalid', part);
  const attrs = attributes(node);
  const items: XlsxPivotCacheItem[] = [];
  for (const [name, itemValue] of Object.entries(node)) {
    const itemName = localName(name);
    if (itemName === 'attrs' || itemName === 'extLst') continue;
    const itemNodes = records(itemValue);
    if (!itemNodes)
      fail(
        'invalid-document-structure',
        'Pivot shared-item collection is invalid',
        part,
      );
    for (const item of itemNodes) {
      items.push(cacheItem(itemName, item, budget, limits, part));
    }
  }
  const count = unsignedInteger(
    attrs.count,
    'Pivot shared-item count is invalid',
    part,
  );
  if (count !== undefined && count !== items.length) {
    fail(
      'invalid-document-structure',
      'Pivot shared-item count does not match',
      part,
    );
  }
  const minimumNumber =
    attrs.minValue === undefined
      ? undefined
      : finiteNumber(
          attrs.minValue,
          'Pivot shared-item minimum is invalid',
          part,
        );
  const maximumNumber =
    attrs.maxValue === undefined
      ? undefined
      : finiteNumber(
          attrs.maxValue,
          'Pivot shared-item maximum is invalid',
          part,
        );
  return {
    items,
    metadata: {
      containsBlank: booleanValue(
        attrs.containsBlank,
        false,
        'Pivot shared-items blank flag is invalid',
        part,
      ),
      containsDate: booleanValue(
        attrs.containsDate,
        false,
        'Pivot shared-items date flag is invalid',
        part,
      ),
      containsInteger: booleanValue(
        attrs.containsInteger,
        false,
        'Pivot shared-items integer flag is invalid',
        part,
      ),
      containsMixedTypes: booleanValue(
        attrs.containsMixedTypes,
        false,
        'Pivot shared-items mixed flag is invalid',
        part,
      ),
      containsNonDate: booleanValue(
        attrs.containsNonDate,
        false,
        'Pivot shared-items non-date flag is invalid',
        part,
      ),
      containsNumber: booleanValue(
        attrs.containsNumber,
        false,
        'Pivot shared-items number flag is invalid',
        part,
      ),
      containsSemiMixedTypes: booleanValue(
        attrs.containsSemiMixedTypes,
        false,
        'Pivot shared-items semi-mixed flag is invalid',
        part,
      ),
      containsString: booleanValue(
        attrs.containsString,
        false,
        'Pivot shared-items string flag is invalid',
        part,
      ),
      longText: booleanValue(
        attrs.longText,
        false,
        'Pivot shared-items long-text flag is invalid',
        part,
      ),
      ...(attrs.maxDate === undefined
        ? {}
        : {
            maximumDate: parseXlsxPivotDateTime(
              attrs.maxDate,
              'Pivot shared-item maximum date is invalid',
              part,
            ),
          }),
      ...(maximumNumber === undefined ? {} : { maximumNumber }),
      ...(attrs.minDate === undefined
        ? {}
        : {
            minimumDate: parseXlsxPivotDateTime(
              attrs.minDate,
              'Pivot shared-item minimum date is invalid',
              part,
            ),
          }),
      ...(minimumNumber === undefined ? {} : { minimumNumber }),
    },
  };
}

function cacheFields(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxPivotCacheField[] {
  if (value === undefined)
    fail('invalid-document-structure', 'Pivot cache fields are missing', part);
  const container = record(value);
  const fieldNodes = container
    ? records(child(container, 'cacheField'))
    : undefined;
  if (!container || !fieldNodes)
    fail('invalid-document-structure', 'Pivot cache fields are invalid', part);
  const expected = unsignedInteger(
    attributes(container).count,
    'Pivot cache field count is invalid',
    part,
  );
  if (expected !== fieldNodes.length) {
    fail(
      'invalid-document-structure',
      'Pivot cache field count does not match',
      part,
    );
  }
  return fieldNodes.map((field) => {
    const attrs = attributes(field);
    if (!attrs.name)
      fail('invalid-document-value', 'Pivot cache field name is invalid', part);
    const shared = sharedItems(
      child(field, 'sharedItems'),
      budget,
      limits,
      part,
    );
    return {
      databaseField: booleanValue(
        attrs.databaseField,
        true,
        'Pivot cache database-field flag is invalid',
        part,
      ),
      ...(shared === undefined
        ? {}
        : { items: shared.items, sharedItems: shared.metadata }),
      name: text(attrs.name, budget, limits, part),
      serverField: booleanValue(
        attrs.serverField,
        false,
        'Pivot cache server-field flag is invalid',
        part,
      ),
      uniqueList: booleanValue(
        attrs.uniqueList,
        true,
        'Pivot cache unique-list flag is invalid',
        part,
      ),
    };
  });
}

function cacheSource(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxPivotCacheSource {
  const node = record(value);
  if (!node)
    fail('invalid-document-structure', 'Pivot cache source is missing', part);
  const attrs = attributes(node);
  if (attrs.type === 'worksheet') {
    const worksheet = record(child(node, 'worksheetSource'));
    if (!worksheet)
      fail(
        'invalid-document-structure',
        'Pivot worksheet source is missing',
        part,
      );
    const sourceAttrs = attributes(worksheet);
    let range: ReturnType<typeof parseXlsxRangeReference>;
    if (sourceAttrs.ref !== undefined) {
      range = parseXlsxRangeReference(sourceAttrs.ref);
      if (!range) {
        fail(
          'invalid-document-value',
          'Pivot worksheet source range is invalid',
          part,
        );
      }
    }
    if (!range && !sourceAttrs.name) {
      fail(
        'invalid-document-value',
        'Pivot worksheet source requires a range or name',
        part,
      );
    }
    consumeXlsxWorksheetBudget(
      budget,
      'rangeAreas',
      range === undefined ? 0 : 1,
      'maxRangeAreas',
      limits,
      part,
    );
    return {
      kind: 'worksheet',
      ...(sourceAttrs.name === undefined
        ? {}
        : { name: text(sourceAttrs.name, budget, limits, part) }),
      ...(range === undefined ? {} : { range }),
      ...(sourceAttrs.sheet === undefined
        ? {}
        : { sheet: text(sourceAttrs.sheet, budget, limits, part) }),
    };
  }
  if (attrs.type === 'external') {
    const connectionId = unsignedInteger(
      attrs.connectionId,
      'Pivot cache connection ID is invalid',
      part,
    );
    return {
      ...(connectionId === undefined ? {} : { connectionId }),
      kind: 'external',
    };
  }
  if (attrs.type === 'consolidation') return { kind: 'consolidation' };
  if (attrs.type === 'scenario') return { kind: 'scenario' };
  fail('invalid-document-value', 'Pivot cache source type is invalid', part);
}

export class XlsxPivotCacheRecordsSink implements XlsxXmlEventSink {
  private declaredCount: number | undefined;
  private current: XlsxPivotCacheRecordValue[] | undefined;
  private records = 0;
  private readonly output: XlsxPivotCacheRecordValue[][] = [];
  private readonly stack: string[] = [];
  private rootSeen = false;

  constructor(
    private readonly namespace: string,
    private readonly fields: readonly XlsxPivotCacheField[],
    private readonly budget: XlsxPivotBudget,
    private readonly worksheetBudget: XlsxWorksheetBudget,
    private readonly limits: ResolvedXlsxResourceLimits,
    private readonly part: string,
  ) {}

  openElement(element: XlsxXmlElement): void {
    if (element.namespace !== this.namespace) {
      fail(
        'invalid-document-structure',
        'Pivot cache-record element has the wrong namespace',
        this.part,
      );
    }
    const parent = this.stack.at(-1);
    if (!parent) {
      if (element.localName !== 'pivotCacheRecords') {
        fail(
          'invalid-document-structure',
          'Pivot cache-record root is missing',
          this.part,
        );
      }
      if (this.rootSeen) {
        fail(
          'invalid-document-structure',
          'Pivot cache-record root is duplicated',
          this.part,
        );
      }
      this.rootSeen = true;
      this.declaredCount = unsignedInteger(
        element.attributes.get('{}count'),
        'Pivot cache-record count is invalid',
        this.part,
      );
    } else if (this.stack.length === 1 && element.localName === 'r') {
      const actual = this.budget.records + 1;
      if (
        !Number.isSafeInteger(actual) ||
        actual > this.limits.maxPivotRecords
      ) {
        throw new XlsxResourceLimitError(
          'maxPivotRecords',
          actual,
          this.limits.maxPivotRecords,
          this.part,
        );
      }
      this.budget.records = actual;
      this.records += 1;
      this.current = [];
    } else if (parent === 'r' && this.current) {
      const value = element.attributes.get('{}v');
      if (element.localName === 'x') {
        const index = unsignedInteger(
          value,
          'Pivot record shared-item index is invalid',
          this.part,
        );
        if (index === undefined)
          fail(
            'invalid-document-value',
            'Pivot record shared-item index is invalid',
            this.part,
          );
        this.current.push({ index, kind: 'shared-item' });
      } else {
        this.current.push(
          cacheItem(
            element.localName,
            { attrs: { v: value } },
            this.worksheetBudget,
            this.limits,
            this.part,
          ),
        );
      }
    } else {
      fail(
        'invalid-document-structure',
        'Pivot cache-record structure is invalid',
        this.part,
      );
    }
    this.stack.push(element.localName);
  }

  closeElement(element: XlsxXmlElement): void {
    if (this.stack.pop() !== element.localName) {
      fail(
        'invalid-document-structure',
        'Pivot cache-record nesting is invalid',
        this.part,
      );
    }
    if (element.localName === 'r') {
      if (!this.current || this.current.length !== this.fields.length) {
        fail(
          'invalid-document-structure',
          'Pivot cache-record field count does not match',
          this.part,
        );
      }
      this.current.forEach((value, fieldIndex) => {
        if (
          value.kind === 'shared-item' &&
          (this.fields[fieldIndex]?.items === undefined ||
            value.index >= this.fields[fieldIndex].items.length)
        ) {
          fail(
            'invalid-document-value',
            'Pivot record shared-item reference is invalid',
            this.part,
          );
        }
      });
      this.output.push(this.current);
      this.current = undefined;
    }
  }

  text(value: string): void {
    if (value.trim().length !== 0) {
      fail(
        'invalid-document-structure',
        'Pivot cache-record text is invalid',
        this.part,
      );
    }
  }

  result(): XlsxPivotCacheRecordValue[][] {
    if (!this.rootSeen || this.stack.length !== 0 || this.current) {
      fail(
        'invalid-document-structure',
        'Pivot cache-record capture is incomplete',
        this.part,
      );
    }
    if (
      this.declaredCount !== undefined &&
      this.declaredCount !== this.records
    ) {
      fail(
        'invalid-document-structure',
        'Pivot cache-record count does not match',
        this.part,
      );
    }
    return this.output;
  }
}

function recordsRelationshipTarget(
  definitionPart: string,
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
): string | undefined {
  const relationshipType = `${relationshipBase(discovery.dialect)}/pivotCacheRecords`;
  const matches = [...relationships.values()].filter(
    (relationship) => relationship.type === relationshipType,
  );
  if (matches.length === 0) return undefined;
  if (matches.length !== 1 || matches[0]!.mode !== 'internal') {
    fail(
      'invalid-document-structure',
      'Pivot cache-record relationship is invalid',
      definitionPart,
    );
  }
  const target = matches[0]!.target;
  if (
    discovery.contentTypes.contentTypeFor(target) !==
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml'
  ) {
    fail(
      'invalid-document-structure',
      'Pivot cache-record target has the wrong content type',
      target,
    );
  }
  return target;
}

export async function loadXlsxPivotCaches(
  declarations: readonly XlsxPivotCacheDeclaration[],
  mode: XlsxPivotCacheMode,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  budget: XlsxPivotBudget,
  worksheetBudget: XlsxWorksheetBudget,
): Promise<XlsxPivotCacheLoadResult> {
  const registry = new Map<number, XlsxPivotCacheRegistryEntry>();
  declarations.forEach((declaration, index) => {
    registry.set(declaration.cacheId, {
      cacheId: declaration.cacheId,
      index,
      target: declaration.target,
    });
  });
  if (mode === 'none') return { caches: [], registry };
  const caches: XlsxPivotCache[] = [];
  const namespace =
    discovery.dialect === 'strict'
      ? 'http://purl.oclc.org/ooxml/spreadsheetml/main'
      : 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  for (const [index, declaration] of declarations.entries()) {
    const value = await reader.readXml(declaration.target, { required: true });
    const definition = root(
      value,
      'pivotCacheDefinition',
      namespace,
      declaration.target,
    );
    const attrs = attributes(definition);
    const fields = cacheFields(
      child(definition, 'cacheFields'),
      worksheetBudget,
      limits,
      declaration.target,
    );
    registry.set(declaration.cacheId, {
      cacheId: declaration.cacheId,
      fieldCount: fields.length,
      fields,
      index,
      target: declaration.target,
    });
    const relationshipPart = getXlsxRelationshipPartName(declaration.target);
    const relationshipsXml = await reader.readXml(relationshipPart);
    const relationships =
      relationshipsXml === null
        ? new Map()
        : parseXlsxRelationships(
            relationshipsXml,
            declaration.target,
            limits.maxRelationships,
          );
    const recordsPart = recordsRelationshipTarget(
      declaration.target,
      relationships,
      discovery,
    );
    const recordCount = parseXlsxPivotRecordCount(
      attrs.recordCount,
      declaration.target,
    );
    if (
      mode === 'records' &&
      (recordCount ?? 0) > 0 &&
      recordsPart === undefined
    ) {
      fail(
        'invalid-document-structure',
        'Pivot cache records are missing',
        declaration.target,
      );
    }
    let loadedRecords: XlsxPivotCacheRecordValue[][] | undefined;
    if (mode === 'records' && recordsPart) {
      const sink = new XlsxPivotCacheRecordsSink(
        namespace,
        fields,
        budget,
        worksheetBudget,
        limits,
        recordsPart,
      );
      await reader.streamXml(recordsPart, sink, { required: true });
      loadedRecords = sink.result();
      if (recordCount !== undefined && recordCount !== loadedRecords.length) {
        fail(
          'invalid-document-structure',
          'Pivot cache definition record count does not match',
          declaration.target,
        );
      }
    }
    const source = cacheSource(
      child(definition, 'cacheSource'),
      worksheetBudget,
      limits,
      declaration.target,
    );
    const refreshedBy =
      attrs.refreshedBy === undefined
        ? undefined
        : text(attrs.refreshedBy, worksheetBudget, limits, declaration.target);
    const missingItemsLimit = unsignedInteger(
      attrs.missingItemsLimit,
      'Pivot cache missing-items limit is invalid',
      declaration.target,
    );
    caches.push({
      backgroundQuery: booleanValue(
        attrs.backgroundQuery,
        false,
        'Pivot cache background-query flag is invalid',
        declaration.target,
      ),
      enableRefresh: booleanValue(
        attrs.enableRefresh,
        true,
        'Pivot cache enable-refresh flag is invalid',
        declaration.target,
      ),
      fields,
      index,
      ...(missingItemsLimit === undefined ? {} : { missingItemsLimit }),
      ...(recordCount === undefined ? {} : { recordCount }),
      ...(loadedRecords === undefined ? {} : { records: loadedRecords }),
      ...(refreshedBy === undefined ? {} : { refreshedBy }),
      ...(attrs.refreshedDate === undefined
        ? {}
        : {
            refreshedDate: finiteNumber(
              attrs.refreshedDate,
              'Pivot cache refreshed date is invalid',
              declaration.target,
            ),
          }),
      refreshOnLoad: booleanValue(
        attrs.refreshOnLoad,
        false,
        'Pivot cache refresh-on-load flag is invalid',
        declaration.target,
      ),
      saveData: booleanValue(
        attrs.saveData,
        true,
        'Pivot cache save-data flag is invalid',
        declaration.target,
      ),
      source,
      supportAdvancedDrill: booleanValue(
        attrs.supportAdvancedDrill,
        true,
        'Pivot cache advanced-drill flag is invalid',
        declaration.target,
      ),
      tupleCache: booleanValue(
        attrs.tupleCache,
        false,
        'Pivot cache tuple flag is invalid',
        declaration.target,
      ),
      upgradeOnRefresh: booleanValue(
        attrs.upgradeOnRefresh,
        false,
        'Pivot cache upgrade-on-refresh flag is invalid',
        declaration.target,
      ),
    });
  }
  return { caches, registry };
}

function requiredUnsignedInteger(
  value: unknown,
  message: string,
  part: string,
): number {
  const parsed = unsignedInteger(value, message, part);
  if (parsed === undefined) fail('invalid-document-value', message, part);
  return parsed;
}

export function parseXlsxPivotSignedInteger(
  value: unknown,
  message: string,
  part: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !/^(?:0|-?[1-9]\d*)$/u.test(value)) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed))
    fail('invalid-document-value', message, part);
  return parsed;
}

const signedInteger = parseXlsxPivotSignedInteger;

export function parseXlsxPivotRecordCount(
  value: unknown,
  part: string,
): number | undefined {
  return unsignedInteger(value, 'Pivot cache record count is invalid', part);
}

function pivotFieldItems(value: unknown, part: string): XlsxPivotFieldItem[] {
  if (value === undefined) return [];
  const container = record(value);
  const itemNodes = container ? records(child(container, 'item')) : undefined;
  if (!container || !itemNodes)
    fail('invalid-document-structure', 'Pivot field items are invalid', part);
  const expected = unsignedInteger(
    attributes(container).count,
    'Pivot field item count is invalid',
    part,
  );
  if (expected !== undefined && expected !== itemNodes.length) {
    fail(
      'invalid-document-structure',
      'Pivot field item count does not match',
      part,
    );
  }
  const itemTypes: Record<string, XlsxPivotFieldItem['type']> = {
    avg: 'average',
    blank: 'blank',
    count: 'count',
    data: 'data',
    default: 'default',
    grand: 'grand-total',
    max: 'maximum',
    min: 'minimum',
    product: 'product',
    stdDev: 'standard-deviation',
    stdDevP: 'standard-deviation-population',
    sum: 'sum',
    var: 'variance',
    varP: 'variance-population',
  };
  return itemNodes.map((item) => {
    const attrs = attributes(item);
    const itemType = attrs.t === undefined ? 'data' : itemTypes[attrs.t];
    if (!itemType)
      fail('invalid-document-value', 'Pivot field item type is invalid', part);
    const sharedItemIndex = unsignedInteger(
      attrs.x,
      'Pivot field shared-item index is invalid',
      part,
    );
    return {
      ...(sharedItemIndex === undefined ? {} : { sharedItemIndex }),
      type: itemType,
    };
  });
}

function pivotFields(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxPivotField[] {
  if (value === undefined)
    fail('invalid-document-structure', 'Pivot fields are missing', part);
  const container = record(value);
  const nodes = container ? records(child(container, 'pivotField')) : undefined;
  if (!container || !nodes)
    fail('invalid-document-structure', 'Pivot fields are invalid', part);
  const expected = requiredUnsignedInteger(
    attributes(container).count,
    'Pivot field count is invalid',
    part,
  );
  if (expected !== nodes.length) {
    fail(
      'invalid-document-structure',
      'Pivot field count does not match',
      part,
    );
  }
  const axes = {
    axisCol: 'column',
    axisPage: 'page',
    axisRow: 'row',
    axisValues: 'values',
  } as const;
  return nodes.map((node) => {
    const attrs = attributes(node);
    let axis: XlsxPivotField['axis'];
    if (attrs.axis !== undefined) {
      axis = axes[attrs.axis as keyof typeof axes];
      if (axis === undefined) {
        fail('invalid-document-value', 'Pivot field axis is invalid', part);
      }
    }
    return {
      ...(axis === undefined ? {} : { axis }),
      compact: booleanValue(
        attrs.compact,
        true,
        'Pivot field compact flag is invalid',
        part,
      ),
      dataField: booleanValue(
        attrs.dataField,
        false,
        'Pivot field data flag is invalid',
        part,
      ),
      items: pivotFieldItems(child(node, 'items'), part),
      ...(attrs.name === undefined
        ? {}
        : { name: text(attrs.name, budget, limits, part) }),
      outline: booleanValue(
        attrs.outline,
        true,
        'Pivot field outline flag is invalid',
        part,
      ),
      showAll: booleanValue(
        attrs.showAll,
        true,
        'Pivot field show-all flag is invalid',
        part,
      ),
      sortType:
        attrs.sortType === undefined
          ? 'manual'
          : attrs.sortType === 'manual' ||
              attrs.sortType === 'ascending' ||
              attrs.sortType === 'descending'
            ? attrs.sortType
            : fail(
                'invalid-document-value',
                'Pivot field sort type is invalid',
                part,
              ),
      subtotalTop: booleanValue(
        attrs.subtotalTop,
        true,
        'Pivot field subtotal position is invalid',
        part,
      ),
    };
  });
}

export function parseXlsxPivotFieldIndexes(
  value: unknown,
  containerName: string,
  part: string,
): number[] {
  if (value === undefined) return [];
  const container = record(value);
  const nodes = container ? records(child(container, 'field')) : undefined;
  if (!container || !nodes)
    fail(
      'invalid-document-structure',
      `Pivot ${containerName} are invalid`,
      part,
    );
  const expected = unsignedInteger(
    attributes(container).count,
    `Pivot ${containerName} count is invalid`,
    part,
  );
  if (expected !== undefined && expected !== nodes.length) {
    fail(
      'invalid-document-structure',
      `Pivot ${containerName} count does not match`,
      part,
    );
  }
  return nodes.map((node) => {
    const index = signedInteger(
      attributes(node).x,
      `Pivot ${containerName} field index is invalid`,
      part,
    );
    if (index === undefined || index < -2) {
      fail(
        'invalid-document-value',
        `Pivot ${containerName} field index is invalid`,
        part,
      );
    }
    return index;
  });
}

const fieldIndexes = parseXlsxPivotFieldIndexes;

function dataFields(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxPivotDataField[] {
  if (value === undefined) return [];
  const container = record(value);
  const nodes = container ? records(child(container, 'dataField')) : undefined;
  if (!container || !nodes)
    fail('invalid-document-structure', 'Pivot data fields are invalid', part);
  const expected = unsignedInteger(
    attributes(container).count,
    'Pivot data-field count is invalid',
    part,
  );
  if (expected !== undefined && expected !== nodes.length) {
    fail(
      'invalid-document-structure',
      'Pivot data-field count does not match',
      part,
    );
  }
  return nodes.map((node) => {
    const attrs = attributes(node);
    const baseField = signedInteger(
      attrs.baseField,
      'Pivot data-field base field is invalid',
      part,
    );
    const baseItem = unsignedInteger(
      attrs.baseItem,
      'Pivot data-field base item is invalid',
      part,
    );
    const showDataAs = parseXlsxPivotDataDisplayMode(attrs.showDataAs, part);
    const subtotal = parseXlsxPivotSubtotal(attrs.subtotal, part);
    return {
      ...(baseField === undefined ? {} : { baseField }),
      ...(baseItem === undefined ? {} : { baseItem }),
      field: requiredUnsignedInteger(
        attrs.fld,
        'Pivot data-field field is invalid',
        part,
      ),
      ...(attrs.name === undefined
        ? {}
        : { name: text(attrs.name, budget, limits, part) }),
      showDataAs,
      subtotal,
    };
  });
}

function pageFields(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxPivotPageField[] {
  if (value === undefined) return [];
  const container = record(value);
  const nodes = container ? records(child(container, 'pageField')) : undefined;
  if (!container || !nodes)
    fail('invalid-document-structure', 'Pivot page fields are invalid', part);
  const expected = unsignedInteger(
    attributes(container).count,
    'Pivot page-field count is invalid',
    part,
  );
  if (expected !== undefined && expected !== nodes.length) {
    fail(
      'invalid-document-structure',
      'Pivot page-field count does not match',
      part,
    );
  }
  return nodes.map((node) => {
    const attrs = attributes(node);
    const hierarchy = unsignedInteger(
      attrs.hier,
      'Pivot page-field hierarchy is invalid',
      part,
    );
    const item = unsignedInteger(
      attrs.item,
      'Pivot page-field item is invalid',
      part,
    );
    return {
      field: requiredUnsignedInteger(
        attrs.fld,
        'Pivot page-field field is invalid',
        part,
      ),
      ...(hierarchy === undefined ? {} : { hierarchy }),
      ...(item === undefined ? {} : { item }),
      ...(attrs.name === undefined
        ? {}
        : { name: text(attrs.name, budget, limits, part) }),
    };
  });
}

function pivotFilters(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxPivotFilter[] {
  if (value === undefined) return [];
  const container = record(value);
  const nodes = container ? records(child(container, 'filter')) : undefined;
  if (!container || !nodes)
    fail('invalid-document-structure', 'Pivot filters are invalid', part);
  const expected = unsignedInteger(
    attributes(container).count,
    'Pivot filter count is invalid',
    part,
  );
  if (expected !== undefined && expected !== nodes.length) {
    fail(
      'invalid-document-structure',
      'Pivot filter count does not match',
      part,
    );
  }
  return nodes.map((node) => {
    const attrs = attributes(node);
    const type = parseXlsxPivotFilterType(attrs.type, part);
    const measureField = unsignedInteger(
      attrs.iMeasureFld,
      'Pivot filter measure field is invalid',
      part,
    );
    const measureHierarchy = unsignedInteger(
      attrs.iMeasureHier,
      'Pivot filter measure hierarchy is invalid',
      part,
    );
    return {
      ...(attrs.description === undefined
        ? {}
        : { description: text(attrs.description, budget, limits, part) }),
      evaluationOrder:
        signedInteger(
          attrs.evalOrder,
          'Pivot filter evaluation order is invalid',
          part,
        ) ?? 0,
      field: requiredUnsignedInteger(
        attrs.fld,
        'Pivot filter field is invalid',
        part,
      ),
      id: requiredUnsignedInteger(attrs.id, 'Pivot filter ID is invalid', part),
      ...(measureField === undefined ? {} : { measureField }),
      ...(measureHierarchy === undefined ? {} : { measureHierarchy }),
      ...(attrs.name === undefined
        ? {}
        : { name: text(attrs.name, budget, limits, part) }),
      ...(attrs.stringValue1 === undefined
        ? {}
        : { stringValue1: text(attrs.stringValue1, budget, limits, part) }),
      ...(attrs.stringValue2 === undefined
        ? {}
        : { stringValue2: text(attrs.stringValue2, budget, limits, part) }),
      type,
    };
  });
}

function selectionRelation(
  range: NonNullable<ReturnType<typeof parseXlsxRangeReference>>,
  selection: XlsxResolvedSheetSelection,
): XlsxPivotTable['selectionRelation'] | null {
  if (selection.kind === 'full-sheet') return 'full-sheet';
  if (selection.kind === 'not-selected') return null;
  return selection.ranges.some(
    (selected) =>
      selected.start.row <= range.end.row &&
      selected.end.row >= range.start.row &&
      selected.start.column <= range.end.column &&
      selected.end.column >= range.start.column,
  )
    ? 'intersects-selection'
    : null;
}

export async function loadXlsxPivotTables(
  relationshipIds: readonly string[],
  worksheetRelationships: ReadonlyMap<string, XlsxRelationship>,
  registry: ReadonlyMap<number, XlsxPivotCacheRegistryEntry>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  worksheetBudget: XlsxWorksheetBudget,
  selection: XlsxResolvedSheetSelection,
  worksheetPart: string,
): Promise<XlsxPivotTable[]> {
  const output: XlsxPivotTable[] = [];
  const names = new Set<string>();
  const relationshipType = `${relationshipBase(discovery.dialect)}/pivotTable`;
  const namespace =
    discovery.dialect === 'strict'
      ? 'http://purl.oclc.org/ooxml/spreadsheetml/main'
      : 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  for (const relationshipId of relationshipIds) {
    const relationship = worksheetRelationships.get(relationshipId);
    if (
      !relationship ||
      relationship.mode !== 'internal' ||
      relationship.type !== relationshipType
    ) {
      fail(
        'invalid-document-structure',
        'Worksheet pivot-table relationship is invalid',
        worksheetPart,
      );
    }
    if (
      discovery.contentTypes.contentTypeFor(relationship.target) !==
      'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml'
    ) {
      fail(
        'invalid-document-structure',
        'Pivot table has the wrong content type',
        relationship.target,
      );
    }
    const value = await reader.readXml(relationship.target, { required: true });
    const definition = root(
      value,
      'pivotTableDefinition',
      namespace,
      relationship.target,
    );
    const attrs = attributes(definition);
    if (!attrs.name || !attrs.dataCaption) {
      fail(
        'invalid-document-value',
        'Pivot table name or data caption is invalid',
        relationship.target,
      );
    }
    const foldedName = attrs.name.toUpperCase();
    if (names.has(foldedName)) {
      fail(
        'invalid-document-value',
        'Worksheet contains duplicate pivot table names',
        worksheetPart,
      );
    }
    names.add(foldedName);
    const cacheId = requiredUnsignedInteger(
      attrs.cacheId,
      'Pivot table cache ID is invalid',
      relationship.target,
    );
    const cache = registry.get(cacheId);
    if (!cache)
      fail(
        'invalid-document-value',
        'Pivot table cache reference is invalid',
        relationship.target,
      );
    const locationNode = record(child(definition, 'location'));
    const location = locationNode
      ? parseXlsxRangeReference(attributes(locationNode).ref)
      : undefined;
    if (!location)
      fail(
        'invalid-document-value',
        'Pivot table location is invalid',
        relationship.target,
      );
    consumeXlsxWorksheetBudget(
      worksheetBudget,
      'rangeAreas',
      1,
      'maxRangeAreas',
      limits,
      relationship.target,
    );
    const relation = selectionRelation(location, selection);
    const fields = pivotFields(
      child(definition, 'pivotFields'),
      worksheetBudget,
      limits,
      relationship.target,
    );
    if (cache.fieldCount !== undefined && cache.fieldCount !== fields.length) {
      fail(
        'invalid-document-structure',
        'Pivot table field count does not match its cache',
        relationship.target,
      );
    }
    fields.forEach((field, fieldIndex) => {
      field.items.forEach((item) => {
        if (
          item.sharedItemIndex !== undefined &&
          cache.fields !== undefined &&
          (cache.fields[fieldIndex]?.items === undefined ||
            item.sharedItemIndex >= cache.fields[fieldIndex].items.length)
        ) {
          fail(
            'invalid-document-value',
            'Pivot field shared-item reference is invalid',
            relationship.target,
          );
        }
      });
    });
    const columnFields = fieldIndexes(
      child(definition, 'colFields'),
      'column fields',
      relationship.target,
    );
    const rowFields = fieldIndexes(
      child(definition, 'rowFields'),
      'row fields',
      relationship.target,
    );
    for (const field of [...columnFields, ...rowFields]) {
      if (field >= fields.length) {
        fail(
          'invalid-document-value',
          'Pivot axis field reference is invalid',
          relationship.target,
        );
      }
    }
    const parsedDataFields = dataFields(
      child(definition, 'dataFields'),
      worksheetBudget,
      limits,
      relationship.target,
    );
    if (parsedDataFields.some((field) => field.field >= fields.length)) {
      fail(
        'invalid-document-value',
        'Pivot data-field reference is invalid',
        relationship.target,
      );
    }
    const parsedPageFields = pageFields(
      child(definition, 'pageFields'),
      worksheetBudget,
      limits,
      relationship.target,
    );
    if (parsedPageFields.some((field) => field.field >= fields.length)) {
      fail(
        'invalid-document-value',
        'Pivot page-field reference is invalid',
        relationship.target,
      );
    }
    const parsedFilters = pivotFilters(
      child(definition, 'filters'),
      worksheetBudget,
      limits,
      relationship.target,
    );
    if (parsedFilters.some((filter) => filter.field >= fields.length)) {
      fail(
        'invalid-document-value',
        'Pivot filter field reference is invalid',
        relationship.target,
      );
    }
    const styleNode = record(child(definition, 'pivotTableStyleInfo'));
    const styleAttrs = styleNode ? attributes(styleNode) : {};
    const table: XlsxPivotTable = {
      cacheIndex: cache.index,
      columnFields,
      compact: booleanValue(
        attrs.compact,
        true,
        'Pivot table compact flag is invalid',
        relationship.target,
      ),
      dataCaption: text(
        attrs.dataCaption,
        worksheetBudget,
        limits,
        relationship.target,
      ),
      dataFields: parsedDataFields,
      fields,
      filters: parsedFilters,
      ...(attrs.grandTotalCaption === undefined
        ? {}
        : {
            grandTotalCaption: text(
              attrs.grandTotalCaption,
              worksheetBudget,
              limits,
              relationship.target,
            ),
          }),
      location,
      name: text(attrs.name, worksheetBudget, limits, relationship.target),
      outline: booleanValue(
        attrs.outline,
        false,
        'Pivot table outline flag is invalid',
        relationship.target,
      ),
      pageFields: parsedPageFields,
      rowFields,
      selectionRelation: relation ?? 'full-sheet',
      showColumnGrandTotals: booleanValue(
        attrs.colGrandTotals,
        true,
        'Pivot table column-grand-total flag is invalid',
        relationship.target,
      ),
      showHeaders: booleanValue(
        attrs.showHeaders,
        true,
        'Pivot table header flag is invalid',
        relationship.target,
      ),
      showRowGrandTotals: booleanValue(
        attrs.rowGrandTotals,
        true,
        'Pivot table row-grand-total flag is invalid',
        relationship.target,
      ),
      style: {
        ...(styleAttrs.name === undefined
          ? {}
          : {
              name: text(
                styleAttrs.name,
                worksheetBudget,
                limits,
                relationship.target,
              ),
            }),
        showColumnHeaders: booleanValue(
          styleAttrs.showColHeaders,
          true,
          'Pivot style column-header flag is invalid',
          relationship.target,
        ),
        showColumnStripes: booleanValue(
          styleAttrs.showColStripes,
          false,
          'Pivot style column-stripe flag is invalid',
          relationship.target,
        ),
        showLastColumn: booleanValue(
          styleAttrs.showLastColumn,
          false,
          'Pivot style last-column flag is invalid',
          relationship.target,
        ),
        showRowHeaders: booleanValue(
          styleAttrs.showRowHeaders,
          true,
          'Pivot style row-header flag is invalid',
          relationship.target,
        ),
        showRowStripes: booleanValue(
          styleAttrs.showRowStripes,
          false,
          'Pivot style row-stripe flag is invalid',
          relationship.target,
        ),
      },
    };
    if (relation) output.push(table);
  }
  return output;
}
