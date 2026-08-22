import type { XmlLookupValue } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type {
  XlsxSlicer,
  XlsxSlicerCache,
  XlsxSlicerPivotTableOwner,
  XlsxTimeline,
  XlsxTimelineCache,
} from '../types';
import { XlsxPartReader } from './part-reader';
import { parseXlsxPivotDateTime } from './pivot';
import type { XlsxRelationship } from './relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import type { XlsxResolvedSheetSelection } from './selection';
import type { XlsxWorkbookDiscovery } from './workbook-discovery';
import {
  consumeXlsxWorksheetBudget,
  type XlsxWorksheetBudget,
} from './worksheet';

type XmlRecord = Record<string, unknown>;

function extensionNamespace(kind: 'slicer' | 'timeline'): string {
  return kind === 'slicer'
    ? 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
    : 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
}

function extensionRelationship(
  kind: 'slicer' | 'timeline',
  cache: boolean,
): string {
  if (kind === 'slicer') {
    return cache
      ? 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
      : 'http://schemas.microsoft.com/office/2007/relationships/slicer';
  }
  return cache
    ? 'http://schemas.microsoft.com/office/2011/relationships/timelineCache'
    : 'http://schemas.microsoft.com/office/2011/relationships/timeline';
}

export interface XlsxSlicerBudget {
  objects: number;
}

export interface XlsxAnalyticCacheRegistryEntry {
  index: number;
  kind: 'slicer' | 'timeline';
}

export interface XlsxAnalyticCacheLoadResult {
  registry: ReadonlyMap<string, XlsxAnalyticCacheRegistryEntry>;
  slicerCaches: XlsxSlicerCache[];
  timelineCaches: XlsxTimelineCache[];
}

export interface XlsxAnalyticDisplayLoadResult {
  slicers: XlsxSlicer[];
  timelines: XlsxTimeline[];
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
  const values = Array.isArray(value) ? value : [value];
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

function root(
  value: XmlLookupValue,
  expectedLocalName: string,
  expectedNamespace: string,
  part: string,
): XmlRecord {
  const entry = Object.entries(value).find(
    ([name]) => localName(name) === expectedLocalName,
  );
  const node = record(entry?.[1]);
  if (!entry || !node) {
    fail(
      'invalid-document-structure',
      `XLSX ${expectedLocalName} root is missing`,
      part,
    );
  }
  const prefix = entry[0].includes(':') ? entry[0].split(':')[0]! : '';
  const sourcePrefix = prefix.startsWith('ns_') ? prefix.slice(3) : prefix;
  const namespace =
    attributes(node)[sourcePrefix ? `xmlns:${sourcePrefix}` : 'xmlns'];
  if (namespace !== expectedNamespace) {
    fail(
      'invalid-document-structure',
      `XLSX ${expectedLocalName} root has the wrong namespace`,
      part,
    );
  }
  return node;
}

function namespaceBindings(
  parent: ReadonlyMap<string, string>,
  node: XmlRecord,
): Map<string, string> {
  const output = new Map(parent);
  for (const [name, value] of Object.entries(attributes(node))) {
    if (name === 'xmlns') output.set('', value);
    else if (name.startsWith('xmlns:')) output.set(name.slice(6), value);
  }
  return output;
}

function elementNamespace(
  key: string,
  node: XmlRecord,
  parent: ReadonlyMap<string, string>,
): string | undefined {
  const rawPrefix = key.includes(':') ? key.split(':')[0]! : '';
  const prefix = rawPrefix.startsWith('ns_') ? rawPrefix.slice(3) : rawPrefix;
  return namespaceBindings(parent, node).get(prefix);
}

function children(
  node: XmlRecord,
  name: string,
  namespace: string,
  part: string,
): XmlRecord[] {
  const output: XmlRecord[] = [];
  const bindings = namespaceBindings(new Map(), node);
  for (const [key, value] of Object.entries(node)) {
    if (localName(key) !== name) continue;
    const values = records(value);
    if (!values) {
      fail(
        'invalid-document-structure',
        `XLSX ${name} collection is invalid`,
        part,
      );
    }
    for (const child of values) {
      if (elementNamespace(key, child, bindings) !== namespace) {
        fail(
          'invalid-document-structure',
          `XLSX ${name} element has the wrong namespace`,
          part,
        );
      }
      output.push(child);
    }
  }
  return output;
}

function descendants(
  node: XmlRecord,
  name: string,
  namespace: string,
  parentBindings: ReadonlyMap<string, string> = new Map(),
): XmlRecord[] {
  const output: XmlRecord[] = [];
  const bindings = namespaceBindings(parentBindings, node);
  for (const [key, value] of Object.entries(node)) {
    const values = records(value);
    if (!values) continue;
    for (const child of values) {
      if (
        localName(key) === name &&
        elementNamespace(key, child, bindings) === namespace
      ) {
        output.push(child);
      }
      output.push(...descendants(child, name, namespace, bindings));
    }
  }
  return output;
}

export function parseXlsxSlicerUnsignedInteger(
  value: string | undefined,
  fallback: number | undefined,
  message: string,
  part: string,
): number | undefined {
  if (value === undefined) return fallback;
  if (!/^(?:0|[1-9]\d*)$/u.test(value)) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > 0xffff_ffff) {
    fail('invalid-document-value', message, part);
  }
  return parsed;
}

const unsignedInteger = parseXlsxSlicerUnsignedInteger;

export function parseXlsxSlicerBoolean(
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

const booleanValue = parseXlsxSlicerBoolean;

export function parseXlsxSlicerSortOrder(
  value: string | undefined,
  part: string,
): XlsxSlicerCache['sortOrder'] {
  if (value !== undefined && value !== 'ascending' && value !== 'descending') {
    fail('invalid-document-value', 'Slicer sort order is invalid', part);
  }
  return value;
}

export function parseXlsxSlicerCrossFilter(
  value: string | undefined,
  part: string,
): XlsxSlicerCache['crossFilter'] {
  if (
    value !== undefined &&
    value !== 'none' &&
    value !== 'showItemsWithDataAtTop' &&
    value !== 'showItemsWithNoData'
  ) {
    fail('invalid-document-value', 'Slicer cross-filter is invalid', part);
  }
  return value;
}

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

function requiredText(
  value: string | undefined,
  message: string,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): string {
  if (!value) fail('invalid-document-value', message, part);
  return text(value, budget, limits, part);
}

function relationshipParts(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  types: ReadonlyMap<
    string,
    { contentType: string; kind: 'slicer' | 'timeline' }
  >,
  discovery: XlsxWorkbookDiscovery,
  ownerPart: string,
): Array<{ kind: 'slicer' | 'timeline'; target: string }> {
  const output: Array<{ kind: 'slicer' | 'timeline'; target: string }> = [];
  const targets = new Set<string>();
  for (const relationship of relationships.values()) {
    const expected = types.get(relationship.type);
    if (!expected) continue;
    if (relationship.mode !== 'internal') {
      fail(
        'invalid-document-structure',
        'XLSX slicer or timeline relationship must be internal',
        ownerPart,
      );
    }
    if (targets.has(relationship.target)) {
      fail(
        'invalid-document-structure',
        'XLSX slicer or timeline relationship target is duplicated',
        ownerPart,
      );
    }
    targets.add(relationship.target);
    if (
      discovery.contentTypes.contentTypeFor(relationship.target) !==
      expected.contentType
    ) {
      fail(
        'invalid-document-structure',
        'XLSX slicer or timeline target has the wrong content type',
        relationship.target,
      );
    }
    output.push({ kind: expected.kind, target: relationship.target });
  }
  return output;
}

function pivotTableOwners(
  definition: XmlRecord,
  namespace: string,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxSlicerPivotTableOwner[] {
  return descendants(definition, 'pivotTable', namespace).map((node) => {
    const attrs = attributes(node);
    const sheetId = unsignedInteger(
      attrs.tabId,
      undefined,
      'Slicer pivot-table sheet ID is invalid',
      part,
    );
    if (sheetId === undefined || !attrs.name) {
      fail(
        'invalid-document-value',
        'Slicer pivot-table owner is invalid',
        part,
      );
    }
    return {
      name: text(attrs.name, budget, limits, part),
      sheetId,
    };
  });
}

function cachePivotId(
  definition: XmlRecord,
  namespace: string,
  part: string,
): number | undefined {
  const value =
    attributes(definition).pivotCacheId ??
    ['tabular', 'olap', 'timelineState', 'timelinePivotCacheDefinition']
      .flatMap((name) => descendants(definition, name, namespace))
      .map((node) => attributes(node).pivotCacheId)
      .find((candidate) => candidate !== undefined);
  return unsignedInteger(
    value,
    undefined,
    'Slicer pivot cache ID is invalid',
    part,
  );
}

export async function loadXlsxAnalyticCaches(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  worksheetBudget: XlsxWorksheetBudget,
): Promise<XlsxAnalyticCacheLoadResult> {
  const types = new Map([
    [
      extensionRelationship('slicer', true),
      {
        contentType: 'application/vnd.ms-excel.slicerCache+xml',
        kind: 'slicer' as const,
      },
    ],
    [
      extensionRelationship('timeline', true),
      {
        contentType: 'application/vnd.ms-excel.timelineCache+xml',
        kind: 'timeline' as const,
      },
    ],
  ]);
  const parts = relationshipParts(
    relationships,
    types,
    discovery,
    discovery.part,
  );
  const slicerCaches: XlsxSlicerCache[] = [];
  const timelineCaches: XlsxTimelineCache[] = [];
  const registry = new Map<string, XlsxAnalyticCacheRegistryEntry>();
  for (const part of parts) {
    const value = await reader.readXml(part.target, { required: true });
    const namespace = extensionNamespace(part.kind);
    const local =
      part.kind === 'slicer'
        ? 'slicerCacheDefinition'
        : 'timelineCacheDefinition';
    const definition = root(value, local, namespace, part.target);
    const attrs = attributes(definition);
    const name = requiredText(
      attrs.name,
      'Slicer or timeline cache name is invalid',
      worksheetBudget,
      limits,
      part.target,
    );
    const sourceName = requiredText(
      attrs.sourceName,
      'Slicer or timeline cache source name is invalid',
      worksheetBudget,
      limits,
      part.target,
    );
    const folded = name.toUpperCase();
    if (registry.has(folded)) {
      fail(
        'invalid-document-value',
        'Workbook contains duplicate slicer or timeline cache names',
        part.target,
      );
    }
    const pivotTables = pivotTableOwners(
      definition,
      namespace,
      worksheetBudget,
      limits,
      part.target,
    );
    const pivotCacheId = cachePivotId(definition, namespace, part.target);
    if (part.kind === 'slicer') {
      const tabular = descendants(
        definition,
        'tabular',
        extensionNamespace('slicer'),
      )[0];
      const olap = descendants(
        definition,
        'olap',
        extensionNamespace('slicer'),
      )[0];
      const table = descendants(
        definition,
        'tableSlicerCache',
        extensionNamespace('timeline'),
      )[0];
      const source = tabular ?? olap ?? table;
      if (!source) {
        fail(
          'invalid-document-structure',
          'Slicer cache source metadata is missing',
          part.target,
        );
      }
      const sourceAttrs = attributes(source);
      const sortOrder = parseXlsxSlicerSortOrder(
        sourceAttrs.sortOrder,
        part.target,
      );
      const crossFilter = parseXlsxSlicerCrossFilter(
        sourceAttrs.crossFilter,
        part.target,
      );
      const tableOwner = table
        ? {
            column: unsignedInteger(
              sourceAttrs.column,
              undefined,
              'Table slicer column is invalid',
              part.target,
            ),
            id: unsignedInteger(
              sourceAttrs.tableId,
              undefined,
              'Table slicer table ID is invalid',
              part.target,
            ),
          }
        : undefined;
      if (
        tableOwner !== undefined &&
        (tableOwner.column === undefined || tableOwner.id === undefined)
      ) {
        fail(
          'invalid-document-value',
          'Table slicer owner is invalid',
          part.target,
        );
      }
      const index = slicerCaches.length;
      slicerCaches.push({
        ...(crossFilter === undefined ? {} : { crossFilter }),
        ...(!olap
          ? {
              customListSort: booleanValue(
                sourceAttrs.customListSort,
                true,
                'Slicer custom-list-sort flag is invalid',
                part.target,
              ),
            }
          : {}),
        index,
        name,
        ...(pivotCacheId === undefined ? {} : { pivotCacheId }),
        pivotTables,
        ...(sortOrder === undefined ? {} : { sortOrder }),
        ...(tabular
          ? {
              showMissing: booleanValue(
                sourceAttrs.showMissing,
                true,
                'Slicer show-missing flag is invalid',
                part.target,
              ),
            }
          : {}),
        sourceKind: tabular ? 'tabular' : olap ? 'olap' : 'table',
        sourceName,
        ...(tableOwner === undefined
          ? {}
          : {
              table: {
                column: tableOwner.column!,
                id: tableOwner.id!,
              },
            }),
      });
      registry.set(folded, { index, kind: 'slicer' });
    } else {
      const index = timelineCaches.length;
      timelineCaches.push({
        index,
        name,
        ...(pivotCacheId === undefined ? {} : { pivotCacheId }),
        pivotTables,
        sourceName,
      });
      registry.set(folded, { index, kind: 'timeline' });
    }
  }
  return { registry, slicerCaches, timelineCaches };
}

function consumeObject(
  budget: XlsxSlicerBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): void {
  const actual = budget.objects + 1;
  if (!Number.isSafeInteger(actual) || actual > limits.maxDrawings) {
    throw new XlsxResourceLimitError(
      'maxDrawings',
      actual,
      limits.maxDrawings,
      part,
    );
  }
  budget.objects = actual;
}

export async function loadXlsxAnalyticDisplays(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  registry: ReadonlyMap<string, XlsxAnalyticCacheRegistryEntry>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  worksheetBudget: XlsxWorksheetBudget,
  budget: XlsxSlicerBudget,
  selection: XlsxResolvedSheetSelection,
  worksheetPart: string,
): Promise<XlsxAnalyticDisplayLoadResult> {
  const types = new Map([
    [
      extensionRelationship('slicer', false),
      {
        contentType: 'application/vnd.ms-excel.slicer+xml',
        kind: 'slicer' as const,
      },
    ],
    [
      extensionRelationship('timeline', false),
      {
        contentType: 'application/vnd.ms-excel.timeline+xml',
        kind: 'timeline' as const,
      },
    ],
  ]);
  const parts = relationshipParts(
    relationships,
    types,
    discovery,
    worksheetPart,
  );
  const slicers: XlsxSlicer[] = [];
  const timelines: XlsxTimeline[] = [];
  const names = new Set<string>();
  for (const part of parts) {
    const value = await reader.readXml(part.target, { required: true });
    const namespace = extensionNamespace(part.kind);
    const rootName = part.kind === 'slicer' ? 'slicers' : 'timelines';
    const itemName = part.kind === 'slicer' ? 'slicer' : 'timeline';
    const definition = root(value, rootName, namespace, part.target);
    const items = children(definition, itemName, namespace, part.target);
    if (items.length === 0) {
      fail(
        'invalid-document-structure',
        'XLSX slicer or timeline display collection is empty',
        part.target,
      );
    }
    for (const item of items) {
      consumeObject(budget, limits, part.target);
      const attrs = attributes(item);
      const name = requiredText(
        attrs.name,
        'Slicer or timeline name is invalid',
        worksheetBudget,
        limits,
        part.target,
      );
      const cacheName = requiredText(
        attrs.cache,
        'Slicer or timeline cache name is invalid',
        worksheetBudget,
        limits,
        part.target,
      );
      const foldedName = name.toUpperCase();
      if (names.has(foldedName)) {
        fail(
          'invalid-document-value',
          'Worksheet contains duplicate slicer or timeline names',
          part.target,
        );
      }
      names.add(foldedName);
      const cache = registry.get(cacheName.toUpperCase());
      if (!cache || cache.kind !== part.kind) {
        fail(
          'invalid-document-value',
          'Slicer or timeline cache reference is invalid',
          part.target,
        );
      }
      const caption =
        attrs.caption === undefined
          ? undefined
          : text(attrs.caption, worksheetBudget, limits, part.target);
      if (part.kind === 'slicer') {
        const columnCount = unsignedInteger(
          attrs.columnCount,
          1,
          'Slicer column count is invalid',
          part.target,
        )!;
        if (columnCount === 0) {
          fail(
            'invalid-document-value',
            'Slicer column count is invalid',
            part.target,
          );
        }
        const rowHeight = unsignedInteger(
          attrs.rowHeight,
          undefined,
          'Slicer row height is invalid',
          part.target,
        );
        if (rowHeight === undefined) {
          fail(
            'invalid-document-value',
            'Slicer row height is invalid',
            part.target,
          );
        }
        const level = unsignedInteger(
          attrs.level,
          undefined,
          'Slicer level is invalid',
          part.target,
        );
        const style =
          attrs.style === undefined
            ? undefined
            : text(attrs.style, worksheetBudget, limits, part.target);
        if (selection.kind === 'full-sheet') {
          slicers.push({
            cacheIndex: cache.index,
            ...(caption === undefined ? {} : { caption }),
            columnCount,
            ...(level === undefined ? {} : { level }),
            lockedPosition: booleanValue(
              attrs.lockedPosition,
              false,
              'Slicer locked-position flag is invalid',
              part.target,
            ),
            name,
            rowHeight,
            selectionRelation: 'full-sheet',
            showCaption: booleanValue(
              attrs.showCaption,
              true,
              'Slicer show-caption flag is invalid',
              part.target,
            ),
            startItem: unsignedInteger(
              attrs.startItem,
              0,
              'Slicer start item is invalid',
              part.target,
            )!,
            ...(style === undefined ? {} : { style }),
          });
        }
      } else {
        const level = unsignedInteger(
          attrs.level,
          undefined,
          'Timeline level is invalid',
          part.target,
        );
        const selectionLevel = unsignedInteger(
          attrs.selectionLevel,
          undefined,
          'Timeline selection level is invalid',
          part.target,
        );
        if (level === undefined || selectionLevel === undefined) {
          fail(
            'invalid-document-value',
            'Timeline level or selection level is invalid',
            part.target,
          );
        }
        const scrollPosition =
          attrs.scrollPosition === undefined
            ? undefined
            : parseXlsxPivotDateTime(
                attrs.scrollPosition,
                'Timeline scroll position is invalid',
                part.target,
              );
        const style =
          attrs.style === undefined
            ? undefined
            : text(attrs.style, worksheetBudget, limits, part.target);
        if (selection.kind === 'full-sheet') {
          timelines.push({
            cacheIndex: cache.index,
            ...(caption === undefined ? {} : { caption }),
            level,
            name,
            ...(scrollPosition === undefined ? {} : { scrollPosition }),
            selectionLevel,
            selectionRelation: 'full-sheet',
            showHeader: booleanValue(
              attrs.showHeader,
              true,
              'Timeline show-header flag is invalid',
              part.target,
            ),
            showHorizontalScrollbar: booleanValue(
              attrs.showHorizontalScrollbar,
              true,
              'Timeline horizontal-scrollbar flag is invalid',
              part.target,
            ),
            showSelectionLabel: booleanValue(
              attrs.showSelectionLabel,
              true,
              'Timeline selection-label flag is invalid',
              part.target,
            ),
            showTimeLevel: booleanValue(
              attrs.showTimeLevel,
              true,
              'Timeline time-level flag is invalid',
              part.target,
            ),
            ...(style === undefined ? {} : { style }),
          });
        }
      }
    }
  }
  return { slicers, timelines };
}
