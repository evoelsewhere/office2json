import type { XmlLookupValue } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type {
  XlsxChartSheet,
  XlsxSheet,
  XlsxWorkbookProperties,
  XlsxWorksheet,
} from '../types';
import { getXlsxRelationshipPartName } from './package-identity';
import { XlsxPartReader } from './part-reader';
import { parseXlsxRelationships, type XlsxRelationship } from './relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import type { XlsxWorkbookDiscovery } from './workbook-discovery';
import { parseXlsxDefinedNames } from './workbook-defined-names';
import { parseXlsxWorkbookViews } from './workbook-views';
import { parseXlsxWorkbookProtection } from './workbook-protection';

const RELATIONSHIP_BASE = {
  strict: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
  transitional:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
} as const;

type XmlRecord = Record<string, unknown>;

function expectedSheetContentType(kind: XlsxSheet['kind']): string {
  return kind === 'worksheet'
    ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
    : 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml';
}

export interface XlsxWorkbookManifest {
  pivotCaches: XlsxPivotCacheDeclaration[];
  properties: XlsxWorkbookProperties;
  protectionTextCharacters: number;
  sheetIdIndexes: ReadonlyMap<number, number>;
  sheetParts: string[];
  sheets: XlsxSheet[];
  workbookRelationships: ReadonlyMap<string, XlsxRelationship>;
}

export interface XlsxPivotCacheDeclaration {
  cacheId: number;
  target: string;
}

function fail(message: string, part: string): never {
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
  const items: unknown[] = Array.isArray(value) ? value : [value];
  const output: XmlRecord[] = [];
  for (const item of items) {
    const parsed = record(item);
    if (!parsed) return undefined;
    output.push(parsed);
  }
  return output;
}

function attributes(value: XmlRecord): XmlRecord {
  return record(value.attrs) ?? {};
}

function rootEntry(value: XmlLookupValue): {
  node: XmlRecord;
  prefix: string;
} {
  const document = value as unknown as XmlRecord;
  const [qualifiedName, node] = Object.entries(document)[0]!;
  const [first, second] = qualifiedName.split(':') as [string, string?];
  return {
    node: node as XmlRecord,
    prefix: second === undefined ? '' : first,
  };
}

function child(node: XmlRecord, prefix: string, localName: string): unknown {
  return node[prefix ? `${prefix}:${localName}` : localName];
}

function parseBoolean(value: unknown, message: string, part: string): boolean {
  if (value === undefined) return false;
  if (value === '1' || value === 'true') return true;
  if (value === '0' || value === 'false') return false;
  fail(message, part);
}

function parseProperties(
  root: XmlRecord,
  prefix: string,
  part: string,
  definedNames: XlsxWorkbookProperties['definedNames'],
  views: XlsxWorkbookProperties['views'],
  protection: XlsxWorkbookProperties['protection'],
): XlsxWorkbookProperties {
  const workbookPr = record(child(root, prefix, 'workbookPr'));
  const workbookAttrs = workbookPr ? attributes(workbookPr) : {};
  const date1904 = parseBoolean(
    workbookAttrs.date1904,
    'Workbook date1904 flag is invalid',
    part,
  );

  const calcPr = record(child(root, prefix, 'calcPr'));
  const calcAttrs = calcPr ? attributes(calcPr) : {};
  const mode = calcAttrs.calcMode ?? 'auto';
  if (mode !== 'auto' && mode !== 'autoNoTable' && mode !== 'manual') {
    fail('Workbook calculation mode is invalid', part);
  }
  const calculationId = unsignedInteger(calcAttrs.calcId);
  if (
    calcAttrs.calcId !== undefined &&
    (calculationId === undefined || calculationId > 0xffff_ffff)
  ) {
    fail('Workbook calculation ID is invalid', part);
  }
  const concurrentManualCount = unsignedInteger(
    calcAttrs.concurrentManualCount,
  );
  if (
    calcAttrs.concurrentManualCount !== undefined &&
    (concurrentManualCount === undefined || concurrentManualCount > 0xffff_ffff)
  ) {
    fail('Workbook concurrent manual count is invalid', part);
  }
  const iterateCount =
    calcAttrs.iterateCount === undefined
      ? 100
      : unsignedInteger(calcAttrs.iterateCount);
  if (iterateCount === undefined || iterateCount > 0xffff_ffff) {
    fail('Workbook iteration count is invalid', part);
  }
  const iterateDeltaText = calcAttrs.iterateDelta ?? '0.001';
  if (
    typeof iterateDeltaText !== 'string' ||
    !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/u.test(iterateDeltaText)
  ) {
    fail('Workbook iteration delta is invalid', part);
  }
  const iterateDelta = Number(iterateDeltaText);
  if (!Number.isFinite(iterateDelta)) {
    fail('Workbook iteration delta is invalid', part);
  }
  const referenceMode = calcAttrs.refMode ?? 'A1';
  if (referenceMode !== 'A1' && referenceMode !== 'R1C1') {
    fail('Workbook calculation reference mode is invalid', part);
  }

  return {
    calculation: {
      calculationCompleted:
        calcAttrs.calcCompleted === undefined
          ? true
          : parseBoolean(
              calcAttrs.calcCompleted,
              'Workbook calculation-completed flag is invalid',
              part,
            ),
      ...(calculationId === undefined ? {} : { calculationId }),
      calculateOnSave:
        calcAttrs.calcOnSave === undefined
          ? true
          : parseBoolean(
              calcAttrs.calcOnSave,
              'Workbook calculate-on-save flag is invalid',
              part,
            ),
      concurrentCalculation:
        calcAttrs.concurrentCalc === undefined
          ? true
          : parseBoolean(
              calcAttrs.concurrentCalc,
              'Workbook concurrent-calculation flag is invalid',
              part,
            ),
      ...(concurrentManualCount === undefined ? {} : { concurrentManualCount }),
      forceFullCalculation: parseBoolean(
        calcAttrs.forceFullCalc,
        'Workbook force-full-calculation flag is invalid',
        part,
      ),
      fullCalculationOnLoad: parseBoolean(
        calcAttrs.fullCalcOnLoad,
        'Workbook full-calculation-on-load flag is invalid',
        part,
      ),
      fullPrecision:
        calcAttrs.fullPrecision === undefined
          ? true
          : parseBoolean(
              calcAttrs.fullPrecision,
              'Workbook full-precision flag is invalid',
              part,
            ),
      iteration: {
        enabled: parseBoolean(
          calcAttrs.iterate,
          'Workbook iteration flag is invalid',
          part,
        ),
        maxChange: Object.is(iterateDelta, -0) ? 0 : iterateDelta,
        maxIterations: iterateCount,
      },
      mode:
        mode === 'autoNoTable'
          ? 'automatic-except-tables'
          : mode === 'manual'
            ? 'manual'
            : 'automatic',
      referenceMode,
    },
    dateSystem: date1904 ? '1904' : '1900',
    definedNames,
    ...(protection === undefined ? {} : { protection }),
    views,
  };
}

function parseSheetState(
  value: unknown,
  part: string,
): 'hidden' | 'very-hidden' | 'visible' {
  if (value === undefined || value === 'visible') return 'visible';
  if (value === 'hidden') return 'hidden';
  if (value === 'veryHidden') return 'very-hidden';
  fail('Workbook sheet state is invalid', part);
}

function validSheetName(value: unknown): value is string {
  if (typeof value !== 'string' || value.length === 0 || value.length > 31) {
    return false;
  }
  for (const character of value) {
    const code = character.codePointAt(0) ?? 0;
    if (code === 0x7f) return false;
    if (
      character === '\\' ||
      character === '/' ||
      character === ':' ||
      character === '?' ||
      character === '*' ||
      character === '[' ||
      character === ']'
    ) {
      return false;
    }
  }
  return true;
}

function positiveSheetId(value: unknown): number | undefined {
  if (typeof value !== 'string' || !/^[1-9]\d*$/u.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed <= 0xffff_ffff
    ? parsed
    : undefined;
}

function unsignedInteger(value: unknown): number | undefined {
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    return undefined;
  }
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}

export async function parseXlsxWorkbookManifest(
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
): Promise<XlsxWorkbookManifest> {
  const { node: root, prefix } = rootEntry(discovery.root);
  const relationshipPart = getXlsxRelationshipPartName(discovery.part);
  const relationshipsXml = await reader.readXml(relationshipPart, {
    required: true,
  });
  const relationships = parseXlsxRelationships(
    relationshipsXml,
    discovery.part,
    limits.maxRelationships,
  );
  const sheetsNode = record(child(root, prefix, 'sheets'));
  if (!sheetsNode)
    fail('Workbook sheets collection is missing', discovery.part);
  const sheetNodes = records(child(sheetsNode, prefix, 'sheet'));
  if (!sheetNodes) {
    fail('Workbook must contain at least one sheet', discovery.part);
  }
  if (sheetNodes.length > limits.maxWorksheets) {
    throw new XlsxResourceLimitError(
      'maxWorksheets',
      sheetNodes.length,
      limits.maxWorksheets,
      discovery.part,
    );
  }

  const names = new Set<string>();
  const sheetIds = new Set<number>();
  const sheetIdIndexes = new Map<number, number>();
  const sheetParts: string[] = [];
  const sheets: XlsxSheet[] = [];
  const relationshipBase = RELATIONSHIP_BASE[discovery.dialect];
  for (const [index, sheetNode] of sheetNodes.entries()) {
    const attrs = attributes(sheetNode);
    if (!validSheetName(attrs.name)) {
      fail('Workbook sheet has an invalid name', discovery.part);
    }
    const foldedName = attrs.name.toUpperCase();
    if (names.has(foldedName)) {
      fail('Workbook contains duplicate sheet names', discovery.part);
    }
    names.add(foldedName);

    const sheetId = positiveSheetId(attrs.sheetId);
    if (sheetId === undefined) {
      fail('Workbook sheet has an invalid sheetId', discovery.part);
    }
    if (sheetIds.has(sheetId)) {
      fail('Workbook contains duplicate sheetId values', discovery.part);
    }
    sheetIds.add(sheetId);
    sheetIdIndexes.set(sheetId, index);

    if (typeof attrs['r:id'] !== 'string' || attrs['r:id'].length === 0) {
      fail(
        'Workbook sheet has an invalid relationship reference',
        discovery.part,
      );
    }
    const relationship = relationships.get(attrs['r:id']);
    if (!relationship || relationship.mode !== 'internal') {
      fail(
        'Workbook sheet relationship is missing or external',
        discovery.part,
      );
    }

    const worksheetType = `${relationshipBase}/worksheet`;
    const chartSheetType = `${relationshipBase}/chartsheet`;
    const kind =
      relationship.type === worksheetType
        ? 'worksheet'
        : relationship.type === chartSheetType
          ? 'chart-sheet'
          : undefined;
    if (!kind) {
      fail(
        'Workbook sheet relationship has an unsupported type',
        discovery.part,
      );
    }
    const expectedContentType = expectedSheetContentType(kind);
    if (
      discovery.contentTypes.contentTypeFor(relationship.target) !==
      expectedContentType
    ) {
      fail(
        'Workbook sheet target has the wrong content type',
        relationship.target,
      );
    }

    const base = {
      index,
      name: attrs.name,
      payload: 'full-sheet' as const,
      state: parseSheetState(attrs.state, discovery.part),
    };
    if (kind === 'worksheet') {
      const worksheet: XlsxWorksheet = {
        ...base,
        columns: [],
        comments: [],
        conditionalFormattings: [],
        dataValidations: [],
        drawings: [],
        hyperlinks: [],
        kind,
        mergedRanges: [],
        protectedRanges: [],
        rows: [],
        tables: [],
        views: [],
      };
      sheets.push(worksheet);
    } else {
      const chartSheet: XlsxChartSheet = { ...base, kind };
      sheets.push(chartSheet);
    }
    sheetParts.push(relationship.target);
  }

  const definedNames = parseXlsxDefinedNames(
    child(root, prefix, 'definedNames'),
    prefix,
    discovery.part,
    sheets.length,
    limits,
  );
  const views = parseXlsxWorkbookViews(
    child(root, prefix, 'bookViews'),
    prefix,
    discovery.part,
    sheets,
  );
  const protection = parseXlsxWorkbookProtection(
    child(root, prefix, 'workbookProtection'),
    discovery.part,
  );
  const pivotCacheValue = child(root, prefix, 'pivotCaches');
  let pivotCacheNodes: XmlRecord[];
  if (pivotCacheValue === undefined) {
    pivotCacheNodes = [];
  } else {
    const pivotCacheContainer = record(pivotCacheValue);
    const nodes = pivotCacheContainer
      ? records(child(pivotCacheContainer, prefix, 'pivotCache'))
      : undefined;
    if (!pivotCacheContainer || !nodes) {
      fail('Workbook pivot-cache declarations are invalid', discovery.part);
    }
    pivotCacheNodes = nodes;
  }
  const pivotCacheIds = new Set<number>();
  const pivotCaches: XlsxPivotCacheDeclaration[] = [];
  for (const pivotCacheNode of pivotCacheNodes) {
    const attrs = attributes(pivotCacheNode);
    const cacheId = unsignedInteger(attrs.cacheId);
    if (cacheId === undefined) {
      fail('Workbook pivot cache ID is invalid', discovery.part);
    }
    if (pivotCacheIds.has(cacheId)) {
      fail('Workbook contains duplicate pivot cache IDs', discovery.part);
    }
    pivotCacheIds.add(cacheId);
    const relationshipId = attrs['r:id'];
    if (typeof relationshipId !== 'string' || relationshipId.length === 0) {
      fail(
        'Workbook pivot cache relationship reference is invalid',
        discovery.part,
      );
    }
    const relationship = relationships.get(relationshipId);
    if (
      !relationship ||
      relationship.mode !== 'internal' ||
      relationship.type !== `${relationshipBase}/pivotCacheDefinition`
    ) {
      fail('Workbook pivot cache relationship is invalid', discovery.part);
    }
    if (
      discovery.contentTypes.contentTypeFor(relationship.target) !==
      'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'
    ) {
      fail(
        'Workbook pivot cache target has the wrong content type',
        relationship.target,
      );
    }
    pivotCaches.push({ cacheId, target: relationship.target });
  }
  const manifest = {
    pivotCaches,
    properties: parseProperties(
      root,
      prefix,
      discovery.part,
      definedNames.definedNames,
      views,
      protection.protection,
    ),
    protectionTextCharacters: protection.textCharacters,
    sheetParts,
    sheets,
    workbookRelationships: relationships,
  } as XlsxWorkbookManifest;
  Object.defineProperty(manifest, 'sheetIdIndexes', {
    enumerable: false,
    value: sheetIdIndexes,
  });
  return manifest;
}
