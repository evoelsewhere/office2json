import type { XmlLookupValue } from '../../../common/xml/tree';

import { XlsxParseError } from '../errors';
import type {
  XlsxRange,
  XlsxTable,
  XlsxTableColumn,
  XlsxTableFormula,
  XlsxTableStyleInfo,
  XlsxTableTotalsFunction,
} from '../types';
import { parseXlsxRangeReference } from './cell-reference';
import { parseXlsxAutoFilter } from './auto-filter';
import { XlsxPartReader } from './part-reader';
import type { XlsxRelationship } from './relationships';
import type { ResolvedXlsxResourceLimits } from './resource-limits';
import { XlsxResourceLimitError } from './resource-limits';
import type { XlsxResolvedSheetSelection } from './selection';
import { isValidXlsxDefinedName } from './workbook-defined-names';
import {
  consumeXlsxWorksheetBudget,
  consumeXlsxWorksheetFormulaCharacters,
  type XlsxWorksheetBudget,
} from './worksheet';
import {
  XLSX_SPREADSHEET_NAMESPACES,
  type XlsxWorkbookDiscovery,
} from './workbook-discovery';

type XmlRecord = Record<string, unknown>;
type XlsxTableSelection = Exclude<
  XlsxResolvedSheetSelection,
  { kind: 'not-selected' }
>;

const TABLE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
const TABLE_RELATIONSHIP_TYPE = {
  strict: 'http://purl.oclc.org/ooxml/officeDocument/relationships/table',
  transitional:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
} as const;
const XML_REFERENCE_PATTERN =
  /&(?:amp|apos|gt|lt|quot|#(?:x[0-9A-Fa-f]+|[0-9]+));/gu;

export interface XlsxTableRegistry {
  count: number;
  ids: Set<number>;
  names: Set<string>;
  parts: Set<string>;
}

export function createXlsxTableRegistry(): XlsxTableRegistry {
  return { count: 0, ids: new Set(), names: new Set(), parts: new Set() };
}

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

function relationshipFailure(
  message: string,
  part: string,
  relationshipType: string,
): never {
  throw new XlsxParseError({
    code: 'invalid-relationship-target',
    message,
    part,
    relationshipType,
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

function child(node: XmlRecord, prefix: string, localName: string): unknown {
  return node[prefix ? `${prefix}:${localName}` : localName];
}

function rootEntry(
  value: XmlLookupValue,
  dialect: XlsxWorkbookDiscovery['dialect'],
  part: string,
): { node: XmlRecord; prefix: string } {
  const entries = Object.entries(value as unknown as XmlRecord);
  const [qualifiedName, sourceNode] = entries[0]!;
  const node = record(sourceNode);
  const [first, second] = qualifiedName.split(':') as [string, string?];
  const prefix = second === undefined ? '' : first;
  if (!node || (second ?? first) !== 'table') {
    structureFailure('Table root is missing', part);
  }
  const namespace = attributes(node)[prefix ? `xmlns:${prefix}` : 'xmlns'];
  if (namespace !== XLSX_SPREADSHEET_NAMESPACES[dialect]) {
    structureFailure('Table root has the wrong namespace', part);
  }
  return { node, prefix };
}

function unsignedInteger(
  value: unknown,
  part: string,
  message: string,
): number {
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > 0xffff_ffff) {
    valueFailure(message, part);
  }
  return parsed;
}

function positiveInteger(
  value: unknown,
  part: string,
  message: string,
): number {
  const parsed = unsignedInteger(value, part, message);
  if (parsed === 0) valueFailure(message, part);
  return parsed;
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

function textAttribute(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined;
}

function accountedTextAttribute(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): string | undefined {
  const parsed = textAttribute(value);
  if (parsed !== undefined) {
    consumeXlsxWorksheetBudget(
      budget,
      'textCharacters',
      parsed.length,
      'maxTextCharacters',
      limits,
      part,
    );
  }
  return parsed;
}

function boundedTableRange(
  value: unknown,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxRange {
  const source = typeof value === 'string' ? value : undefined;
  const parsed = parseXlsxRangeReference(source);
  if (!parsed || source?.includes('$')) {
    valueFailure('Table range is invalid', part);
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
  return parsed;
}

function formula(
  value: unknown,
  part: string,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  message: string,
): XlsxTableFormula | undefined {
  if (value === undefined) return undefined;
  if (Array.isArray(value)) structureFailure(message, part);
  const node = record(value);
  const expression =
    typeof value === 'string'
      ? decodeXmlEntities(value)
      : typeof node?.value === 'string'
        ? decodeXmlEntities(node.value)
        : undefined;
  if (expression === undefined) {
    valueFailure(message, part);
  }
  consumeXlsxWorksheetFormulaCharacters(budget, expression, limits, part);
  const attrs = node ? attributes(node) : {};
  return {
    array:
      optionalBoolean(
        attrs.array,
        part,
        'Table formula array flag is invalid',
      ) ?? false,
    expression,
  };
}

function differentialStyle(
  value: unknown,
  count: number,
  part: string,
  message: string,
): number | undefined {
  if (value === undefined) return undefined;
  const index = unsignedInteger(value, part, message);
  if (index >= count) valueFailure(message, part);
  return index;
}

function totalsFunction(value: unknown, part: string): XlsxTableTotalsFunction {
  if (value === undefined || value === 'none') return 'none';
  const functions: Record<string, XlsxTableTotalsFunction> = {
    average: 'average',
    count: 'count',
    countNums: 'count-numbers',
    custom: 'custom',
    max: 'maximum',
    min: 'minimum',
    stdDev: 'standard-deviation',
    sum: 'sum',
    var: 'variance',
  };
  const normalized = typeof value === 'string' ? functions[value] : undefined;
  if (!normalized) valueFailure('Table totals-row function is invalid', part);
  return normalized;
}

function tableColumn(
  value: XmlRecord,
  prefix: string,
  part: string,
  differentialStyleCount: number,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
): XlsxTableColumn {
  const attrs = attributes(value);
  const id = positiveInteger(attrs.id, part, 'Table column ID is invalid');
  const name = textAttribute(attrs.name);
  if (name === undefined || name.length === 0) {
    valueFailure('Table column name is invalid', part);
  }
  consumeXlsxWorksheetBudget(
    budget,
    'textCharacters',
    name.length,
    'maxTextCharacters',
    limits,
    part,
  );
  const totalsLabel = textAttribute(attrs.totalsRowLabel);
  if (totalsLabel !== undefined) {
    consumeXlsxWorksheetBudget(
      budget,
      'textCharacters',
      totalsLabel.length,
      'maxTextCharacters',
      limits,
      part,
    );
  }
  const uniqueName = accountedTextAttribute(
    attrs.uniqueName,
    budget,
    limits,
    part,
  );
  const headerCellStyle = accountedTextAttribute(
    attrs.headerRowCellStyle,
    budget,
    limits,
    part,
  );
  const dataCellStyle = accountedTextAttribute(
    attrs.dataCellStyle,
    budget,
    limits,
    part,
  );
  const totalsCellStyle = accountedTextAttribute(
    attrs.totalsRowCellStyle,
    budget,
    limits,
    part,
  );
  const queryTableFieldId =
    attrs.queryTableFieldId === undefined
      ? undefined
      : unsignedInteger(
          attrs.queryTableFieldId,
          part,
          'Table column query-field ID is invalid',
        );
  const totals = totalsFunction(attrs.totalsRowFunction, part);
  const totalsFormula = formula(
    child(value, prefix, 'totalsRowFormula'),
    part,
    budget,
    limits,
    'Table totals-row formula is invalid',
  );
  if (totals === 'custom' && totalsFormula === undefined) {
    valueFailure('Custom table totals formula is missing', part);
  }
  if (totalsFormula !== undefined && totals !== 'custom') {
    valueFailure('Table totals formula requires the custom function', part);
  }
  const calculatedFormula = formula(
    child(value, prefix, 'calculatedColumnFormula'),
    part,
    budget,
    limits,
    'Table calculated-column formula is invalid',
  );
  const dataDifferentialStyle = differentialStyle(
    attrs.dataDxfId,
    differentialStyleCount,
    part,
    'Table data differential-style reference is invalid',
  );
  const headerDifferentialStyle = differentialStyle(
    attrs.headerRowDxfId,
    differentialStyleCount,
    part,
    'Table header differential-style reference is invalid',
  );
  const totalsDifferentialStyle = differentialStyle(
    attrs.totalsRowDxfId,
    differentialStyleCount,
    part,
    'Table totals differential-style reference is invalid',
  );
  return {
    ...(calculatedFormula === undefined ? {} : { calculatedFormula }),
    ...(dataCellStyle === undefined ? {} : { dataCellStyle }),
    ...(dataDifferentialStyle === undefined ? {} : { dataDifferentialStyle }),
    ...(headerCellStyle === undefined ? {} : { headerCellStyle }),
    ...(headerDifferentialStyle === undefined
      ? {}
      : { headerDifferentialStyle }),
    id,
    name,
    ...(queryTableFieldId === undefined ? {} : { queryTableFieldId }),
    ...(totalsCellStyle === undefined ? {} : { totalsCellStyle }),
    ...(totalsDifferentialStyle === undefined
      ? {}
      : { totalsDifferentialStyle }),
    ...(totalsFormula === undefined ? {} : { totalsFormula }),
    ...(totalsLabel === undefined ? {} : { totalsLabel }),
    totalsFunction: totals,
    ...(uniqueName === undefined ? {} : { uniqueName }),
  };
}

function tableStyle(
  value: unknown,
  part: string,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
): XlsxTableStyleInfo | undefined {
  if (value === undefined) return undefined;
  const node = record(value);
  if (!node || Array.isArray(value)) {
    structureFailure('Table style info is invalid', part);
  }
  const attrs = attributes(node);
  const name = textAttribute(attrs.name);
  if (name !== undefined) {
    consumeXlsxWorksheetBudget(
      budget,
      'textCharacters',
      name.length,
      'maxTextCharacters',
      limits,
      part,
    );
  }
  return {
    ...(name === undefined ? {} : { name }),
    showColumnStripes:
      optionalBoolean(
        attrs.showColumnStripes,
        part,
        'Table column-stripe flag is invalid',
      ) ?? false,
    showFirstColumn:
      optionalBoolean(
        attrs.showFirstColumn,
        part,
        'Table first-column flag is invalid',
      ) ?? false,
    showLastColumn:
      optionalBoolean(
        attrs.showLastColumn,
        part,
        'Table last-column flag is invalid',
      ) ?? false,
    showRowStripes:
      optionalBoolean(
        attrs.showRowStripes,
        part,
        'Table row-stripe flag is invalid',
      ) ?? false,
  };
}

function selectionRelation(
  selection: XlsxTableSelection,
  range: XlsxRange,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxTable['selectionRelation'] | null {
  if (selection.kind === 'full-sheet') return 'full-sheet';
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
      selected.start.row <= range.end.row &&
      selected.end.row >= range.start.row &&
      selected.start.column <= range.end.column &&
      selected.end.column >= range.start.column
    ) {
      return 'intersects-selection';
    }
  }
  return null;
}

function tableType(value: unknown, part: string): XlsxTable['tableType'] {
  if (value === undefined || value === 'worksheet') return 'worksheet';
  if (value === 'xml') return 'xml';
  if (value === 'queryTable') return 'query-table';
  valueFailure('Table type is invalid', part);
}

async function parseTable(
  part: string,
  dialect: XlsxWorkbookDiscovery['dialect'],
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  registry: XlsxTableRegistry,
  differentialStyleCount: number,
  budget: XlsxWorksheetBudget,
  selection: XlsxTableSelection,
): Promise<XlsxTable | null> {
  const xml = await reader.readXml(part, { required: true });
  const { node, prefix } = rootEntry(xml, dialect, part);
  const attrs = attributes(node);
  const id = positiveInteger(attrs.id, part, 'Table ID is invalid');
  const name = textAttribute(attrs.name);
  const displayName = textAttribute(attrs.displayName);
  if (!isValidXlsxDefinedName(name) || !isValidXlsxDefinedName(displayName)) {
    valueFailure('Table name is invalid', part);
  }
  const range = boundedTableRange(attrs.ref, limits, part);
  consumeXlsxWorksheetBudget(
    budget,
    'rangeAreas',
    1,
    'maxRangeAreas',
    limits,
    part,
  );
  const comment = accountedTextAttribute(attrs.comment, budget, limits, part);
  const headerCellStyle = accountedTextAttribute(
    attrs.headerRowCellStyle,
    budget,
    limits,
    part,
  );
  const dataCellStyle = accountedTextAttribute(
    attrs.dataCellStyle,
    budget,
    limits,
    part,
  );
  const totalsCellStyle = accountedTextAttribute(
    attrs.totalsRowCellStyle,
    budget,
    limits,
    part,
  );
  const connectionId =
    attrs.connectionId === undefined
      ? undefined
      : unsignedInteger(
          attrs.connectionId,
          part,
          'Table connection ID is invalid',
        );
  const dataDifferentialStyle = differentialStyle(
    attrs.dataDxfId,
    differentialStyleCount,
    part,
    'Table data differential-style reference is invalid',
  );
  const headerDifferentialStyle = differentialStyle(
    attrs.headerRowDxfId,
    differentialStyleCount,
    part,
    'Table header differential-style reference is invalid',
  );
  const totalsDifferentialStyle = differentialStyle(
    attrs.totalsRowDxfId,
    differentialStyleCount,
    part,
    'Table totals differential-style reference is invalid',
  );
  const headerRowBorderDifferentialStyle = differentialStyle(
    attrs.headerRowBorderDxfId,
    differentialStyleCount,
    part,
    'Table header-border differential-style reference is invalid',
  );
  const tableBorderDifferentialStyle = differentialStyle(
    attrs.tableBorderDxfId,
    differentialStyleCount,
    part,
    'Table border differential-style reference is invalid',
  );
  const totalsRowBorderDifferentialStyle = differentialStyle(
    attrs.totalsRowBorderDxfId,
    differentialStyleCount,
    part,
    'Table totals-border differential-style reference is invalid',
  );
  consumeXlsxWorksheetBudget(
    budget,
    'textCharacters',
    name.length + displayName.length,
    'maxTextCharacters',
    limits,
    part,
  );
  const foldedNames = [name.toUpperCase(), displayName.toUpperCase()];
  if (
    registry.ids.has(id) ||
    foldedNames.some((candidate) => registry.names.has(candidate))
  ) {
    valueFailure('Workbook contains duplicate table identities', part);
  }
  registry.ids.add(id);
  for (const folded of foldedNames) registry.names.add(folded);

  const columnsNode = record(child(node, prefix, 'tableColumns'));
  if (!columnsNode)
    structureFailure('Table columns collection is missing', part);
  const columnNodes = records(child(columnsNode, prefix, 'tableColumn'));
  if (!columnNodes) {
    structureFailure('Table columns collection is invalid', part);
  }
  const declaredColumnCount = unsignedInteger(
    attributes(columnsNode).count,
    part,
    'Table column count is invalid',
  );
  if (declaredColumnCount !== columnNodes.length) {
    structureFailure('Table column count does not match', part);
  }
  const width = range.end.column - range.start.column + 1;
  if (columnNodes.length !== width) {
    valueFailure('Table column count does not match its range', part);
  }
  const columns = columnNodes.map((column) =>
    tableColumn(column, prefix, part, differentialStyleCount, budget, limits),
  );
  if (new Set(columns.map((column) => column.id)).size !== columns.length) {
    valueFailure('Table contains duplicate column IDs', part);
  }
  const foldedColumns = columns.map((column) => column.name.toUpperCase());
  if (new Set(foldedColumns).size !== foldedColumns.length) {
    valueFailure('Table contains duplicate column names', part);
  }
  const headerRowCount =
    attrs.headerRowCount === undefined
      ? 1
      : unsignedInteger(
          attrs.headerRowCount,
          part,
          'Table header-row count is invalid',
        );
  const totalsRowCount =
    attrs.totalsRowCount === undefined
      ? 0
      : unsignedInteger(
          attrs.totalsRowCount,
          part,
          'Table totals-row count is invalid',
        );
  if (headerRowCount > 1 || totalsRowCount > 1) {
    valueFailure('Table header or totals row count is invalid', part);
  }
  if (range.end.row - range.start.row + 1 < headerRowCount + totalsRowCount) {
    valueFailure(
      'Table range is too short for its header and totals rows',
      part,
    );
  }
  const relation = selectionRelation(selection, range, budget, limits, part);
  const autoFilterSource = child(node, prefix, 'autoFilter');
  if (autoFilterSource !== undefined) {
    const autoFilterNode = record(autoFilterSource);
    const autoFilterRange = autoFilterNode
      ? parseXlsxRangeReference(attributes(autoFilterNode).ref)
      : undefined;
    if (
      !autoFilterNode ||
      !autoFilterRange ||
      autoFilterRange.start.row < range.start.row ||
      autoFilterRange.end.row > range.end.row ||
      autoFilterRange.start.column < range.start.column ||
      autoFilterRange.end.column > range.end.column
    ) {
      valueFailure('Table auto-filter range is outside the table', part);
    }
  }
  const autoFilter = parseXlsxAutoFilter(
    autoFilterSource,
    prefix,
    differentialStyleCount,
    selection,
    budget,
    limits,
    part,
  );
  const style = tableStyle(
    child(node, prefix, 'tableStyleInfo'),
    part,
    budget,
    limits,
  );
  const table: XlsxTable = {
    ...(autoFilter === undefined ? {} : { autoFilter }),
    columns,
    ...(comment === undefined ? {} : { comment }),
    ...(connectionId === undefined ? {} : { connectionId }),
    ...(dataCellStyle === undefined ? {} : { dataCellStyle }),
    ...(dataDifferentialStyle === undefined ? {} : { dataDifferentialStyle }),
    displayName,
    ...(headerCellStyle === undefined ? {} : { headerCellStyle }),
    ...(headerDifferentialStyle === undefined
      ? {}
      : { headerDifferentialStyle }),
    headerRow: headerRowCount === 1,
    ...(headerRowBorderDifferentialStyle === undefined
      ? {}
      : { headerRowBorderDifferentialStyle }),
    id,
    insertRow:
      optionalBoolean(
        attrs.insertRow,
        part,
        'Table insert-row flag is invalid',
      ) ?? false,
    insertRowShift:
      optionalBoolean(
        attrs.insertRowShift,
        part,
        'Table insert-row-shift flag is invalid',
      ) ?? false,
    name,
    published:
      optionalBoolean(
        attrs.published,
        part,
        'Table published flag is invalid',
      ) ?? false,
    range,
    selectionRelation: relation ?? 'full-sheet',
    ...(style === undefined ? {} : { style }),
    ...(tableBorderDifferentialStyle === undefined
      ? {}
      : { tableBorderDifferentialStyle }),
    tableType: tableType(attrs.tableType, part),
    ...(totalsCellStyle === undefined ? {} : { totalsCellStyle }),
    ...(totalsDifferentialStyle === undefined
      ? {}
      : { totalsDifferentialStyle }),
    totalsRow: totalsRowCount === 1,
    ...(totalsRowBorderDifferentialStyle === undefined
      ? {}
      : { totalsRowBorderDifferentialStyle }),
    totalsRowShown:
      optionalBoolean(
        attrs.totalsRowShown,
        part,
        'Table totals-row shown flag is invalid',
      ) ?? true,
  };
  return relation === null ? null : table;
}

export async function loadXlsxTables(
  relationshipIds: readonly string[],
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: Pick<XlsxWorkbookDiscovery, 'contentTypes' | 'dialect'>,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  registry: XlsxTableRegistry,
  differentialStyleCount: number,
  budget: XlsxWorksheetBudget,
  selection: XlsxTableSelection,
  worksheetPart: string,
): Promise<XlsxTable[]> {
  const tables: XlsxTable[] = [];
  const expectedType = TABLE_RELATIONSHIP_TYPE[discovery.dialect];
  for (const relationshipId of relationshipIds) {
    const relationship = relationships.get(relationshipId);
    if (!relationship || relationship.type !== expectedType) {
      structureFailure(
        'Worksheet table relationship is missing',
        worksheetPart,
      );
    }
    if (relationship.mode !== 'internal') {
      relationshipFailure(
        'Worksheet table relationship must be internal',
        worksheetPart,
        relationship.type,
      );
    }
    if (
      discovery.contentTypes.contentTypeFor(relationship.target) !==
      TABLE_CONTENT_TYPE
    ) {
      structureFailure(
        'Worksheet table target has the wrong content type',
        relationship.target,
      );
    }
    if (registry.parts.has(relationship.target)) {
      structureFailure(
        'Workbook references a table part more than once',
        worksheetPart,
      );
    }
    registry.parts.add(relationship.target);
    registry.count += 1;
    if (registry.count > limits.maxTables) {
      throw new XlsxResourceLimitError(
        'maxTables',
        registry.count,
        limits.maxTables,
        relationship.target,
      );
    }
    const table = await parseTable(
      relationship.target,
      discovery.dialect,
      reader,
      limits,
      registry,
      differentialStyleCount,
      budget,
      selection,
    );
    if (table) tables.push(table);
  }
  return tables;
}
