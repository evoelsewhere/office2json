import type { XmlLookupValue } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type { XlsxCellMetadataEntry } from '../types';
import { XlsxPartReader } from './part-reader';
import type { XlsxRelationship } from './relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import type { XlsxWorkbookDiscovery } from './workbook-discovery';
import {
  cloneXlsxRichValueForOutput,
  EMPTY_XLSX_RICH_VALUES,
  type XlsxRichValueRegistry,
} from './rich-value';

type XmlRecord = Record<string, unknown>;

type RegistryEntry = XlsxCellMetadataEntry | { kind: 'unsupported' };

export interface XlsxCellMetadataBudget {
  records: number;
}

export interface XlsxCellMetadataRegistry {
  cellBlocks: readonly (readonly RegistryEntry[])[];
  part: string | null;
  valueBlocks: readonly (readonly RegistryEntry[])[];
}

export interface XlsxResolvedCellMetadata {
  entries: XlsxCellMetadataEntry[];
  unsupported: boolean;
}

export const EMPTY_XLSX_CELL_METADATA: XlsxCellMetadataRegistry = Object.freeze(
  {
    cellBlocks: Object.freeze([]),
    part: null,
    valueBlocks: Object.freeze([]),
  },
);

function spreadsheetNamespace(
  dialect: XlsxWorkbookDiscovery['dialect'],
): string {
  return dialect === 'strict'
    ? 'http://purl.oclc.org/ooxml/spreadsheetml/main'
    : 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
}

function relationshipType(dialect: XlsxWorkbookDiscovery['dialect']): string {
  const base =
    dialect === 'strict'
      ? 'http://purl.oclc.org/ooxml/officeDocument/relationships'
      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return `${base}/sheetMetadata`;
}

function dynamicArrayNamespace(): string {
  return 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
}

function richDataNamespace(): string {
  return 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';
}

function fail(
  code:
    | 'invalid-document-structure'
    | 'invalid-document-value'
    | 'missing-required-part',
  message: string,
  part: string,
  cell?: string,
): never {
  throw new XlsxParseError({
    ...(cell === undefined ? {} : { cell }),
    code,
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

function sourcePrefix(name: string): string {
  const parts = name.split(':', 2);
  if (parts.length === 1) return '';
  const prefix = parts[0]!;
  return prefix.startsWith('ns_') ? prefix.slice(3) : prefix;
}

function namespaceFor(
  qualifiedName: string,
  node: XmlRecord,
  inherited: Readonly<Record<string, string>>,
): string | undefined {
  const prefix = sourcePrefix(qualifiedName);
  const declaration = prefix ? `xmlns:${prefix}` : 'xmlns';
  return attributes(node)[declaration] ?? inherited[declaration];
}

interface RootResult {
  attrs: Record<string, string>;
  node: XmlRecord;
}

function root(
  value: XmlLookupValue,
  namespace: string,
  part: string,
): RootResult {
  const entries = Object.entries(value).filter(
    ([name]) => localName(name) === 'metadata',
  );
  if (entries.length !== 1) {
    fail(
      'invalid-document-structure',
      'Cell metadata root is missing or duplicated',
      part,
    );
  }
  const [qualifiedName, rawNode] = entries[0]!;
  const node = record(rawNode);
  if (
    !node ||
    namespaceFor(qualifiedName, node, attributes(node)) !== namespace
  ) {
    fail(
      'invalid-document-structure',
      'Cell metadata root has the wrong namespace',
      part,
    );
  }
  return { attrs: attributes(node), node };
}

function children(
  node: XmlRecord,
  name: string,
  namespace: string,
  inherited: Readonly<Record<string, string>>,
  part: string,
): XmlRecord[] {
  const output: XmlRecord[] = [];
  for (const [qualifiedName, value] of Object.entries(node)) {
    if (localName(qualifiedName) !== name) continue;
    const values = records(value);
    if (!values) {
      fail(
        'invalid-document-structure',
        'Cell metadata collection is invalid',
        part,
      );
    }
    for (const child of values) {
      if (namespaceFor(qualifiedName, child, inherited) !== namespace) {
        fail(
          'invalid-document-structure',
          'Cell metadata element has the wrong namespace',
          part,
        );
      }
      output.push(child);
    }
  }
  return output;
}

function onlyChild(
  node: XmlRecord,
  name: string,
  namespace: string,
  inherited: Readonly<Record<string, string>>,
  part: string,
): XmlRecord | undefined {
  const values = children(node, name, namespace, inherited, part);
  if (values.length > 1) {
    fail(
      'invalid-document-structure',
      `Cell metadata ${name} is duplicated`,
      part,
    );
  }
  return values[0];
}

function descendants(
  node: XmlRecord,
  name: string,
): Array<[string, XmlRecord]> {
  const output: Array<[string, XmlRecord]> = [];
  for (const [qualifiedName, value] of Object.entries(node)) {
    const values = records(value);
    if (!values) continue;
    for (const child of values) {
      if (localName(qualifiedName) === name)
        output.push([qualifiedName, child]);
      output.push(...descendants(child, name));
    }
  }
  return output;
}

export function parseXlsxCellMetadataUnsignedInteger(
  value: unknown,
  message: string,
  part: string,
): number {
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (parsed > 0xffff_ffff) {
    fail('invalid-document-value', message, part);
  }
  return parsed;
}

function boolean(
  value: string | undefined,
  message: string,
  part: string,
): boolean {
  if (value === undefined || value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  fail('invalid-document-value', message, part);
}

function assertCount(
  value: string | undefined,
  actual: number,
  message: string,
  part: string,
): void {
  if (
    value !== undefined &&
    parseXlsxCellMetadataUnsignedInteger(value, message, part) !== actual
  ) {
    fail('invalid-document-structure', message, part);
  }
}

function consumeMetadataRecord(
  budget: XlsxCellMetadataBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
  amount = 1,
): void {
  const actual = budget.records + amount;
  if (!Number.isSafeInteger(actual) || actual > limits.maxMetadataRecords) {
    throw new XlsxResourceLimitError(
      'maxMetadataRecords',
      actual,
      limits.maxMetadataRecords,
      part,
    );
  }
  budget.records = actual;
}

interface MetadataTypeDefinition {
  name: string;
}

function metadataTypes(
  definition: RootResult,
  namespace: string,
  part: string,
): MetadataTypeDefinition[] {
  const container = onlyChild(
    definition.node,
    'metadataTypes',
    namespace,
    definition.attrs,
    part,
  );
  if (!container) return [];
  const nodes = children(
    container,
    'metadataType',
    namespace,
    definition.attrs,
    part,
  );
  if (nodes.length === 0) {
    fail(
      'invalid-document-structure',
      'Cell metadata type collection is empty',
      part,
    );
  }
  assertCount(
    attributes(container).count,
    nodes.length,
    'Cell metadata type count does not match',
    part,
  );
  const names = new Set<string>();
  return nodes.map((node) => {
    const attrs = attributes(node);
    if (!attrs.name || names.has(attrs.name)) {
      fail(
        'invalid-document-value',
        'Cell metadata type name is invalid or duplicated',
        part,
      );
    }
    names.add(attrs.name);
    parseXlsxCellMetadataUnsignedInteger(
      attrs.minSupportedVersion,
      'Cell metadata minimum version is invalid',
      part,
    );
    for (const [name, value] of Object.entries(attrs)) {
      if (name === 'name' || name === 'minSupportedVersion') continue;
      boolean(value, 'Cell metadata type flag is invalid', part);
    }
    return { name: attrs.name };
  });
}

function knownFutureBlock(
  typeName: string,
  block: XmlRecord,
  rootAttrs: Readonly<Record<string, string>>,
  part: string,
  richValues: XlsxRichValueRegistry,
): RegistryEntry {
  if (typeName === 'XLDAPR') {
    const values = descendants(block, 'dynamicArrayProperties');
    if (values.length !== 1) {
      fail(
        'invalid-document-structure',
        'Dynamic-array metadata block is invalid',
        part,
      );
    }
    const [qualifiedName, node] = values[0]!;
    if (
      namespaceFor(qualifiedName, node, rootAttrs) !== dynamicArrayNamespace()
    ) {
      fail(
        'invalid-document-structure',
        'Dynamic-array metadata has the wrong namespace',
        part,
      );
    }
    const attrs = attributes(node);
    return {
      collapsed: boolean(
        attrs.fCollapsed,
        'Dynamic-array collapsed flag is invalid',
        part,
      ),
      dynamic: boolean(
        attrs.fDynamic,
        'Dynamic-array dynamic flag is invalid',
        part,
      ),
      kind: 'dynamic-array',
    };
  }
  if (typeName === 'XLRICHVALUE') {
    const values = descendants(block, 'rvb');
    if (values.length !== 1) {
      fail(
        'invalid-document-structure',
        'Rich-value metadata block is invalid',
        part,
      );
    }
    const [qualifiedName, node] = values[0]!;
    if (namespaceFor(qualifiedName, node, rootAttrs) !== richDataNamespace()) {
      fail(
        'invalid-document-structure',
        'Rich-value metadata has the wrong namespace',
        part,
      );
    }
    const valueIndex = parseXlsxCellMetadataUnsignedInteger(
      attributes(node).i,
      'Rich-value metadata index is invalid',
      part,
    );
    if (
      richValues.part !== null &&
      richValues.values[valueIndex] === undefined
    ) {
      fail(
        'invalid-document-value',
        'Rich-value metadata index is invalid',
        part,
      );
    }
    return {
      ...(richValues.values[valueIndex] === undefined
        ? {}
        : { data: richValues.values[valueIndex] }),
      kind: 'rich-value',
      valueIndex,
    };
  }
  return { kind: 'unsupported' };
}

function futureMetadata(
  definition: RootResult,
  namespace: string,
  part: string,
  richValues: XlsxRichValueRegistry,
): ReadonlyMap<string, readonly RegistryEntry[]> {
  const output = new Map<string, readonly RegistryEntry[]>();
  for (const container of children(
    definition.node,
    'futureMetadata',
    namespace,
    definition.attrs,
    part,
  )) {
    const attrs = attributes(container);
    const name = attrs.name;
    if (!name || output.has(name)) {
      fail(
        'invalid-document-value',
        'Future metadata name is invalid or duplicated',
        part,
      );
    }
    const blocks = children(container, 'bk', namespace, definition.attrs, part);
    assertCount(
      attrs.count,
      blocks.length,
      'Future metadata block count does not match',
      part,
    );
    output.set(
      name,
      blocks.map((block) =>
        knownFutureBlock(name, block, definition.attrs, part, richValues),
      ),
    );
  }
  return output;
}

function metadataBlocks(
  definition: RootResult,
  containerName: 'cellMetadata' | 'valueMetadata',
  types: readonly MetadataTypeDefinition[],
  future: ReadonlyMap<string, readonly RegistryEntry[]>,
  namespace: string,
  limits: ResolvedXlsxResourceLimits,
  budget: XlsxCellMetadataBudget,
  part: string,
): readonly (readonly RegistryEntry[])[] {
  const container = onlyChild(
    definition.node,
    containerName,
    namespace,
    definition.attrs,
    part,
  );
  if (!container) return [];
  const blocks = children(container, 'bk', namespace, definition.attrs, part);
  if (blocks.length === 0) {
    fail(
      'invalid-document-structure',
      `Cell metadata ${containerName} is empty`,
      part,
    );
  }
  assertCount(
    attributes(container).count,
    blocks.length,
    `Cell metadata ${containerName} count does not match`,
    part,
  );
  return blocks.map((block) => {
    const records = children(block, 'rc', namespace, definition.attrs, part);
    if (records.length === 0) {
      fail('invalid-document-structure', 'Cell metadata block is empty', part);
    }
    return records.map((metadataRecord) => {
      consumeMetadataRecord(budget, limits, part);
      const attrs = attributes(metadataRecord);
      const typeIndex = parseXlsxCellMetadataUnsignedInteger(
        attrs.t,
        'Cell metadata type index is invalid',
        part,
      );
      const valueIndex = parseXlsxCellMetadataUnsignedInteger(
        attrs.v,
        'Cell metadata value index is invalid',
        part,
      );
      if (typeIndex < 1 || typeIndex > types.length) {
        fail(
          'invalid-document-value',
          'Cell metadata type reference is invalid',
          part,
        );
      }
      const values = future.get(types[typeIndex - 1]!.name);
      if (!values || valueIndex >= values.length) {
        fail(
          'invalid-document-value',
          'Cell metadata value reference is invalid',
          part,
        );
      }
      return values[valueIndex]!;
    });
  });
}

function metadataTarget(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
): string | undefined {
  const type = relationshipType(discovery.dialect);
  const matches = [...relationships.values()].filter(
    (relationship) => relationship.type === type,
  );
  if (matches.length > 1) {
    fail(
      'invalid-document-structure',
      'Workbook cell-metadata relationship is duplicated',
      discovery.part,
    );
  }
  const relationship = matches[0];
  if (!relationship) return undefined;
  if (relationship.mode !== 'internal') {
    fail(
      'invalid-document-structure',
      'Workbook cell-metadata relationship must be internal',
      discovery.part,
    );
  }
  if (
    discovery.contentTypes.contentTypeFor(relationship.target) !==
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'
  ) {
    fail(
      'invalid-document-structure',
      'Cell metadata target has the wrong content type',
      relationship.target,
    );
  }
  return relationship.target;
}

export async function loadXlsxCellMetadata(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  budget: XlsxCellMetadataBudget,
  richValues: XlsxRichValueRegistry = EMPTY_XLSX_RICH_VALUES,
): Promise<XlsxCellMetadataRegistry> {
  const part = metadataTarget(relationships, discovery);
  if (!part) return EMPTY_XLSX_CELL_METADATA;
  return parseXlsxCellMetadataPart(
    await reader.readXml(part, { required: true }),
    discovery.dialect,
    part,
    limits,
    budget,
    richValues,
  );
}

export function parseXlsxCellMetadataPart(
  value: XmlLookupValue,
  dialect: XlsxWorkbookDiscovery['dialect'],
  part: string,
  limits: ResolvedXlsxResourceLimits,
  budget: XlsxCellMetadataBudget,
  richValues: XlsxRichValueRegistry = EMPTY_XLSX_RICH_VALUES,
): XlsxCellMetadataRegistry {
  const namespace = spreadsheetNamespace(dialect);
  const definition = root(value, namespace, part);
  const types = metadataTypes(definition, namespace, part);
  const future = futureMetadata(definition, namespace, part, richValues);
  return {
    cellBlocks: metadataBlocks(
      definition,
      'cellMetadata',
      types,
      future,
      namespace,
      limits,
      budget,
      part,
    ),
    part,
    valueBlocks: metadataBlocks(
      definition,
      'valueMetadata',
      types,
      future,
      namespace,
      limits,
      budget,
      part,
    ),
  };
}

export function resolveXlsxCellMetadata(
  registry: XlsxCellMetadataRegistry,
  source: 'cell' | 'value',
  index: number | undefined,
  budget: XlsxCellMetadataBudget,
  textBudget: { textCharacters: number },
  limits: ResolvedXlsxResourceLimits,
  worksheetPart: string,
  cell: string,
): XlsxResolvedCellMetadata | undefined {
  if (index === undefined) return undefined;
  if (registry.part === null) {
    fail(
      'missing-required-part',
      'Worksheet cell metadata part is missing',
      worksheetPart,
      cell,
    );
  }
  const blocks = source === 'cell' ? registry.cellBlocks : registry.valueBlocks;
  if (index < 1 || index > blocks.length) {
    fail(
      'invalid-document-value',
      'Worksheet cell metadata reference is invalid',
      worksheetPart,
      cell,
    );
  }
  const block = blocks[index - 1]!;
  consumeMetadataRecord(budget, limits, worksheetPart, block.length);
  return {
    entries: block
      .filter(
        (entry): entry is XlsxCellMetadataEntry => entry.kind !== 'unsupported',
      )
      .map((entry) =>
        entry.kind === 'rich-value' && entry.data !== undefined
          ? {
              ...entry,
              data: cloneXlsxRichValueForOutput(
                entry.data,
                budget,
                textBudget,
                limits,
                worksheetPart,
              ),
            }
          : { ...entry },
      ),
    unsupported: block.some((entry) => entry.kind === 'unsupported'),
  };
}
