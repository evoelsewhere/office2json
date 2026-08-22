import type { XmlLookupValue } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type { XlsxPartReader } from './part-reader';
import type { XlsxRelationship } from './relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import type { XlsxWorkbookDiscovery } from './workbook-discovery';

type XmlRecord = Record<string, unknown>;

interface FeatureBag {
  node: XmlRecord;
  type: string;
}

export interface XlsxFeaturePropertyBagRegistry {
  checkboxComplements: ReadonlySet<number>;
  part: string | null;
  records: number;
}

export const EMPTY_XLSX_FEATURE_PROPERTY_BAGS: XlsxFeaturePropertyBagRegistry =
  Object.freeze({
    checkboxComplements: new Set<number>(),
    part: null,
    records: 0,
  });

export function xlsxFeaturePropertyBagNamespace(): string {
  return 'http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag';
}

function fail(
  code:
    | 'invalid-document-structure'
    | 'invalid-document-value'
    | 'missing-required-part',
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

function attributes(node: XmlRecord): Record<string, string> {
  return (record(node.attrs) ?? {}) as Record<string, string>;
}

function localName(name: string): string {
  return name.slice(name.lastIndexOf(':') + 1);
}

function sourcePrefix(name: string): string {
  const parts = name.split(':', 2);
  if (parts.length === 1) return '';
  return parts[0]!.startsWith('ns_') ? parts[0]!.slice(3) : parts[0]!;
}

function namespaceFor(
  qualifiedName: string,
  node: unknown,
  inherited: Readonly<Record<string, string>>,
): string | undefined {
  const prefix = sourcePrefix(qualifiedName);
  const declaration = prefix ? `xmlns:${prefix}` : 'xmlns';
  const child = record(node);
  return (
    (child === undefined ? undefined : attributes(child)[declaration]) ??
    inherited[declaration]
  );
}

function children(
  node: XmlRecord,
  name: string,
  inherited: Readonly<Record<string, string>>,
  part: string,
): XmlRecord[] {
  const output: XmlRecord[] = [];
  for (const [qualifiedName, value] of Object.entries(node)) {
    if (localName(qualifiedName) !== name) continue;
    const values = Array.isArray(value) ? value : [value];
    for (const child of values) {
      const childRecord = record(child);
      if (!childRecord) {
        fail(
          'invalid-document-structure',
          'Feature property bag collection is invalid',
          part,
        );
      }
      if (
        namespaceFor(qualifiedName, childRecord, inherited) !==
        xlsxFeaturePropertyBagNamespace()
      ) {
        fail(
          'invalid-document-structure',
          'Feature property bag element has the wrong namespace',
          part,
        );
      }
      output.push(childRecord);
    }
  }
  return output;
}

function elementValues(
  node: XmlRecord,
  name: string,
  inherited: Readonly<Record<string, string>>,
  part: string,
): unknown[] {
  const output: unknown[] = [];
  for (const [qualifiedName, value] of Object.entries(node)) {
    if (localName(qualifiedName) !== name) continue;
    const values = Array.isArray(value) ? value : [value];
    for (const child of values) {
      if (
        namespaceFor(qualifiedName, child, inherited) !==
        xlsxFeaturePropertyBagNamespace()
      ) {
        fail(
          'invalid-document-structure',
          'Feature property bag element has the wrong namespace',
          part,
        );
      }
      output.push(child);
    }
  }
  return output;
}

function scalar(node: unknown, part: string): string {
  if (typeof node === 'string') return node;
  const nodeRecord = record(node);
  if (!nodeRecord) {
    fail(
      'invalid-document-value',
      'Feature property bag reference is invalid',
      part,
    );
  }
  const nested = Object.keys(nodeRecord).filter(
    (name) => name !== 'attrs' && name !== 'value',
  );
  if (nested.length !== 0 || typeof nodeRecord.value !== 'string') {
    fail(
      'invalid-document-value',
      'Feature property bag reference is invalid',
      part,
    );
  }
  return nodeRecord.value;
}

function unsignedInteger(value: string, part: string): number {
  if (!/^(?:0|[1-9]\d*)$/u.test(value)) {
    fail(
      'invalid-document-value',
      'Feature property bag reference is invalid',
      part,
    );
  }
  return Number(value);
}

function root(
  value: XmlLookupValue,
  part: string,
): {
  attrs: Record<string, string>;
  node: XmlRecord;
} {
  const entries = Object.entries(value).filter(
    ([name]) => localName(name) === 'FeaturePropertyBags',
  );
  if (entries.length !== 1) {
    fail(
      'invalid-document-structure',
      'Feature property bag root is missing or duplicated',
      part,
    );
  }
  const [qualifiedName, rawNode] = entries[0]!;
  const node = record(rawNode);
  if (
    !node ||
    namespaceFor(qualifiedName, node, attributes(node)) !==
      xlsxFeaturePropertyBagNamespace()
  ) {
    fail(
      'invalid-document-structure',
      'Feature property bag root has the wrong namespace',
      part,
    );
  }
  return { attrs: attributes(node), node };
}

function reference(
  node: XmlRecord,
  key: string | undefined,
  bags: readonly FeatureBag[],
  rootAttrs: Readonly<Record<string, string>>,
  part: string,
): number {
  const ids = elementValues(node, 'bagId', rootAttrs, part).filter((id) => {
    if (key === undefined) return true;
    const idRecord = record(id);
    return idRecord !== undefined && attributes(idRecord).k === key;
  });
  if (ids.length !== 1) {
    fail(
      'invalid-document-structure',
      'Feature property bag mapping is invalid',
      part,
    );
  }
  const index = unsignedInteger(scalar(ids[0], part), part);
  if (bags[index] === undefined) {
    fail(
      'invalid-document-value',
      'Feature property bag reference is invalid',
      part,
    );
  }
  return index;
}

export function parseXlsxFeaturePropertyBagPart(
  value: XmlLookupValue,
  part: string,
  limits: ResolvedXlsxResourceLimits,
): XlsxFeaturePropertyBagRegistry {
  const definition = root(value, part);
  const bags: FeatureBag[] = children(
    definition.node,
    'bag',
    definition.attrs,
    part,
  ).map((node) => {
    const type = attributes(node).type;
    if (!type) {
      fail(
        'invalid-document-value',
        'Feature property bag type is invalid',
        part,
      );
    }
    return { node, type };
  });
  let records = bags.length;
  for (const bag of bags) {
    const directIds = elementValues(bag.node, 'bagId', definition.attrs, part);
    const mappings = children(bag.node, 'a', definition.attrs, part);
    records += directIds.length + mappings.length;
    for (const mapping of mappings) {
      records += elementValues(mapping, 'bagId', definition.attrs, part).length;
    }
  }
  if (records > limits.maxStyles) {
    throw new XlsxResourceLimitError(
      'maxStyles',
      records,
      limits.maxStyles,
      part,
    );
  }
  const checkboxComplements = new Set<number>();
  const complementCollections = bags.filter(
    (bag) => bag.type === 'XFComplements',
  );
  if (complementCollections.length > 1) {
    fail(
      'invalid-document-structure',
      'Feature property bag complement collection is duplicated',
      part,
    );
  }
  for (const bag of complementCollections) {
    const mappings = children(bag.node, 'a', definition.attrs, part);
    for (const [index, mapping] of mappings.entries()) {
      if (attributes(mapping).k !== 'MappedFeaturePropertyBags') continue;
      const complementIndex = reference(
        mapping,
        undefined,
        bags,
        definition.attrs,
        part,
      );
      const complement = bags[complementIndex]!;
      if (complement.type !== 'XFComplement') continue;
      const controlsIndex = reference(
        complement.node,
        'XFControls',
        bags,
        definition.attrs,
        part,
      );
      const controls = bags[controlsIndex]!;
      if (controls.type !== 'XFControls') continue;
      const controlIndex = reference(
        controls.node,
        'CellControl',
        bags,
        definition.attrs,
        part,
      );
      if (bags[controlIndex]!.type === 'Checkbox') {
        checkboxComplements.add(index);
      }
    }
  }
  return Object.freeze({ checkboxComplements, part, records });
}

export async function loadXlsxFeaturePropertyBags(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
): Promise<XlsxFeaturePropertyBagRegistry> {
  const relationshipType =
    'http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag';
  const candidates = [...relationships.values()].filter(
    (relationship) => relationship.type === relationshipType,
  );
  if (candidates.length === 0) return EMPTY_XLSX_FEATURE_PROPERTY_BAGS;
  if (candidates.length !== 1) {
    fail(
      'invalid-document-structure',
      'Feature property bag relationship is duplicated',
      discovery.part,
    );
  }
  const relationship = candidates[0]!;
  if (relationship.mode !== 'internal') {
    fail(
      'invalid-document-structure',
      'Feature property bag relationship must be internal',
      discovery.part,
    );
  }
  if (
    discovery.contentTypes.contentTypeFor(relationship.target) !==
    'application/vnd.ms-excel.featurepropertybag+xml'
  ) {
    fail(
      'invalid-document-structure',
      'Feature property bag target has the wrong content type',
      relationship.target,
    );
  }
  return parseXlsxFeaturePropertyBagPart(
    await reader.readXml(relationship.target, { required: true }),
    relationship.target,
    limits,
  );
}
