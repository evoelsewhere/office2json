import type { XmlLookupValue } from '../../../common/xml/tree';
import { decodeXmlEntities } from '../../../common/text/html';
import { XlsxParseError } from '../errors';
import type {
  XlsxApplicationDocumentProperties,
  XlsxCoreDocumentProperties,
  XlsxCustomDocumentProperty,
  XlsxCustomDocumentPropertyValue,
  XlsxDocumentProperties,
} from '../types';
import { XlsxPartReader } from './part-reader';
import { parseXlsxPivotDateTime } from './pivot';
import { parseXlsxRelationships, type XlsxRelationship } from './relationships';
import type { ResolvedXlsxResourceLimits } from './resource-limits';
import type { XlsxWorkbookDiscovery } from './workbook-discovery';
import {
  consumeXlsxWorksheetBudget,
  type XlsxWorksheetBudget,
} from './worksheet';

type XmlRecord = Record<string, unknown>;

function coreNamespace(): string {
  return 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
}

function dcNamespace(): string {
  return 'http://purl.org/dc/elements/1.1/';
}

function dctermsNamespace(): string {
  return 'http://purl.org/dc/terms/';
}

function xsiNamespace(): string {
  return 'http://www.w3.org/2001/XMLSchema-instance';
}

function applicationNamespaces(): ReadonlySet<string> {
  return new Set([
    'http://purl.oclc.org/ooxml/officeDocument/extendedProperties',
    'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  ]);
}

function customNamespaces(): ReadonlySet<string> {
  return new Set([
    'http://purl.oclc.org/ooxml/officeDocument/customProperties',
    'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
  ]);
}

function valueTypeNamespaces(): ReadonlySet<string> {
  return new Set([
    'http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes',
    'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
  ]);
}

function coreRelationshipTypes(): ReadonlySet<string> {
  return new Set([
    'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  ]);
}

function applicationRelationshipTypes(): ReadonlySet<string> {
  return new Set([
    'http://purl.oclc.org/ooxml/officeDocument/relationships/extended-properties',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  ]);
}

function customRelationshipTypes(): ReadonlySet<string> {
  return new Set([
    'http://purl.oclc.org/ooxml/officeDocument/relationships/custom-properties',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
  ]);
}

function coreContentType(): string {
  return 'application/vnd.openxmlformats-package.core-properties+xml';
}

function applicationContentType(): string {
  return 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
}

function customContentType(): string {
  return 'application/vnd.openxmlformats-officedocument.custom-properties+xml';
}

function fail(
  code:
    | 'invalid-document-structure'
    | 'invalid-document-value'
    | 'unsupported-feature',
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
  const prefix = parts[0]!;
  return prefix.startsWith('ns_') ? prefix.slice(3) : prefix;
}

function namespaceFor(
  qualifiedName: string,
  node: unknown,
  inherited: Readonly<Record<string, string>>,
): string | undefined {
  const prefix = sourcePrefix(qualifiedName);
  const declaration = prefix ? `xmlns:${prefix}` : 'xmlns';
  return attributes(record(node) ?? {})[declaration] ?? inherited[declaration];
}

function hasNamespace(
  namespaces: ReadonlySet<string>,
  qualifiedName: string,
  node: unknown,
  inherited: Readonly<Record<string, string>>,
): boolean {
  return (namespaces as ReadonlySet<unknown>).has(
    namespaceFor(qualifiedName, node, inherited),
  );
}

interface RootResult {
  attrs: Record<string, string>;
  node: XmlRecord;
}

function root(
  value: XmlLookupValue,
  expectedName: string,
  namespaces: ReadonlySet<string>,
  part: string,
): RootResult {
  const entries = Object.entries(value).filter(
    ([name]) => localName(name) === expectedName,
  );
  if (entries.length !== 1) {
    fail(
      'invalid-document-structure',
      `XLSX ${expectedName} properties root is missing or duplicated`,
      part,
    );
  }
  const [qualifiedName, rawNode] = entries[0]!;
  const node = record(rawNode);
  if (!node) {
    fail(
      'invalid-document-structure',
      'Document properties root is invalid',
      part,
    );
  }
  const attrs = attributes(node);
  if (node.value !== undefined) {
    fail(
      'invalid-document-structure',
      'Document properties root contains text',
      part,
    );
  }
  if (!hasNamespace(namespaces, qualifiedName, node, attrs)) {
    fail(
      'invalid-document-structure',
      'Document properties root has the wrong namespace',
      part,
    );
  }
  return { attrs, node };
}

function directChildren(node: XmlRecord): Array<[string, unknown]> {
  const output: Array<[string, unknown]> = [];
  for (const [name, value] of Object.entries(node)) {
    if (name === 'attrs' || name === 'value') continue;
    const children = Array.isArray(value) ? value : [value];
    for (const child of children) output.push([name, child]);
  }
  return output;
}

function requiredRecord(
  value: unknown,
  message: string,
  part: string,
): XmlRecord {
  const node = record(value);
  if (!node) fail('invalid-document-structure', message, part);
  return node;
}

function scalar(node: unknown, message: string, part: string): string {
  if (typeof node === 'string') return decodeXmlEntities(node);
  if (node === undefined) return '';
  const definition = requiredRecord(node, message, part);
  if (directChildren(definition).length !== 0) {
    fail('invalid-document-structure', message, part);
  }
  const value = definition.value;
  if (value === undefined) return '';
  if (typeof value !== 'string') fail('invalid-document-value', message, part);
  return decodeXmlEntities(value);
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

export function parseXlsxDocumentUnsignedInteger(
  value: unknown,
  message: string,
  part: string,
): number {
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > 0xffff_ffff) {
    fail('invalid-document-value', message, part);
  }
  return parsed;
}

export function parseXlsxDocumentBoolean(
  value: string,
  message: string,
  part: string,
): boolean {
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  fail('invalid-document-value', message, part);
}

function propertyTarget(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  relationshipTypes: ReadonlySet<string>,
  contentType: string,
  discovery: XlsxWorkbookDiscovery,
): string | undefined {
  const matches = [...relationships.values()].filter((relationship) =>
    relationshipTypes.has(relationship.type),
  );
  if (matches.length > 1) {
    fail(
      'invalid-document-structure',
      'Package document-property relationship is duplicated',
      '_rels/.rels',
    );
  }
  const relationship = matches[0];
  if (!relationship) return undefined;
  if (relationship.mode !== 'internal') {
    fail(
      'invalid-document-structure',
      'Package document-property relationship must be internal',
      '_rels/.rels',
    );
  }
  if (
    discovery.contentTypes.contentTypeFor(relationship.target) !== contentType
  ) {
    fail(
      'invalid-document-structure',
      'Document-property target has the wrong content type',
      relationship.target,
    );
  }
  return relationship.target;
}

const CORE_FIELDS = {
  category: ['category', coreNamespace],
  contentStatus: ['contentStatus', coreNamespace],
  contentType: ['contentType', coreNamespace],
  created: ['created', dctermsNamespace],
  creator: ['creator', dcNamespace],
  description: ['description', dcNamespace],
  identifier: ['identifier', dcNamespace],
  keywords: ['keywords', coreNamespace],
  language: ['language', dcNamespace],
  lastModifiedBy: ['lastModifiedBy', coreNamespace],
  lastPrinted: ['lastPrinted', coreNamespace],
  modified: ['modified', dctermsNamespace],
  revision: ['revision', coreNamespace],
  subject: ['subject', dcNamespace],
  title: ['title', dcNamespace],
  version: ['version', coreNamespace],
} as const satisfies Record<
  keyof XlsxCoreDocumentProperties,
  readonly [string, () => string]
>;

export function parseXlsxCoreDocumentProperties(
  value: XmlLookupValue,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxCoreDocumentProperties {
  const definition = root(
    value,
    'coreProperties',
    new Set([coreNamespace()]),
    part,
  );
  const bySourceName = new Map<
    string,
    { namespace: string; property: keyof XlsxCoreDocumentProperties }
  >(
    Object.entries(CORE_FIELDS).map(([property, [source, namespace]]) => [
      source,
      {
        namespace: namespace(),
        property: property as keyof XlsxCoreDocumentProperties,
      },
    ]),
  );
  const output: XlsxCoreDocumentProperties = {};
  for (const [qualifiedName, node] of directChildren(definition.node)) {
    const source = localName(qualifiedName);
    const field = bySourceName.get(source);
    if (!field) {
      fail(
        'unsupported-feature',
        'Core document property is unsupported',
        part,
      );
    }
    if (
      namespaceFor(qualifiedName, node, definition.attrs) !== field.namespace
    ) {
      fail(
        'invalid-document-structure',
        'Core property has the wrong namespace',
        part,
      );
    }
    if (output[field.property] !== undefined) {
      fail(
        'invalid-document-structure',
        'Core document property is duplicated',
        part,
      );
    }
    const raw = scalar(node, 'Core document property is invalid', part);
    if (field.property === 'created' || field.property === 'modified') {
      const type = attributes(record(node) ?? {})['xsi:type'];
      if (typeof type !== 'string') {
        fail(
          'invalid-document-value',
          'Core property date type is invalid',
          part,
        );
      }
      const [prefix, name] = type.split(':');
      if (
        name !== 'W3CDTF' ||
        definition.attrs[`xmlns:${prefix}`] !== dctermsNamespace() ||
        definition.attrs['xmlns:xsi'] !== xsiNamespace()
      ) {
        fail(
          'invalid-document-value',
          'Core property date type is invalid',
          part,
        );
      }
      parseXlsxPivotDateTime(raw, 'Core property date is invalid', part);
    } else if (field.property === 'lastPrinted') {
      parseXlsxPivotDateTime(raw, 'Core property date is invalid', part);
    }
    output[field.property] = text(raw, budget, limits, part);
  }
  return output;
}

const APP_STRING_FIELDS = {
  AppVersion: 'applicationVersion',
  Application: 'application',
  Company: 'company',
  HyperlinkBase: 'hyperlinkBase',
  Manager: 'manager',
  PresentationFormat: 'presentationFormat',
  Template: 'template',
} as const satisfies Record<string, keyof XlsxApplicationDocumentProperties>;

const APP_INTEGER_FIELDS = {
  Characters: 'characters',
  CharactersWithSpaces: 'charactersWithSpaces',
  DocSecurity: 'documentSecurity',
  HiddenSlides: 'hiddenSlides',
  Lines: 'lines',
  MMClips: 'multimediaClips',
  Notes: 'notes',
  Pages: 'pages',
  Paragraphs: 'paragraphs',
  Slides: 'slides',
  TotalTime: 'totalTimeMinutes',
  Words: 'words',
} as const satisfies Record<string, keyof XlsxApplicationDocumentProperties>;

const APP_BOOLEAN_FIELDS = {
  HyperlinksChanged: 'hyperlinksChanged',
  LinksUpToDate: 'linksUpToDate',
  ScaleCrop: 'scaleCrop',
  SharedDoc: 'sharedDocument',
} as const satisfies Record<string, keyof XlsxApplicationDocumentProperties>;

function valueTypeChild(
  node: unknown,
  inherited: Readonly<Record<string, string>>,
  part: string,
): [string, unknown] {
  const children = directChildren(
    requiredRecord(node, 'Document-property typed value is invalid', part),
  );
  if (children.length !== 1) {
    fail(
      'invalid-document-structure',
      'Document-property typed value is invalid',
      part,
    );
  }
  const child = children[0]!;
  if (!hasNamespace(valueTypeNamespaces(), child[0], child[1], inherited)) {
    fail(
      'invalid-document-structure',
      'Document-property value has the wrong namespace',
      part,
    );
  }
  return child;
}

function parseHeadingPairs(
  node: unknown,
  inherited: Readonly<Record<string, string>>,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): Array<{ count: number; heading: string }> {
  const [vectorName, vectorValue] = valueTypeChild(node, inherited, part);
  const vector = requiredRecord(
    vectorValue,
    'Heading-pair vector is invalid',
    part,
  );
  const attrs = attributes(vector);
  if (localName(vectorName) !== 'vector' || attrs.baseType !== 'variant') {
    fail('invalid-document-value', 'Heading-pair vector is invalid', part);
  }
  const variants = directChildren(vector);
  if (
    variants.some(
      ([name, child]) =>
        localName(name) !== 'variant' ||
        !hasNamespace(valueTypeNamespaces(), name, child, inherited),
    ) ||
    variants.length % 2 !== 0 ||
    parseXlsxDocumentUnsignedInteger(
      attrs.size,
      'Heading-pair vector size is invalid',
      part,
    ) !== variants.length
  ) {
    fail('invalid-document-value', 'Heading-pair vector is invalid', part);
  }
  const output: Array<{ count: number; heading: string }> = [];
  for (let index = 0; index < variants.length; index += 2) {
    const [headingName, headingNode] = valueTypeChild(
      variants[index]![1],
      inherited,
      part,
    );
    const [countName, countNode] = valueTypeChild(
      variants[index + 1]![1],
      inherited,
      part,
    );
    if (
      !['lpstr', 'lpwstr'].includes(localName(headingName)) ||
      !['i4', 'int', 'ui4', 'uint'].includes(localName(countName))
    ) {
      fail('invalid-document-value', 'Heading-pair value is invalid', part);
    }
    output.push({
      count: parseXlsxDocumentUnsignedInteger(
        scalar(countNode, 'Heading-pair count is invalid', part),
        'Heading-pair count is invalid',
        part,
      ),
      heading: text(
        scalar(headingNode, 'Heading-pair name is invalid', part),
        budget,
        limits,
        part,
      ),
    });
  }
  return output;
}

function parseTitles(
  node: unknown,
  inherited: Readonly<Record<string, string>>,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): string[] {
  const [vectorName, vectorValue] = valueTypeChild(node, inherited, part);
  const vector = requiredRecord(
    vectorValue,
    'Part-title vector is invalid',
    part,
  );
  const attrs = attributes(vector);
  const children = directChildren(vector);
  if (
    localName(vectorName) !== 'vector' ||
    !['lpstr', 'lpwstr'].includes(attrs.baseType as string) ||
    children.some(([name]) => localName(name) !== attrs.baseType) ||
    parseXlsxDocumentUnsignedInteger(
      attrs.size,
      'Part-title vector size is invalid',
      part,
    ) !== children.length
  ) {
    fail('invalid-document-value', 'Part-title vector is invalid', part);
  }
  return children.map(([, child]) =>
    text(scalar(child, 'Part title is invalid', part), budget, limits, part),
  );
}

export function parseXlsxApplicationDocumentProperties(
  value: XmlLookupValue,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxApplicationDocumentProperties {
  const definition = root(value, 'Properties', applicationNamespaces(), part);
  const output: XlsxApplicationDocumentProperties = {};
  const seen = new Set<string>();
  for (const [qualifiedName, node] of directChildren(definition.node)) {
    const name = localName(qualifiedName);
    if (seen.has(name)) {
      fail(
        'invalid-document-structure',
        'Application property is duplicated',
        part,
      );
    }
    seen.add(name);
    if (
      !hasNamespace(
        applicationNamespaces(),
        qualifiedName,
        node,
        definition.attrs,
      )
    ) {
      fail(
        'invalid-document-structure',
        'Application property has the wrong namespace',
        part,
      );
    }
    const stringField =
      APP_STRING_FIELDS[name as keyof typeof APP_STRING_FIELDS];
    if (stringField) {
      (output as Record<string, unknown>)[stringField] = text(
        scalar(node, 'Application property is invalid', part),
        budget,
        limits,
        part,
      );
      continue;
    }
    const integerField =
      APP_INTEGER_FIELDS[name as keyof typeof APP_INTEGER_FIELDS];
    if (integerField) {
      (output as Record<string, unknown>)[integerField] =
        parseXlsxDocumentUnsignedInteger(
          scalar(node, 'Application property is invalid', part),
          'Application property is invalid',
          part,
        );
      continue;
    }
    const booleanField =
      APP_BOOLEAN_FIELDS[name as keyof typeof APP_BOOLEAN_FIELDS];
    if (booleanField) {
      (output as Record<string, unknown>)[booleanField] =
        parseXlsxDocumentBoolean(
          scalar(node, 'Application property is invalid', part),
          'Application property is invalid',
          part,
        );
      continue;
    }
    if (name === 'HeadingPairs') {
      output.headingPairs = parseHeadingPairs(
        node,
        definition.attrs,
        budget,
        limits,
        part,
      );
      continue;
    }
    if (name === 'TitlesOfParts') {
      output.titlesOfParts = parseTitles(
        node,
        definition.attrs,
        budget,
        limits,
        part,
      );
      continue;
    }
    if (name === 'DigSig' || name === 'HLinks') {
      fail(
        'unsupported-feature',
        'Application property is not safely representable',
        part,
      );
    }
    fail(
      'unsupported-feature',
      'Application document property is unsupported',
      part,
    );
  }
  return output;
}

const SIGNED_INTEGER_BOUNDS = {
  i1: [-128n, 127n],
  i2: [-32768n, 32767n],
  i4: [-2147483648n, 2147483647n],
  i8: [-9223372036854775808n, 9223372036854775807n],
  int: [-2147483648n, 2147483647n],
} as const;
const UNSIGNED_INTEGER_BOUNDS = {
  ui1: 255n,
  ui2: 65535n,
  ui4: 4294967295n,
  ui8: 18446744073709551615n,
  uint: 4294967295n,
} as const;

function integerValue(
  kind: string,
  value: string,
  part: string,
): XlsxCustomDocumentPropertyValue {
  if (!/^-?(?:0|[1-9]\d*)$/u.test(value) || value === '-0') {
    fail('invalid-document-value', 'Custom property integer is invalid', part);
  }
  const parsed = BigInt(value);
  const signed =
    SIGNED_INTEGER_BOUNDS[kind as keyof typeof SIGNED_INTEGER_BOUNDS];
  const unsigned =
    UNSIGNED_INTEGER_BOUNDS[kind as keyof typeof UNSIGNED_INTEGER_BOUNDS];
  if (
    (signed && (parsed < signed[0] || parsed > signed[1])) ||
    (unsigned !== undefined && (parsed < 0n || parsed > unsigned))
  ) {
    fail('invalid-document-value', 'Custom property integer is invalid', part);
  }
  return { kind: 'integer', value };
}

function customValue(
  name: string,
  node: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxCustomDocumentPropertyValue {
  const kind = localName(name);
  const raw = scalar(node, 'Custom property value is invalid', part);
  if (kind === 'empty') {
    if (raw !== '')
      fail('invalid-document-value', 'Custom empty property is invalid', part);
    return { kind: 'empty' };
  }
  if (kind === 'null') {
    if (raw !== '')
      fail('invalid-document-value', 'Custom null property is invalid', part);
    return { kind: 'null' };
  }
  if (['lpstr', 'lpwstr', 'bstr'].includes(kind)) {
    return { kind: 'string', value: text(raw, budget, limits, part) };
  }
  if (kind in SIGNED_INTEGER_BOUNDS || kind in UNSIGNED_INTEGER_BOUNDS) {
    return integerValue(kind, text(raw, budget, limits, part), part);
  }
  if (kind === 'bool') {
    return {
      kind: 'boolean',
      value: parseXlsxDocumentBoolean(
        raw,
        'Custom property boolean is invalid',
        part,
      ),
    };
  }
  if (kind === 'r4' || kind === 'r8') {
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/u.test(raw)) {
      fail('invalid-document-value', 'Custom property number is invalid', part);
    }
    const parsed = Number(raw);
    if (!Number.isFinite(parsed)) {
      fail('invalid-document-value', 'Custom property number is invalid', part);
    }
    return { kind: 'number', value: Object.is(parsed, -0) ? 0 : parsed };
  }
  if (kind === 'decimal' || kind === 'cy') {
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/u.test(raw)) {
      fail(
        'invalid-document-value',
        'Custom property decimal is invalid',
        part,
      );
    }
    return { kind: 'decimal', value: text(raw, budget, limits, part) };
  }
  if (kind === 'date' || kind === 'filetime') {
    return {
      kind: 'date-time',
      value: text(
        parseXlsxPivotDateTime(raw, 'Custom property date is invalid', part),
        budget,
        limits,
        part,
      ),
    };
  }
  fail(
    'unsupported-feature',
    'Custom property value type is unsupported',
    part,
  );
}

export function parseXlsxCustomDocumentProperties(
  value: XmlLookupValue,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxCustomDocumentProperty[] {
  const definition = root(value, 'Properties', customNamespaces(), part);
  const output: XlsxCustomDocumentProperty[] = [];
  const ids = new Set<number>();
  const names = new Set<string>();
  for (const [qualifiedName, nodeValue] of directChildren(definition.node)) {
    const node = requiredRecord(
      nodeValue,
      'Custom document property is invalid',
      part,
    );
    if (
      localName(qualifiedName) !== 'property' ||
      !hasNamespace(customNamespaces(), qualifiedName, node, definition.attrs)
    ) {
      fail(
        'unsupported-feature',
        'Custom document property is unsupported',
        part,
      );
    }
    const attrs = attributes(node);
    const propertyId = parseXlsxDocumentUnsignedInteger(
      attrs.pid,
      'Custom property ID is invalid',
      part,
    );
    if (propertyId < 2 || ids.has(propertyId)) {
      fail(
        'invalid-document-value',
        'Custom property IDs are invalid or duplicated',
        part,
      );
    }
    ids.add(propertyId);
    const propertyName = decodeXmlEntities(attrs.name ?? '');
    if (!propertyName || names.has(propertyName.toUpperCase())) {
      fail(
        'invalid-document-value',
        'Custom property names are invalid or duplicated',
        part,
      );
    }
    names.add(propertyName.toUpperCase());
    const formatId = attrs.fmtid?.trim();
    if (
      !formatId ||
      !/^\{[0-9A-Fa-f]{8}(?:-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}\}$/u.test(
        formatId,
      )
    ) {
      fail(
        'invalid-document-value',
        'Custom property format ID is invalid',
        part,
      );
    }
    const [valueName, valueNode] = valueTypeChild(node, definition.attrs, part);
    output.push({
      formatId: text(formatId, budget, limits, part),
      ...(attrs.linkTarget === undefined
        ? {}
        : {
            linkTarget: text(
              decodeXmlEntities(attrs.linkTarget),
              budget,
              limits,
              part,
            ),
          }),
      name: text(propertyName, budget, limits, part),
      propertyId,
      value: customValue(valueName, valueNode, budget, limits, part),
    });
  }
  return output;
}

export async function loadXlsxDocumentProperties(
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  budget: XlsxWorksheetBudget,
): Promise<XlsxDocumentProperties | undefined> {
  const relationshipsXml = await reader.readXml('_rels/.rels');
  const relationships = parseXlsxRelationships(
    relationshipsXml,
    null,
    limits.maxRelationships,
  );
  const corePart = propertyTarget(
    relationships,
    coreRelationshipTypes(),
    coreContentType(),
    discovery,
  );
  const applicationPart = propertyTarget(
    relationships,
    applicationRelationshipTypes(),
    applicationContentType(),
    discovery,
  );
  const customPart = propertyTarget(
    relationships,
    customRelationshipTypes(),
    customContentType(),
    discovery,
  );
  if (!corePart && !applicationPart && !customPart) return undefined;
  return {
    ...(applicationPart === undefined
      ? {}
      : {
          application: parseXlsxApplicationDocumentProperties(
            await reader.readXml(applicationPart, { required: true }),
            budget,
            limits,
            applicationPart,
          ),
        }),
    ...(corePart === undefined
      ? {}
      : {
          core: parseXlsxCoreDocumentProperties(
            await reader.readXml(corePart, { required: true }),
            budget,
            limits,
            corePart,
          ),
        }),
    ...(customPart === undefined
      ? {}
      : {
          custom: parseXlsxCustomDocumentProperties(
            await reader.readXml(customPart, { required: true }),
            budget,
            limits,
            customPart,
          ),
        }),
  };
}
