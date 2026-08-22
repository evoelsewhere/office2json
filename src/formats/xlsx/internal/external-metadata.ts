import type { XmlLookupValue } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type {
  XlsxConnection,
  XlsxExternalDefinedName,
  XlsxExternalLink,
  XlsxExternalLinkTarget,
  XlsxQueryTable,
} from '../types';
import { getXlsxRelationshipPartName } from './package-identity';
import { XlsxPartReader } from './part-reader';
import { parseXlsxRelationships, type XlsxRelationship } from './relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import type { XlsxWorkbookDiscovery } from './workbook-discovery';
import {
  consumeXlsxWorksheetBudget,
  consumeXlsxWorksheetFormulaCharacters,
  type XlsxWorksheetBudget,
} from './worksheet';

type XmlRecord = Record<string, unknown>;

export interface XlsxExternalMetadataLoadResult {
  connections: XlsxConnection[];
  externalLinks: XlsxExternalLink[];
}

function fail(
  code:
    | 'invalid-document-structure'
    | 'invalid-document-value'
    | 'security-rejected-content',
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

function descendants(node: XmlRecord, name: string): XmlRecord[] {
  const output: XmlRecord[] = [];
  for (const [key, value] of Object.entries(node)) {
    const values = records(value);
    if (!values) continue;
    for (const child of values) {
      if (localName(key) === name) output.push(child);
      output.push(...descendants(child, name));
    }
  }
  return output;
}

function scalarText(value: unknown): string | undefined {
  if (typeof value === 'string') return value;
  const node = record(value);
  return typeof node?.value === 'string' ? node.value : undefined;
}

function root(
  value: XmlLookupValue,
  expectedLocalName: string,
  namespace: string,
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
  if (
    attributes(node)[sourcePrefix ? `xmlns:${sourcePrefix}` : 'xmlns'] !==
    namespace
  ) {
    fail(
      'invalid-document-structure',
      `XLSX ${expectedLocalName} root has the wrong namespace`,
      part,
    );
  }
  return node;
}

export function parseXlsxExternalUnsignedInteger(
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

function requiredUnsignedInteger(
  value: string | undefined,
  message: string,
  part: string,
): number {
  const parsed = parseXlsxExternalUnsignedInteger(
    value,
    undefined,
    message,
    part,
  );
  if (parsed === undefined) fail('invalid-document-value', message, part);
  return parsed;
}

export function parseXlsxExternalBoolean(
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

function relationshipBase(dialect: XlsxWorkbookDiscovery['dialect']): string {
  return dialect === 'strict'
    ? 'http://purl.oclc.org/ooxml/officeDocument/relationships'
    : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
}

function spreadsheetNamespace(
  dialect: XlsxWorkbookDiscovery['dialect'],
): string {
  return dialect === 'strict'
    ? 'http://purl.oclc.org/ooxml/spreadsheetml/main'
    : 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
}

function internalTargets(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  type: string,
  contentType: string,
  discovery: XlsxWorkbookDiscovery,
  owner: string,
): string[] {
  const output: string[] = [];
  const targets = new Set<string>();
  for (const relationship of relationships.values()) {
    if (relationship.type !== type) continue;
    if (relationship.mode !== 'internal') {
      fail(
        'invalid-document-structure',
        'External metadata relationship must be internal',
        owner,
      );
    }
    if (targets.has(relationship.target)) {
      fail(
        'invalid-document-structure',
        'External metadata relationship target is duplicated',
        owner,
      );
    }
    targets.add(relationship.target);
    if (
      discovery.contentTypes.contentTypeFor(relationship.target) !== contentType
    ) {
      fail(
        'invalid-document-structure',
        'External metadata target has the wrong content type',
        relationship.target,
      );
    }
    output.push(relationship.target);
  }
  return output;
}

export function redactXlsxExternalTarget(
  value: string,
  part: string,
): XlsxExternalLinkTarget {
  if (/^[^/:?#]+(?:[\\/]|$)/u.test(value)) {
    return { kind: 'relative', redacted: true };
  }
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    fail('invalid-document-value', 'External workbook target is invalid', part);
  }
  const protocol = url.protocol.toLowerCase();
  if (protocol === 'file:') return { kind: 'file', redacted: true };
  if (protocol === 'http:' || protocol === 'https:') {
    return {
      host: url.hostname.toLowerCase(),
      kind: protocol === 'http:' ? 'http' : 'https',
      redacted: true,
    };
  }
  fail(
    'security-rejected-content',
    'External workbook protocol is not allowed',
    part,
  );
}

async function externalLinkTarget(
  part: string,
  definition: XmlRecord,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
): Promise<XlsxExternalLinkTarget | undefined> {
  const books = descendants(definition, 'externalBook');
  if (books.length === 0) return undefined;
  if (books.length !== 1) {
    fail(
      'invalid-document-structure',
      'External workbook owner is duplicated',
      part,
    );
  }
  const relationshipId = attributes(books[0]!)['r:id'];
  if (!relationshipId) {
    fail(
      'invalid-document-value',
      'External workbook relationship reference is invalid',
      part,
    );
  }
  const relationshipPart = getXlsxRelationshipPartName(part);
  const relationshipXml = await reader.readXml(relationshipPart);
  if (relationshipXml === null) {
    fail(
      'invalid-document-structure',
      'External workbook relationships are missing',
      part,
    );
  }
  const relationships = parseXlsxRelationships(
    relationshipXml,
    part,
    limits.maxRelationships,
  );
  const expectedType = `${relationshipBase(discovery.dialect)}/externalLinkPath`;
  const relationship = relationships.get(relationshipId);
  if (
    !relationship ||
    relationship.type !== expectedType ||
    relationship.mode !== 'external'
  ) {
    fail(
      'invalid-document-structure',
      'External workbook path relationship is invalid',
      part,
    );
  }
  return redactXlsxExternalTarget(relationship.target, part);
}

function externalDefinedNames(
  definition: XmlRecord,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxExternalDefinedName[] {
  return descendants(definition, 'definedName').map((node) => {
    const attrs = attributes(node);
    if (!attrs.name) {
      fail('invalid-document-value', 'External defined name is invalid', part);
    }
    const formula = attrs.refersTo ?? scalarText(node);
    if (!formula || formula !== formula.trim()) {
      fail(
        'invalid-document-value',
        'External defined-name formula is invalid',
        part,
      );
    }
    consumeXlsxWorksheetFormulaCharacters(budget, formula, limits, part);
    const sheetId = parseXlsxExternalUnsignedInteger(
      attrs.sheetId,
      undefined,
      'External defined-name sheet ID is invalid',
      part,
    );
    return {
      formula,
      name: text(attrs.name, budget, limits, part),
      ...(sheetId === undefined ? {} : { sheetId }),
    };
  });
}

async function loadExternalLinks(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  budget: XlsxWorksheetBudget,
): Promise<XlsxExternalLink[]> {
  const targets = internalTargets(
    relationships,
    `${relationshipBase(discovery.dialect)}/externalLink`,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
    discovery,
    discovery.part,
  );
  const output: XlsxExternalLink[] = [];
  for (const [index, part] of targets.entries()) {
    const value = await reader.readXml(part, { required: true });
    const definition = root(
      value,
      'externalLink',
      spreadsheetNamespace(discovery.dialect),
      part,
    );
    if (
      descendants(definition, 'ddeLink').length > 0 ||
      descendants(definition, 'oleLink').length > 0
    ) {
      fail(
        'security-rejected-content',
        'External DDE or OLE link is not allowed',
        part,
      );
    }
    const sheetNames = descendants(definition, 'sheetName').map((node) => {
      const value = attributes(node).val;
      if (!value) {
        fail('invalid-document-value', 'External sheet name is invalid', part);
      }
      return text(value, budget, limits, part);
    });
    const target = await externalLinkTarget(
      part,
      definition,
      discovery,
      reader,
      limits,
    );
    output.push({
      definedNames: externalDefinedNames(definition, budget, limits, part),
      index,
      sheetNames,
      ...(target === undefined ? {} : { target }),
    });
  }
  return output;
}

export function xlsxConnectionCredentialsOmitted(
  attrs: Readonly<Record<string, string>>,
): boolean {
  return (
    attrs.credentials !== undefined ||
    attrs.singleSignOnId !== undefined ||
    attrs.ssoId !== undefined
  );
}

export function xlsxConnectionSourceDataOmitted(value: unknown): boolean {
  const definition = record(value);
  if (!definition) return false;
  const secretNames = new Set([
    'command',
    'connection',
    'credentials',
    'odcFile',
    'serverCredentialsMethod',
    'sourceFile',
    'ssoId',
    'url',
  ]);
  const containsSecret = (node: XmlRecord): boolean => {
    if (Object.keys(attributes(node)).some((name) => secretNames.has(name))) {
      return true;
    }
    return Object.values(node).some((value) => {
      const values = records(value);
      return values?.some((child) => containsSecret(child)) ?? false;
    });
  };
  return containsSecret(definition);
}

async function loadConnections(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  budget: XlsxWorksheetBudget,
): Promise<XlsxConnection[]> {
  const targets = internalTargets(
    relationships,
    `${relationshipBase(discovery.dialect)}/connections`,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml',
    discovery,
    discovery.part,
  );
  const output: XlsxConnection[] = [];
  const ids = new Set<number>();
  for (const part of targets) {
    const value = await reader.readXml(part, { required: true });
    const definition = root(
      value,
      'connections',
      spreadsheetNamespace(discovery.dialect),
      part,
    );
    const connectionEntries = Object.entries(definition).filter(
      ([name]) => localName(name) === 'connection',
    );
    if (connectionEntries.length > 1) {
      fail(
        'invalid-document-structure',
        'Connection collection is duplicated',
        part,
      );
    }
    const nodes = records(connectionEntries[0]?.[1]);
    if (!nodes) {
      fail(
        'invalid-document-structure',
        'Connection collection is invalid',
        part,
      );
    }
    const expected = parseXlsxExternalUnsignedInteger(
      attributes(definition).count,
      undefined,
      'Connection count is invalid',
      part,
    );
    if (expected !== undefined && expected !== nodes.length) {
      fail(
        'invalid-document-structure',
        'Connection count does not match',
        part,
      );
    }
    if (nodes.length > limits.maxTables) {
      throw new XlsxResourceLimitError(
        'maxTables',
        nodes.length,
        limits.maxTables,
        part,
      );
    }
    for (const node of nodes) {
      const attrs = attributes(node);
      const id = requiredUnsignedInteger(
        attrs.id,
        'Connection ID is invalid',
        part,
      );
      if (ids.has(id)) {
        fail('invalid-document-value', 'Connection IDs are duplicated', part);
      }
      ids.add(id);
      const description =
        attrs.description === undefined
          ? undefined
          : text(attrs.description, budget, limits, part);
      const name =
        attrs.name === undefined
          ? undefined
          : text(attrs.name, budget, limits, part);
      output.push({
        background: parseXlsxExternalBoolean(
          attrs.background,
          false,
          'Connection background flag is invalid',
          part,
        ),
        credentialsOmitted: xlsxConnectionCredentialsOmitted(attrs),
        deleted: parseXlsxExternalBoolean(
          attrs.deleted,
          false,
          'Connection deleted flag is invalid',
          part,
        ),
        ...(description === undefined ? {} : { description }),
        id,
        keepAlive: parseXlsxExternalBoolean(
          attrs.keepAlive,
          false,
          'Connection keep-alive flag is invalid',
          part,
        ),
        ...(name === undefined ? {} : { name }),
        ...(attrs.refreshedVersion === undefined
          ? {}
          : {
              refreshedVersion: requiredUnsignedInteger(
                attrs.refreshedVersion,
                'Connection refreshed version is invalid',
                part,
              ),
            }),
        ...(attrs.interval === undefined
          ? {}
          : {
              refreshInterval: requiredUnsignedInteger(
                attrs.interval,
                'Connection refresh interval is invalid',
                part,
              ),
            }),
        refreshOnLoad: parseXlsxExternalBoolean(
          attrs.refreshOnLoad,
          false,
          'Connection refresh-on-load flag is invalid',
          part,
        ),
        saveData: parseXlsxExternalBoolean(
          attrs.saveData,
          false,
          'Connection save-data flag is invalid',
          part,
        ),
        sourceDataOmitted: xlsxConnectionSourceDataOmitted(node),
        ...(attrs.type === undefined
          ? {}
          : {
              type: requiredUnsignedInteger(
                attrs.type,
                'Connection type is invalid',
                part,
              ),
            }),
      });
    }
  }
  return output;
}

export async function loadXlsxExternalMetadata(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  budget: XlsxWorksheetBudget,
): Promise<XlsxExternalMetadataLoadResult> {
  return {
    connections: await loadConnections(
      relationships,
      discovery,
      reader,
      limits,
      budget,
    ),
    externalLinks: await loadExternalLinks(
      relationships,
      discovery,
      reader,
      limits,
      budget,
    ),
  };
}

export async function loadXlsxQueryTables(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  connectionIds: ReadonlySet<number>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  budget: XlsxWorksheetBudget,
  worksheetPart: string,
): Promise<XlsxQueryTable[]> {
  const targets = internalTargets(
    relationships,
    `${relationshipBase(discovery.dialect)}/queryTable`,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
    discovery,
    worksheetPart,
  );
  const output: XlsxQueryTable[] = [];
  const names = new Set<string>();
  for (const part of targets) {
    const value = await reader.readXml(part, { required: true });
    const definition = root(
      value,
      'queryTable',
      spreadsheetNamespace(discovery.dialect),
      part,
    );
    const attrs = attributes(definition);
    if (!attrs.name) {
      fail('invalid-document-value', 'Query table name is invalid', part);
    }
    const name = text(attrs.name, budget, limits, part);
    const folded = name.toUpperCase();
    if (names.has(folded)) {
      fail('invalid-document-value', 'Query table names are duplicated', part);
    }
    names.add(folded);
    const connectionId = requiredUnsignedInteger(
      attrs.connectionId,
      'Query table connection ID is invalid',
      part,
    );
    if (!connectionIds.has(connectionId)) {
      fail(
        'invalid-document-value',
        'Query table connection reference is invalid',
        part,
      );
    }
    output.push({
      adjustColumnWidth: parseXlsxExternalBoolean(
        attrs.adjustColumnWidth,
        true,
        'Query table adjust-column-width flag is invalid',
        part,
      ),
      applyAlignmentFormats: parseXlsxExternalBoolean(
        attrs.applyAlignmentFormats,
        false,
        'Query table alignment-format flag is invalid',
        part,
      ),
      applyBorderFormats: parseXlsxExternalBoolean(
        attrs.applyBorderFormats,
        false,
        'Query table border-format flag is invalid',
        part,
      ),
      applyFontFormats: parseXlsxExternalBoolean(
        attrs.applyFontFormats,
        false,
        'Query table font-format flag is invalid',
        part,
      ),
      applyNumberFormats: parseXlsxExternalBoolean(
        attrs.applyNumberFormats,
        false,
        'Query table number-format flag is invalid',
        part,
      ),
      applyPatternFormats: parseXlsxExternalBoolean(
        attrs.applyPatternFormats,
        false,
        'Query table pattern-format flag is invalid',
        part,
      ),
      applyWidthHeightFormats: parseXlsxExternalBoolean(
        attrs.applyWidthHeightFormats,
        false,
        'Query table width-height-format flag is invalid',
        part,
      ),
      backgroundRefresh: parseXlsxExternalBoolean(
        attrs.backgroundRefresh,
        true,
        'Query table background-refresh flag is invalid',
        part,
      ),
      connectionId,
      disableEdit: parseXlsxExternalBoolean(
        attrs.disableEdit,
        false,
        'Query table disable-edit flag is invalid',
        part,
      ),
      name,
      preserveFormatting: parseXlsxExternalBoolean(
        attrs.preserveFormatting,
        true,
        'Query table preserve-formatting flag is invalid',
        part,
      ),
      refreshOnLoad: parseXlsxExternalBoolean(
        attrs.refreshOnLoad,
        false,
        'Query table refresh-on-load flag is invalid',
        part,
      ),
    });
  }
  return output;
}
