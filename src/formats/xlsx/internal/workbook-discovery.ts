import type { XmlLookupValue } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import {
  parseXlsxContentTypes,
  type XlsxContentTypeTable,
} from './content-types';
import { XlsxPartReader } from './part-reader';
import { parseXlsxRelationships, type XlsxRelationship } from './relationships';
import type { ResolvedXlsxResourceLimits } from './resource-limits';

export const XLSX_SPREADSHEET_NAMESPACES = {
  strict: 'http://purl.oclc.org/ooxml/spreadsheetml/main',
  transitional: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
} as const;

export const XLSX_OFFICE_DOCUMENT_RELATIONSHIP_TYPES = new Set([
  'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
]);

export const XLSX_WORKBOOK_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';

const ACTIVE_OR_BINARY_MAIN_CONTENT_TYPES = new Set([
  'application/vnd.ms-excel.addin.macroEnabled.main+xml',
  'application/vnd.ms-excel.sheet.binary.macroEnabled.main',
  'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
  'application/vnd.ms-excel.template.macroEnabled.main+xml',
]);

type XmlRecord = Record<string, unknown>;

export interface XlsxWorkbookDiscovery {
  contentTypes: XlsxContentTypeTable;
  dialect: keyof typeof XLSX_SPREADSHEET_NAMESPACES;
  part: string;
  root: XmlLookupValue;
}

function attributes(value: XmlRecord): XmlRecord {
  return (value.attrs ?? {}) as XmlRecord;
}

function fail(
  code:
    | 'invalid-document-structure'
    | 'invalid-relationship-target'
    | 'security-rejected-content',
  message: string,
  part: string,
): never {
  throw new XlsxParseError({ code, message, part, severity: 'error' });
}

function selectMainRelationship(
  relationships: ReadonlyMap<string, XlsxRelationship>,
): Extract<XlsxRelationship, { mode: 'internal' }> {
  const candidates = [...relationships.values()].filter((relationship) =>
    XLSX_OFFICE_DOCUMENT_RELATIONSHIP_TYPES.has(relationship.type),
  );
  if (candidates.length !== 1) {
    fail(
      'invalid-document-structure',
      'Package must contain exactly one office-document relationship',
      '_rels/.rels',
    );
  }
  const relationship = candidates[0]!;
  if (relationship.mode !== 'internal') {
    fail(
      'invalid-relationship-target',
      'Package office-document relationship must be internal',
      '_rels/.rels',
    );
  }
  return relationship;
}

function workbookDialect(
  value: XmlLookupValue,
  part: string,
): keyof typeof XLSX_SPREADSHEET_NAMESPACES {
  const document = value as unknown as XmlRecord;
  for (const [qualifiedName, candidate] of Object.entries(document)) {
    const [first, second] = qualifiedName.split(':') as [string, string?];
    const prefix = second === undefined ? '' : first;
    const localName = second ?? first;
    if (localName !== 'workbook') continue;
    const attrs = attributes(candidate as XmlRecord);
    const namespace = attrs[prefix ? `xmlns:${prefix}` : 'xmlns'];
    for (const [dialect, expected] of Object.entries(
      XLSX_SPREADSHEET_NAMESPACES,
    )) {
      if (namespace === expected) {
        return dialect as keyof typeof XLSX_SPREADSHEET_NAMESPACES;
      }
    }
  }
  fail(
    'invalid-document-structure',
    'Workbook root is missing or has an unsupported namespace',
    part,
  );
}

export async function discoverXlsxWorkbook(
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
): Promise<XlsxWorkbookDiscovery> {
  const contentTypesXml = await reader.readXml('[Content_Types].xml', {
    required: true,
  });
  const packageRelationshipsXml = await reader.readXml('_rels/.rels', {
    required: true,
  });
  const contentTypes = parseXlsxContentTypes(contentTypesXml);
  const packageRelationships = parseXlsxRelationships(
    packageRelationshipsXml,
    null,
    limits.maxRelationships,
  );
  const mainRelationship = selectMainRelationship(packageRelationships);
  const contentType = contentTypes.contentTypeFor(mainRelationship.target);
  if (
    contentType !== undefined &&
    ACTIVE_OR_BINARY_MAIN_CONTENT_TYPES.has(contentType)
  ) {
    fail(
      'security-rejected-content',
      'Macro-enabled or binary spreadsheet main parts are not accepted',
      mainRelationship.target,
    );
  }
  if (contentType !== XLSX_WORKBOOK_CONTENT_TYPE) {
    fail(
      'invalid-document-structure',
      'Office-document relationship does not target an XLSX workbook main part',
      mainRelationship.target,
    );
  }

  const root = await reader.readXml(mainRelationship.target, {
    required: true,
  });
  return {
    contentTypes,
    dialect: workbookDialect(root, mainRelationship.target),
    part: mainRelationship.target,
    root,
  };
}
