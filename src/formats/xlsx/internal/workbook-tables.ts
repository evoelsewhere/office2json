import { XlsxParseError } from '../errors';
import { getXlsxRelationshipPartName } from './package-identity';
import { XlsxPartReader } from './part-reader';
import { parseXlsxRelationships } from './relationships';
import type { ResolvedXlsxResourceLimits } from './resource-limits';
import {
  parseXlsxSharedStringPart,
  type XlsxSharedStringTable,
} from './shared-strings';
import type { XlsxWorkbookDiscovery } from './workbook-discovery';

const SHARED_STRINGS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
const SHARED_STRINGS_RELATIONSHIP_TYPE = {
  strict:
    'http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings',
  transitional:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
} as const;

const EMPTY_SHARED_STRINGS: XlsxSharedStringTable = Object.freeze({
  part: null,
  values: Object.freeze([]),
});

function structureFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-structure',
    message,
    part,
    severity: 'error',
  });
}

function externalFailure(part: string, relationshipType: string): never {
  throw new XlsxParseError({
    code: 'invalid-relationship-target',
    message: 'Workbook shared-string relationship must be internal',
    part,
    relationshipType,
    severity: 'error',
  });
}

export async function loadXlsxSharedStrings(
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
): Promise<XlsxSharedStringTable> {
  const relationshipPart = getXlsxRelationshipPartName(discovery.part);
  const relationshipXml = await reader.readXml(relationshipPart, {
    required: true,
  });
  const relationships = parseXlsxRelationships(
    relationshipXml,
    discovery.part,
    limits.maxRelationships,
  );
  const relationshipType = SHARED_STRINGS_RELATIONSHIP_TYPE[discovery.dialect];
  const candidates = [...relationships.values()].filter(
    (relationship) => relationship.type === relationshipType,
  );
  if (candidates.length === 0) return EMPTY_SHARED_STRINGS;
  if (candidates.length !== 1) {
    structureFailure(
      'Workbook contains multiple shared-string relationships',
      relationshipPart,
    );
  }
  const relationship = candidates[0]!;
  if (relationship.mode !== 'internal') {
    externalFailure(relationshipPart, relationship.type);
  }
  if (
    discovery.contentTypes.contentTypeFor(relationship.target) !==
    SHARED_STRINGS_CONTENT_TYPE
  ) {
    structureFailure(
      'Workbook shared-string target has the wrong content type',
      relationship.target,
    );
  }
  return parseXlsxSharedStringPart(
    relationship.target,
    discovery.dialect,
    reader,
    limits,
  );
}
