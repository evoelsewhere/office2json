import {
  decodeXmlEntities,
  sanitizeHyperlink,
} from '../../../common/text/html';
import { XlsxParseError } from '../errors';
import type { XlsxHyperlinkTarget, XlsxRange } from '../types';
import { parseXlsxRangeReference } from './cell-reference';
import type { XlsxRelationship } from './relationships';
import type { XlsxXmlElement } from './streaming-xml';

const OFFICE_RELATIONSHIP_NAMESPACE = {
  strict: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
  transitional:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
} as const;

export interface ParsedXlsxHyperlink {
  display?: string;
  range: XlsxRange;
  target: XlsxHyperlinkTarget;
  textCharacters: number;
  tooltip?: string;
}

function fail(
  part: string,
  message: string,
  code: 'invalid-document-value' | 'invalid-relationship-target',
): never {
  throw new XlsxParseError({ code, message, part, severity: 'error' });
}

function securityFailure(part: string): never {
  throw new XlsxParseError({
    code: 'security-rejected-content',
    message: 'Worksheet hyperlink protocol is not allowed',
    part,
    relationshipType: 'hyperlink',
    severity: 'error',
  });
}

function attribute(
  element: XlsxXmlElement,
  localName: string,
): string | undefined {
  return element.attributes.get(`{}${localName}`);
}

function relationshipId(
  element: XlsxXmlElement,
  dialect: 'strict' | 'transitional',
): string | undefined {
  return element.attributes.get(
    `{${OFFICE_RELATIONSHIP_NAMESPACE[dialect]}}id`,
  );
}

function normalizedExternalTarget(value: string, part: string): string {
  const decoded = decodeXmlEntities(value);
  const safe = sanitizeHyperlink(decoded);
  if (!safe) securityFailure(part);
  const url = new URL(safe);
  url.username = '';
  url.password = '';
  return url.toString();
}

export function parseXlsxHyperlink(
  element: XlsxXmlElement,
  dialect: 'strict' | 'transitional',
  relationships: ReadonlyMap<string, XlsxRelationship>,
  part: string,
): ParsedXlsxHyperlink {
  const source = attribute(element, 'ref');
  const range = parseXlsxRangeReference(source);
  if (!range || source?.includes('$')) {
    fail(
      part,
      'Worksheet hyperlink range is invalid',
      'invalid-document-value',
    );
  }
  const id = relationshipId(element, dialect);
  const location = attribute(element, 'location');
  let target: XlsxHyperlinkTarget;
  if (id === undefined) {
    if (location === undefined || location.length === 0) {
      fail(
        part,
        'Worksheet hyperlink target is missing',
        'invalid-document-value',
      );
    }
    target = { kind: 'internal', location };
  } else {
    const relationship = relationships.get(id);
    const expectedType = `${OFFICE_RELATIONSHIP_NAMESPACE[dialect]}/hyperlink`;
    if (
      !relationship ||
      relationship.mode !== 'external' ||
      relationship.type !== expectedType
    ) {
      fail(
        part,
        'Worksheet hyperlink relationship is invalid',
        'invalid-relationship-target',
      );
    }
    target = {
      kind: 'external',
      ...(location === undefined ? {} : { location }),
      url: normalizedExternalTarget(relationship.target, part),
    };
  }
  const display = attribute(element, 'display');
  const tooltip = attribute(element, 'tooltip');
  return {
    ...(display === undefined ? {} : { display }),
    range,
    target,
    textCharacters:
      (display?.length ?? 0) +
      (tooltip?.length ?? 0) +
      (target.kind === 'internal'
        ? target.location.length
        : target.url.length + (target.location?.length ?? 0)),
    ...(tooltip === undefined ? {} : { tooltip }),
  };
}
