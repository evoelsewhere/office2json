import { XlsxParseError } from '../errors';
import {
  resolveXlsxPartTarget,
  resolveXlsxRootTarget,
} from './package-identity';
import { XlsxResourceLimitError } from './resource-limits';

export const XLSX_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';

type XmlRecord = Record<string, unknown>;

export type XlsxRelationship =
  | { id: string; mode: 'internal'; target: string; type: string }
  | { id: string; mode: 'external'; target: string; type: string };

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function records(value: unknown): XmlRecord[] | undefined {
  if (value === undefined) return [];
  const values: unknown[] = Array.isArray(value) ? value : [value];
  const output: XmlRecord[] = [];
  for (const item of values) {
    const itemRecord = record(item);
    if (!itemRecord) return undefined;
    output.push(itemRecord);
  }
  return output;
}

function attributes(value: XmlRecord): XmlRecord {
  return record(value.attrs) ?? {};
}

function relationshipFailure(
  part: string,
  message: string,
  code: 'invalid-document-structure' | 'invalid-relationship-target',
  cause?: unknown,
): never {
  throw new XlsxParseError(
    { code, message, part, severity: 'error' },
    { cause },
  );
}

function validToken(value: unknown): value is string {
  return (
    typeof value === 'string' && value.length > 0 && value === value.trim()
  );
}

export function parseXlsxRelationships(
  value: unknown,
  ownerPart: string | null,
  maxRelationships: number,
): ReadonlyMap<string, XlsxRelationship> {
  const relationshipPart = ownerPart ?? '_rels/.rels';
  const root = record(record(value)?.Relationships);
  if (!root || attributes(root).xmlns !== XLSX_RELATIONSHIPS_NAMESPACE) {
    relationshipFailure(
      relationshipPart,
      'Relationship root is missing or has the wrong namespace',
      'invalid-document-structure',
    );
  }
  const relationshipNodes = records(root.Relationship);
  if (!relationshipNodes) {
    relationshipFailure(
      relationshipPart,
      'Relationships contain an invalid entry collection',
      'invalid-document-structure',
    );
  }
  if (relationshipNodes.length > maxRelationships) {
    throw new XlsxResourceLimitError(
      'maxRelationships',
      relationshipNodes.length,
      maxRelationships,
      relationshipPart,
    );
  }

  const relationships = new Map<string, XlsxRelationship>();
  for (const relationshipNode of relationshipNodes) {
    const attrs = attributes(relationshipNode);
    if (!validToken(attrs.Id)) {
      relationshipFailure(
        relationshipPart,
        'Relationship has an invalid ID',
        'invalid-document-structure',
      );
    }
    if (!validToken(attrs.Type)) {
      relationshipFailure(
        relationshipPart,
        'Relationship has an invalid type',
        'invalid-document-structure',
      );
    }
    if (relationships.has(attrs.Id)) {
      relationshipFailure(
        relationshipPart,
        'Relationships contain a duplicate ID',
        'invalid-document-structure',
      );
    }

    let relationship: XlsxRelationship;
    if (attrs.TargetMode === undefined) {
      if (typeof attrs.Target !== 'string') {
        relationshipFailure(
          relationshipPart,
          'Relationship has an invalid internal target',
          'invalid-relationship-target',
        );
      }
      let target: string;
      try {
        target =
          ownerPart === null
            ? resolveXlsxRootTarget(attrs.Target)
            : resolveXlsxPartTarget(ownerPart, attrs.Target);
      } catch (cause) {
        relationshipFailure(
          relationshipPart,
          'Relationship has an invalid internal target',
          'invalid-relationship-target',
          cause,
        );
      }
      relationship = {
        id: attrs.Id,
        mode: 'internal',
        target,
        type: attrs.Type,
      };
    } else if (attrs.TargetMode === 'External') {
      if (!validToken(attrs.Target)) {
        relationshipFailure(
          relationshipPart,
          'Relationship has an invalid external target',
          'invalid-relationship-target',
        );
      }
      relationship = {
        id: attrs.Id,
        mode: 'external',
        target: attrs.Target,
        type: attrs.Type,
      };
    } else {
      relationshipFailure(
        relationshipPart,
        'Relationship has an invalid TargetMode',
        'invalid-document-structure',
      );
    }
    relationships.set(relationship.id, relationship);
  }
  return relationships;
}
