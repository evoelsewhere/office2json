import { readZipEntryBytes } from '../../../../common/archive/read-entry';
import JSZip from 'jszip';

import { assertXlsxArchiveWithinLimits } from '../../internal/archive';
import { XlsxPartReader } from '../../internal/part-reader';
import {
  parseXlsxRelationships,
  type XlsxRelationship,
} from '../../internal/relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from '../../internal/resource-limits';
import { discoverXlsxWorkbook } from '../../internal/workbook-discovery';
import { canonicalXlsxSha256, sha256XlsxBytes } from '../digest';
import { XlsxWriteError } from '../errors';

export interface XlsxPackageGraphPart {
  byteLength: number;
  contentType: string;
  name: string;
  relationshipPart: boolean;
  sha256: string;
}

export type XlsxPackageGraphRelationship = XlsxRelationship & {
  owner: string | null;
};

export interface XlsxPackageGraph {
  conformance: 'strict' | 'transitional';
  containsActiveContent: boolean;
  containsDigitalSignatures: boolean;
  containsExternalRelationships: boolean;
  containsOpaqueContent: boolean;
  manifestHash: string;
  parts: XlsxPackageGraphPart[];
  relationships: XlsxPackageGraphRelationship[];
}

export function xlsxRelationshipOwner(part: string): string | null | undefined {
  if (part === '_rels/.rels') return null;
  const marker = '/_rels/';
  const markerIndex = part.lastIndexOf(marker);
  if (markerIndex <= 0 || !part.endsWith('.rels')) return undefined;
  const directory = part.slice(0, markerIndex);
  const filename = part.slice(markerIndex + marker.length, -5);
  return `${directory}/${filename}`;
}

export function xlsxActiveContent(name: string, contentType: string): boolean {
  const foldedName = name.toLowerCase();
  const foldedType = contentType.toLowerCase();
  return (
    foldedName.includes('vbaproject') ||
    foldedName.includes('/activex/') ||
    foldedName.includes('/embeddings/') ||
    foldedType.includes('macroenabled') ||
    foldedType.includes('activex') ||
    foldedType.includes('oleobject')
  );
}

export function xlsxOpaqueContent(name: string, contentType: string): boolean {
  if (name.endsWith('.rels')) return false;
  return !(
    name.endsWith('.xml') ||
    contentType.startsWith('image/') ||
    contentType === 'application/xml'
  );
}

export function xlsxLexicalCompare(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

export function xlsxPackagePartIsXml(
  name: string,
  contentType: string,
): boolean {
  return (
    name.endsWith('.xml') ||
    contentType === 'application/xml' ||
    contentType.endsWith('+xml')
  );
}

export function consumeXlsxGraphExpandedBytes(
  current: number,
  amount: number,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): number {
  const actual = current + amount;
  if (
    !Number.isSafeInteger(actual) ||
    actual > limits.maxTotalUncompressedBytes
  ) {
    throw new XlsxResourceLimitError(
      'maxTotalUncompressedBytes',
      actual,
      limits.maxTotalUncompressedBytes,
      part,
    );
  }
  return actual;
}

export function assertXlsxGraphRelationshipTargets(
  parts: readonly Pick<XlsxPackageGraphPart, 'name'>[],
  relationships: readonly XlsxPackageGraphRelationship[],
): void {
  const names = new Set(parts.map((part) => part.name));
  for (const relationship of relationships) {
    if (relationship.owner !== null && !names.has(relationship.owner)) {
      throw new XlsxWriteError(
        'relationship-graph-invalid',
        'XLSX relationship owner part is missing',
        { part: relationship.owner },
      );
    }
    if (relationship.mode === 'internal' && !names.has(relationship.target)) {
      throw new XlsxWriteError(
        'relationship-graph-invalid',
        'XLSX internal relationship target part is missing',
        { part: relationship.owner ?? '_rels/.rels' },
      );
    }
  }
}

export async function inspectXlsxPackageGraph(
  bytes: Uint8Array,
  limits: ResolvedXlsxResourceLimits,
): Promise<XlsxPackageGraph> {
  const archive = await JSZip.loadAsync(bytes);
  assertXlsxArchiveWithinLimits(archive, limits);
  const reader = new XlsxPartReader(archive, [], limits);
  const discovery = await discoverXlsxWorkbook(reader, limits);
  const names = Object.values(archive.files)
    .filter((entry) => !entry.dir)
    .map((entry) => entry.name)
    .sort();
  let expandedBytes = 0;
  const parts: XlsxPackageGraphPart[] = [];
  for (const name of names) {
    const entry = archive.file(name)!;
    const partBytes = await readZipEntryBytes(entry, limits.maxPartBytes);
    expandedBytes = consumeXlsxGraphExpandedBytes(
      expandedBytes,
      partBytes.byteLength,
      limits,
      name,
    );
    const relationshipPart = xlsxRelationshipOwner(name) !== undefined;
    const contentType =
      name === '[Content_Types].xml'
        ? 'application/vnd.openxmlformats-package.content-types+xml'
        : relationshipPart
          ? 'application/vnd.openxmlformats-package.relationships+xml'
          : discovery.contentTypes.contentTypeFor(name);
    if (contentType === undefined) {
      throw new XlsxWriteError(
        'relationship-graph-invalid',
        'XLSX package part has no declared content type',
        { part: name },
      );
    }
    parts.push({
      byteLength: partBytes.byteLength,
      contentType,
      name,
      relationshipPart,
      sha256: await sha256XlsxBytes(partBytes),
    });
  }

  const xmlValidator = new XlsxPartReader(archive, [], limits);
  for (const part of parts) {
    if (xlsxPackagePartIsXml(part.name, part.contentType)) {
      await xmlValidator.streamXml(part.name, {});
    }
  }

  const relationships: XlsxPackageGraphRelationship[] = [];
  for (const part of parts.filter((candidate) => candidate.relationshipPart)) {
    const owner = xlsxRelationshipOwner(part.name)!;
    const xml = (await reader.readXml(part.name))!;
    const parsed = parseXlsxRelationships(xml, owner, limits.maxRelationships);
    for (const relationship of parsed.values()) {
      relationships.push({ ...relationship, owner });
    }
  }
  relationships.sort(
    (left, right) =>
      xlsxLexicalCompare(String(left.owner), String(right.owner)) ||
      xlsxLexicalCompare(left.id, right.id),
  );
  assertXlsxGraphRelationshipTargets(parts, relationships);

  const containsActiveContent = parts.some((part) =>
    xlsxActiveContent(part.name, part.contentType),
  );
  const containsDigitalSignatures = parts.some((part) =>
    part.name.toLowerCase().startsWith('_xmlsignatures/'),
  );
  const containsExternalRelationships = relationships.some(
    (relationship) => relationship.mode === 'external',
  );
  const containsOpaqueContent = parts.some((part) =>
    xlsxOpaqueContent(part.name, part.contentType),
  );
  const manifestHash = await canonicalXlsxSha256({ parts, relationships });
  return {
    conformance: discovery.dialect,
    containsActiveContent,
    containsDigitalSignatures,
    containsExternalRelationships,
    containsOpaqueContent,
    manifestHash,
    parts,
    relationships,
  };
}
