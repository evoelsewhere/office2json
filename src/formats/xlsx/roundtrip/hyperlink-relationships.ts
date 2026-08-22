import { XLSX_RELATIONSHIPS_NAMESPACE } from '../internal/relationships';
import type { ResolvedXlsxResourceLimits } from '../internal/resource-limits';
import type { XlsxHyperlink } from '../types';
import { XlsxWriteError } from './errors';
import {
  readXlsxHyperlinkRelationshipIds,
  xlsxMatchingCloseToken,
} from './hyperlink-patch';
import type { XlsxPackageGraphRelationship } from './internal/package-graph';
import type { ResolvedXlsxWriteLimits } from './types';
import {
  decodeXlsxXml,
  encodeXlsxXml,
  tokenizeXlsxXml,
  xlsxXmlLocalName,
} from './worksheet-patch';
import { writeLimitFailure } from './write-limits';

interface TextPatch {
  end: number;
  replacement: string;
  start: number;
}

export interface XlsxExternalHyperlinkRelationshipPlan {
  changed: boolean;
  idsByCell: ReadonlyMap<string, string>;
  removeIds: ReadonlySet<string>;
  targets: ReadonlyMap<string, string>;
}

export interface XlsxHyperlinkRelationshipPatchResult {
  data: Uint8Array;
  patchBytes: number;
  patchCount: number;
}

function escapeAttribute(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('"', '&quot;')
    .replaceAll('\r', '&#13;')
    .replaceAll('\n', '&#10;')
    .replaceAll('\t', '&#9;');
}

function nextRelationshipId(used: Set<string>): string {
  let maximum = 0;
  for (const id of used) {
    const match = /^rId([1-9]\d*)$/u.exec(id);
    if (match) maximum = Math.max(maximum, Number(match[1]));
  }
  if (!Number.isSafeInteger(maximum) || maximum >= Number.MAX_SAFE_INTEGER) {
    throw new XlsxWriteError(
      'identifier-allocation-failed',
      'XLSX relationship IDs are exhausted',
      { featureClass: 'relationship-id' },
    );
  }
  maximum += 1;
  const candidate = `rId${maximum}`;
  used.add(candidate);
  return candidate;
}

export function planXlsxExternalHyperlinkRelationships(
  worksheetBytes: Uint8Array,
  relationships: readonly XlsxPackageGraphRelationship[],
  hyperlinks: readonly XlsxHyperlink[],
  part: string,
): XlsxExternalHyperlinkRelationshipPlan {
  const sourceIds = readXlsxHyperlinkRelationshipIds(worksheetBytes, part);
  const sourceById = new Map(relationships.map((value) => [value.id, value]));
  const used = new Set(relationships.map((value) => value.id));
  const idsByCell = new Map<string, string>();
  const targets = new Map<string, string>();
  for (const hyperlink of hyperlinks) {
    if (hyperlink.target.kind !== 'external') continue;
    const sourceId = sourceIds.get(hyperlink.range.reference);
    const sourceRelationship =
      sourceId === undefined ? undefined : sourceById.get(sourceId);
    const id =
      sourceRelationship?.mode === 'external' &&
      sourceRelationship.type.endsWith('/hyperlink')
        ? sourceId!
        : nextRelationshipId(used);
    idsByCell.set(hyperlink.range.reference, id);
    targets.set(id, hyperlink.target.url);
  }
  const finalIds = new Set(idsByCell.values());
  const removeIds = new Set(
    [...sourceIds.values()].filter((id) => !finalIds.has(id)),
  );
  const changed =
    removeIds.size !== 0 ||
    [...targets].some(([id, target]) => sourceById.get(id)?.target !== target);
  return { changed, idsByCell, removeIds, targets };
}

function relationshipElement(
  name: string,
  id: string,
  target: string,
  relationshipType: string,
): string {
  return `<${name} Id="${escapeAttribute(id)}" Type="${relationshipType}" Target="${escapeAttribute(target)}" TargetMode="External"/>`;
}

function newRelationshipDocument(
  plan: XlsxExternalHyperlinkRelationshipPlan,
  relationshipType: string,
): Uint8Array {
  const items = [...plan.targets]
    .map(([id, target]) =>
      relationshipElement('Relationship', id, target, relationshipType),
    )
    .join('');
  return new TextEncoder().encode(
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${XLSX_RELATIONSHIPS_NAMESPACE}">${items}</Relationships>`,
  );
}

export function patchXlsxHyperlinkRelationships(
  bytes: Uint8Array | null,
  plan: XlsxExternalHyperlinkRelationshipPlan,
  relationshipType: string,
  writeLimits: ResolvedXlsxWriteLimits,
  readerLimits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxHyperlinkRelationshipPatchResult {
  if (bytes === null) {
    const data = newRelationshipDocument(plan, relationshipType);
    if (plan.targets.size > readerLimits.maxRelationships) {
      throw new XlsxWriteError(
        'resource-limit-exceeded',
        'XLSX generated relationships exceed the reader limit',
        {
          actual: plan.targets.size,
          limit: readerLimits.maxRelationships,
          limitName: 'maxRelationships',
          part,
        },
      );
    }
    if (data.byteLength > writeLimits.maxGeneratedXmlBytes) {
      writeLimitFailure(
        'maxGeneratedXmlBytes',
        data.byteLength,
        writeLimits.maxGeneratedXmlBytes,
        part,
      );
    }
    return {
      data,
      patchBytes: data.byteLength,
      patchCount: plan.targets.size === 0 ? 0 : 1,
    };
  }
  const decoded = decodeXlsxXml(bytes, part);
  const tokens = tokenizeXlsxXml(decoded.text, part);
  const root = tokens.find(
    (token) =>
      token.depth === 0 && xlsxXmlLocalName(token.name) === 'Relationships',
  );
  if (!root) {
    throw new XlsxWriteError(
      'preservation-conflict',
      'XLSX relationship root cannot be patched',
      { featureClass: 'relationship-xml', part },
    );
  }
  const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(root));
  const elementName = `${root.name.slice(0, -'Relationships'.length)}Relationship`;
  const existing = new Set<string>();
  const patches: TextPatch[] = [];
  let relationshipCount = 0;
  for (const token of tokens) {
    if (
      token.closing ||
      token.depth !== root.depth + 1 ||
      token.name !== elementName
    ) {
      continue;
    }
    relationshipCount += 1;
    const id = token.attributes.find(
      (attribute) => attribute.name === 'Id',
    )?.value;
    if (id === undefined) continue;
    existing.add(id);
    if (plan.removeIds.has(id)) {
      patches.push({ end: token.end, replacement: '', start: token.start });
      relationshipCount -= 1;
      continue;
    }
    const target = plan.targets.get(id);
    if (target === undefined) continue;
    const type = token.attributes.find(
      (attribute) => attribute.name === 'Type',
    )?.value;
    if (type !== relationshipType) continue;
    patches.push({
      end: token.end,
      replacement: relationshipElement(elementName, id, target, type),
      start: token.start,
    });
  }
  const additions = [...plan.targets].filter(([id]) => !existing.has(id));
  relationshipCount += additions.length;
  if (relationshipCount > readerLimits.maxRelationships) {
    throw new XlsxWriteError(
      'resource-limit-exceeded',
      'XLSX generated relationships exceed the reader limit',
      {
        actual: relationshipCount,
        limit: readerLimits.maxRelationships,
        limitName: 'maxRelationships',
        part,
      },
    );
  }
  if (additions.length !== 0) {
    const insertion = additions
      .map(([id, target]) =>
        relationshipElement(elementName, id, target, relationshipType),
      )
      .join('');
    if (root.selfClosing) {
      const raw = decoded.text
        .slice(root.start, root.end)
        .replace(/\/\s*>$/u, '>');
      patches.push({
        end: root.end,
        replacement: `${raw}${insertion}</${root.name}>`,
        start: root.start,
      });
    } else {
      patches.push({
        end: close.start,
        replacement: insertion,
        start: close.start,
      });
    }
  }
  let patchBytes = 0;
  for (const patch of patches) {
    patchBytes += encodeXlsxXml({
      bom: false,
      encoding: decoded.encoding,
      text: patch.replacement,
    }).byteLength;
    if (patchBytes > writeLimits.maxPatchBytes) {
      writeLimitFailure(
        'maxPatchBytes',
        patchBytes,
        writeLimits.maxPatchBytes,
        part,
      );
    }
  }
  patches.sort((left, right) => right.start - left.start);
  let output = decoded.text;
  for (const patch of patches) {
    output = `${output.slice(0, patch.start)}${patch.replacement}${output.slice(patch.end)}`;
  }
  const data = encodeXlsxXml({ ...decoded, text: output });
  if (data.byteLength > writeLimits.maxGeneratedXmlBytes) {
    writeLimitFailure(
      'maxGeneratedXmlBytes',
      data.byteLength,
      writeLimits.maxGeneratedXmlBytes,
      part,
    );
  }
  return { data, patchBytes, patchCount: patches.length };
}
