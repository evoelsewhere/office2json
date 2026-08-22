import { parseXlsxRangeReference } from '../internal/cell-reference';
import { XlsxWriteError } from './errors';
import { xlsxMatchingCloseToken } from './hyperlink-patch';
import { transformXlsxStructuralRange } from './structural-reference';
import type { ResolvedXlsxWriteLimits } from './types';
import {
  decodeXlsxXml,
  encodeXlsxXml,
  tokenizeXlsxXml,
  xlsxXmlLocalName,
  type XlsxXmlAttributeSpan,
  type XlsxXmlTagToken,
} from './worksheet-patch';
import type { XlsxWorksheetStructurePatch } from './worksheet-structure-patch';
import { writeLimitFailure } from './write-limits';

interface TextPatch {
  end: number;
  replacement: string;
  start: number;
}

export interface XlsxTableStructurePatchResult {
  data: Uint8Array;
  patchBytes: number;
  patchCount: number;
}

function failure(
  message: string,
  part: string,
  request: XlsxWorksheetStructurePatch,
): never {
  throw new XlsxWriteError('preservation-conflict', message, {
    featureClass: 'table-structure-xml',
    operationId: request.operationId,
    part,
    range: `${request.index}:${request.index + request.count - 1}`,
  });
}

function attribute(
  token: XlsxXmlTagToken,
  name: string,
): XlsxXmlAttributeSpan | undefined {
  return token.attributes.find((candidate) => candidate.name === name);
}

function attributePatch(span: XlsxXmlAttributeSpan, value: string): TextPatch {
  return {
    end: span.end,
    replacement: ` ${span.name}="${value}"`,
    start: span.start,
  };
}

function rangeAttribute(
  token: XlsxXmlTagToken,
  part: string,
  request: XlsxWorksheetStructurePatch,
  message: string,
): {
  range: NonNullable<ReturnType<typeof parseXlsxRangeReference>>;
  span: XlsxXmlAttributeSpan;
} {
  const span = attribute(token, 'ref');
  const range = parseXlsxRangeReference(span?.value);
  if (!span || !range) failure(message, part, request);
  return { range, span };
}

function tablePatches(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  request: XlsxWorksheetStructurePatch,
  part: string,
): TextPatch[] {
  const patches: TextPatch[] = [];
  const tableReference = rangeAttribute(
    root,
    part,
    request,
    'XLSX structural table range is invalid',
  );
  const transformedTable = transformXlsxStructuralRange(
    tableReference.range,
    request,
  );
  if (transformedTable === null) {
    failure('XLSX structural table range cannot be deleted', part, request);
  }
  if (transformedTable.reference !== tableReference.range.reference) {
    patches.push(
      attributePatch(tableReference.span, transformedTable.reference),
    );
  }

  const prefix = root.name.slice(0, -'table'.length);
  const autoFilter = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}autoFilter`,
  );
  if (!autoFilter) return patches;
  const filterClose = xlsxMatchingCloseToken(
    tokens,
    tokens.indexOf(autoFilter),
  );
  const filterReference = rangeAttribute(
    autoFilter,
    part,
    request,
    'XLSX structural table auto-filter range is invalid',
  );
  const transformedFilter = transformXlsxStructuralRange(
    filterReference.range,
    request,
  );
  if (transformedFilter === null) {
    patches.push({
      end: filterClose.end,
      replacement: '',
      start: autoFilter.start,
    });
    return patches;
  }
  if (transformedFilter.reference !== filterReference.range.reference) {
    patches.push(
      attributePatch(filterReference.span, transformedFilter.reference),
    );
  }
  const autoFilterIndex = tokens.indexOf(autoFilter);
  const sortState = tokens
    .slice(autoFilterIndex, tokens.indexOf(filterClose))
    .find(
      (token) =>
        !token.closing &&
        token.depth === autoFilter.depth + 1 &&
        token.name === `${prefix}sortState`,
    );
  if (!sortState) return patches;
  const sortClose = xlsxMatchingCloseToken(tokens, tokens.indexOf(sortState));
  const sortReference = rangeAttribute(
    sortState,
    part,
    request,
    'XLSX structural table sort range is invalid',
  );
  const transformedSort = transformXlsxStructuralRange(
    sortReference.range,
    request,
  );
  if (transformedSort === null) {
    patches.push({
      end: sortClose.end,
      replacement: '',
      start: sortState.start,
    });
    return patches;
  }
  if (transformedSort.reference !== sortReference.range.reference) {
    patches.push(attributePatch(sortReference.span, transformedSort.reference));
  }
  const sortStateIndex = tokens.indexOf(sortState);
  const conditions = tokens
    .slice(sortStateIndex, tokens.indexOf(sortClose))
    .filter(
      (token) =>
        !token.closing &&
        token.depth === sortState.depth + 1 &&
        token.name === `${prefix}sortCondition`,
    );
  for (const condition of conditions) {
    const conditionReference = rangeAttribute(
      condition,
      part,
      request,
      'XLSX structural table sort-condition range is invalid',
    );
    const transformed = transformXlsxStructuralRange(
      conditionReference.range,
      request,
    );
    if (transformed === null) {
      const conditionClose = xlsxMatchingCloseToken(
        tokens,
        tokens.indexOf(condition),
      );
      patches.push({
        end: conditionClose.end,
        replacement: '',
        start: condition.start,
      });
    } else if (transformed.reference !== conditionReference.range.reference) {
      patches.push(
        attributePatch(conditionReference.span, transformed.reference),
      );
    }
  }
  return patches;
}

function patchOne(
  bytes: Uint8Array,
  request: XlsxWorksheetStructurePatch,
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxTableStructurePatchResult {
  const decoded = decodeXlsxXml(bytes, part);
  const tokens = tokenizeXlsxXml(decoded.text, part);
  const root = tokens.find(
    (token) => token.depth === 0 && xlsxXmlLocalName(token.name) === 'table',
  );
  if (!root) failure('XLSX table root cannot patch structure', part, request);
  const patches = tablePatches(tokens, root, request, part);
  let patchBytes = 0;
  for (const patch of patches) {
    patchBytes += encodeXlsxXml({
      bom: false,
      encoding: decoded.encoding,
      text: patch.replacement,
    }).byteLength;
  }
  patches.sort((left, right) => right.start - left.start);
  let output = decoded.text;
  for (const patch of patches) {
    output = `${output.slice(0, patch.start)}${patch.replacement}${output.slice(patch.end)}`;
  }
  const data = encodeXlsxXml({ ...decoded, text: output });
  if (data.byteLength > limits.maxGeneratedXmlBytes) {
    writeLimitFailure(
      'maxGeneratedXmlBytes',
      data.byteLength,
      limits.maxGeneratedXmlBytes,
      part,
    );
  }
  return { data, patchBytes, patchCount: patches.length };
}

export function patchXlsxTableStructure(
  bytes: Uint8Array,
  requested: readonly XlsxWorksheetStructurePatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxTableStructurePatchResult {
  let data: Uint8Array = bytes.slice();
  let patchBytes = 0;
  let patchCount = 0;
  for (const request of requested) {
    const result = patchOne(data, request, limits, part);
    data = result.data;
    patchBytes += result.patchBytes;
    patchCount += result.patchCount;
    if (patchBytes > limits.maxPatchBytes) {
      writeLimitFailure(
        'maxPatchBytes',
        patchBytes,
        limits.maxPatchBytes,
        part,
      );
    }
    if (patchCount > limits.maxPatchCount) {
      writeLimitFailure(
        'maxPatchCount',
        patchCount,
        limits.maxPatchCount,
        part,
      );
    }
  }
  return { data, patchBytes, patchCount };
}
