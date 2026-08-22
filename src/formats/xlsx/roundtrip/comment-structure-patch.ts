import { parseXlsxCellReference } from '../internal/cell-reference';
import { XlsxWriteError } from './errors';
import { xlsxMatchingCloseToken } from './hyperlink-patch';
import { transformXlsxStructuralCell } from './structural-reference';
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

export interface XlsxCommentStructurePatchResult {
  data: Uint8Array;
  patchBytes: number;
  patchCount: number;
}

function failure(
  message: string,
  part: string,
  request: XlsxWorksheetStructurePatch,
  featureClass = 'comment-structure-xml',
): never {
  throw new XlsxWriteError('preservation-conflict', message, {
    featureClass,
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

function transformedCommentReference(
  token: XlsxXmlTagToken,
  part: string,
  request: XlsxWorksheetStructurePatch,
): { previous: string; transformed: string } {
  const reference = attribute(token, 'ref');
  const parsed = parseXlsxCellReference(reference?.value);
  if (!reference || !parsed || parsed.absoluteColumn || parsed.absoluteRow) {
    failure('XLSX structural comment reference is invalid', part, request);
  }
  const transformed = transformXlsxStructuralCell(
    parsed.row,
    parsed.column,
    request,
  );
  if (transformed === null) {
    failure(
      'XLSX structural edit would delete a comment anchor',
      part,
      request,
      'comment-anchor-deletion',
    );
  }
  return { previous: parsed.address, transformed: transformed.address };
}

function commentPatches(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  part: string,
  request: XlsxWorksheetStructurePatch,
): TextPatch[] {
  const localRoot = xlsxXmlLocalName(root.name);
  const prefix = root.name.slice(0, -localRoot.length);
  let comments: XlsxXmlTagToken[];
  if (localRoot === 'ThreadedComments') {
    comments = tokens.filter(
      (token) =>
        !token.closing &&
        token.depth === root.depth + 1 &&
        token.name === `${prefix}threadedComment`,
    );
  } else if (localRoot === 'comments') {
    const commentList = tokens.find(
      (token) =>
        !token.closing &&
        token.depth === root.depth + 1 &&
        token.name === `${prefix}commentList`,
    );
    if (!commentList) {
      failure('XLSX comment list cannot patch structure', part, request);
    }
    const listIndex = tokens.indexOf(commentList);
    const close = xlsxMatchingCloseToken(tokens, listIndex);
    comments = tokens
      .slice(listIndex, tokens.indexOf(close))
      .filter(
        (token) =>
          !token.closing &&
          token.depth === commentList.depth + 1 &&
          token.name === `${prefix}comment`,
      );
  } else {
    failure('XLSX comment root cannot patch structure', part, request);
  }
  const patches: TextPatch[] = [];
  for (const comment of comments) {
    const reference = transformedCommentReference(comment, part, request);
    if (reference.transformed !== reference.previous) {
      patches.push(
        attributePatch(attribute(comment, 'ref')!, reference.transformed),
      );
    }
  }
  return patches;
}

function unsignedText(
  text: string,
  part: string,
  request: XlsxWorksheetStructurePatch,
): number {
  const value = text.trim();
  if (!/^(?:0|[1-9]\d*)$/u.test(value)) {
    failure('XLSX structural VML comment anchor is invalid', part, request);
  }
  return Number(value);
}

function vmlPatches(
  text: string,
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  part: string,
  request: XlsxWorksheetStructurePatch,
): TextPatch[] {
  const shapes = tokens.filter(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      xlsxXmlLocalName(token.name) === 'shape',
  );
  const patches: TextPatch[] = [];
  for (const shape of shapes) {
    const shapeIndex = tokens.indexOf(shape);
    const shapeClose = xlsxMatchingCloseToken(tokens, shapeIndex);
    const shapeTokens = tokens.slice(shapeIndex, tokens.indexOf(shapeClose));
    const client = shapeTokens.find(
      (token) =>
        !token.closing &&
        token.depth === shape.depth + 1 &&
        xlsxXmlLocalName(token.name) === 'ClientData' &&
        attribute(token, 'ObjectType')?.value === 'Note',
    );
    if (!client) continue;
    const clientIndex = tokens.indexOf(client);
    const clientClose = xlsxMatchingCloseToken(tokens, clientIndex);
    const clientTokens = tokens.slice(clientIndex, tokens.indexOf(clientClose));
    const rowToken = clientTokens.find(
      (token) =>
        !token.closing &&
        token.depth === client.depth + 1 &&
        xlsxXmlLocalName(token.name) === 'Row',
    );
    const columnToken = clientTokens.find(
      (token) =>
        !token.closing &&
        token.depth === client.depth + 1 &&
        xlsxXmlLocalName(token.name) === 'Column',
    );
    if (!rowToken || !columnToken) {
      failure('XLSX structural VML comment anchor is invalid', part, request);
    }
    const rowClose = xlsxMatchingCloseToken(tokens, tokens.indexOf(rowToken));
    const columnClose = xlsxMatchingCloseToken(
      tokens,
      tokens.indexOf(columnToken),
    );
    const row = unsignedText(
      text.slice(rowToken.end, rowClose.start),
      part,
      request,
    );
    const column = unsignedText(
      text.slice(columnToken.end, columnClose.start),
      part,
      request,
    );
    const transformed = transformXlsxStructuralCell(
      row + 1,
      column + 1,
      request,
    );
    if (transformed === null) {
      failure(
        'XLSX structural edit would delete a comment anchor',
        part,
        request,
        'comment-anchor-deletion',
      );
    }
    if (transformed.row !== row + 1) {
      patches.push({
        end: rowClose.start,
        replacement: String(transformed.row - 1),
        start: rowToken.end,
      });
    }
    if (transformed.column !== column + 1) {
      patches.push({
        end: columnClose.start,
        replacement: String(transformed.column - 1),
        start: columnToken.end,
      });
    }
  }
  return patches;
}

function patchOne(
  bytes: Uint8Array,
  request: XlsxWorksheetStructurePatch,
  limits: ResolvedXlsxWriteLimits,
  part: string,
  kind: 'comments' | 'vml',
): XlsxCommentStructurePatchResult {
  const decoded = decodeXlsxXml(bytes, part);
  const tokens = tokenizeXlsxXml(decoded.text, part);
  const root = tokens[0];
  if (!root) failure('XLSX comment root cannot patch structure', part, request);
  const patches =
    kind === 'vml'
      ? vmlPatches(decoded.text, tokens, root, part, request)
      : commentPatches(tokens, root, part, request);
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

function patchRequested(
  bytes: Uint8Array,
  requested: readonly XlsxWorksheetStructurePatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
  kind: 'comments' | 'vml',
): XlsxCommentStructurePatchResult {
  let data: Uint8Array = bytes.slice();
  let patchBytes = 0;
  let patchCount = 0;
  for (const request of requested) {
    const result = patchOne(data, request, limits, part, kind);
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

export function patchXlsxCommentAnchors(
  bytes: Uint8Array,
  requested: readonly XlsxWorksheetStructurePatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxCommentStructurePatchResult {
  return patchRequested(bytes, requested, limits, part, 'comments');
}

export function patchXlsxCommentVmlAnchors(
  bytes: Uint8Array,
  requested: readonly XlsxWorksheetStructurePatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxCommentStructurePatchResult {
  return patchRequested(bytes, requested, limits, part, 'vml');
}
