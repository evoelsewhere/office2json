import { XlsxWriteError } from './errors';
import { xlsxMatchingCloseToken } from './hyperlink-patch';
import {
  transformXlsxStructuralDrawingAnchor,
  type XlsxStructuralDrawingAnchor,
} from './structural-reference';
import type { ResolvedXlsxWriteLimits } from './types';
import {
  decodeXlsxXml,
  encodeXlsxXml,
  tokenizeXlsxXml,
  xlsxXmlLocalName,
  type XlsxXmlTagToken,
} from './worksheet-patch';
import type { XlsxWorksheetStructurePatch } from './worksheet-structure-patch';
import { writeLimitFailure } from './write-limits';

interface TextPatch {
  end: number;
  replacement: string;
  start: number;
}

export interface XlsxDrawingStructurePatchResult {
  data: Uint8Array;
  patchBytes: number;
  patchCount: number;
}

function failure(
  message: string,
  part: string,
  request: XlsxWorksheetStructurePatch,
  featureClass = 'drawing-structure-xml',
): never {
  throw new XlsxWriteError('preservation-conflict', message, {
    featureClass,
    operationId: request.operationId,
    part,
    range: `${request.index}:${request.index + request.count - 1}`,
  });
}

function unsignedText(
  text: string,
  part: string,
  request: XlsxWorksheetStructurePatch,
): number {
  const value = text.trim();
  if (!/^(?:0|[1-9]\d*)$/u.test(value)) {
    failure('XLSX structural drawing marker is invalid', part, request);
  }
  return Number(value);
}

function marker(
  text: string,
  tokens: readonly XlsxXmlTagToken[],
  source: XlsxXmlTagToken,
  prefix: string,
  part: string,
  request: XlsxWorksheetStructurePatch,
): {
  marker: {
    column: number;
    columnOffset: number;
    row: number;
    rowOffset: number;
  };
  column: { close: XlsxXmlTagToken; open: XlsxXmlTagToken };
  row: { close: XlsxXmlTagToken; open: XlsxXmlTagToken };
} {
  const sourceIndex = tokens.indexOf(source);
  const sourceClose = xlsxMatchingCloseToken(tokens, sourceIndex);
  const children = tokens.slice(sourceIndex, tokens.indexOf(sourceClose));
  const field = (name: 'col' | 'row') => {
    const open = children.find(
      (token) =>
        !token.closing &&
        token.depth === source.depth + 1 &&
        token.name === `${prefix}${name}`,
    );
    if (!open) {
      failure('XLSX structural drawing marker is invalid', part, request);
    }
    const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(open));
    return { close, open };
  };
  const column = field('col');
  const row = field('row');
  return {
    column,
    marker: {
      column:
        unsignedText(
          text.slice(column.open.end, column.close.start),
          part,
          request,
        ) + 1,
      columnOffset: 0,
      row:
        unsignedText(text.slice(row.open.end, row.close.start), part, request) +
        1,
      rowOffset: 0,
    },
    row,
  };
}

function drawingPatches(
  text: string,
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  part: string,
  request: XlsxWorksheetStructurePatch,
): TextPatch[] {
  const localRoot = xlsxXmlLocalName(root.name);
  if (localRoot !== 'wsDr') {
    failure('XLSX drawing root cannot patch structure', part, request);
  }
  const prefix = root.name.slice(0, root.name.indexOf(localRoot));
  const anchors = tokens.filter(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      (token.name === `${prefix}oneCellAnchor` ||
        token.name === `${prefix}twoCellAnchor`),
  );
  const patches: TextPatch[] = [];
  for (const anchor of anchors) {
    const local = xlsxXmlLocalName(anchor.name);
    const anchorIndex = tokens.indexOf(anchor);
    const anchorClose = xlsxMatchingCloseToken(tokens, anchorIndex);
    const children = tokens.slice(anchorIndex, tokens.indexOf(anchorClose));
    const fromToken = children.find(
      (token) =>
        !token.closing &&
        token.depth === anchor.depth + 1 &&
        token.name === `${prefix}from`,
    );
    if (!fromToken) {
      failure('XLSX structural drawing anchor is invalid', part, request);
    }
    const from = marker(text, tokens, fromToken, prefix, part, request);
    const kind = local === 'oneCellAnchor' ? 'one-cell' : 'two-cell';
    const editAsSource = anchor.attributes.find(
      (candidate) => candidate.name === 'editAs',
    )?.value;
    const editAs =
      editAsSource === undefined
        ? undefined
        : (
            {
              absolute: 'absolute',
              oneCell: 'one-cell',
              twoCell: 'two-cell',
            } as const
          )[editAsSource];
    if (editAsSource !== undefined && editAs === undefined) {
      failure('XLSX structural drawing edit mode is invalid', part, request);
    }
    let to: ReturnType<typeof marker> | undefined;
    if (kind === 'two-cell') {
      const toToken = children.find(
        (token) =>
          !token.closing &&
          token.depth === anchor.depth + 1 &&
          token.name === `${prefix}to`,
      );
      if (!toToken) {
        failure('XLSX structural drawing anchor is invalid', part, request);
      }
      to = marker(text, tokens, toToken, prefix, part, request);
    }
    const sourceAnchor: XlsxStructuralDrawingAnchor = {
      ...(editAs === undefined ? {} : { editAs }),
      from: from.marker,
      kind,
      ...(to === undefined ? {} : { to: to.marker }),
    };
    const transformed = transformXlsxStructuralDrawingAnchor(
      sourceAnchor,
      request,
    );
    if (transformed === null) {
      failure(
        'XLSX structural edit would delete a drawing anchor',
        part,
        request,
        'drawing-anchor-deletion',
      );
    }
    for (const [sourceMarker, outputMarker] of [
      [from, transformed.from!],
      ...(to === undefined ? [] : ([[to, transformed.to!]] as const)),
    ] as const) {
      for (const [field, previous, output] of [
        ['column', sourceMarker.marker.column, outputMarker.column],
        ['row', sourceMarker.marker.row, outputMarker.row],
      ] as const) {
        if (previous === output) continue;
        const target =
          field === 'column' ? sourceMarker.column : sourceMarker.row;
        patches.push({
          end: target.close.start,
          replacement: String(output - 1),
          start: target.open.end,
        });
      }
    }
  }
  return patches;
}

function patchOne(
  bytes: Uint8Array,
  request: XlsxWorksheetStructurePatch,
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxDrawingStructurePatchResult {
  const decoded = decodeXlsxXml(bytes, part);
  const tokens = tokenizeXlsxXml(decoded.text, part);
  const root = tokens[0];
  if (!root) failure('XLSX drawing root cannot patch structure', part, request);
  const patches = drawingPatches(decoded.text, tokens, root, part, request);
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

export function patchXlsxDrawingStructure(
  bytes: Uint8Array,
  requested: readonly XlsxWorksheetStructurePatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxDrawingStructurePatchResult {
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
