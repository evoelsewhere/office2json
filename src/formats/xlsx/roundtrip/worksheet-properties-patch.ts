import { XlsxWriteError } from './errors';
import { xlsxMatchingCloseToken } from './hyperlink-patch';
import type { ResolvedXlsxWriteLimits } from './types';
import {
  decodeXlsxXml,
  encodeXlsxXml,
  tokenizeXlsxXml,
  xlsxXmlLocalName,
  type XlsxXmlTagToken,
} from './worksheet-patch';
import { writeLimitFailure } from './write-limits';

interface TextPatch {
  end: number;
  replacement: string;
  start: number;
}

export type XlsxWorksheetPropertyPatch =
  | {
      height?: number;
      hidden?: boolean;
      kind: 'set-row';
      operationId: string;
      row: number;
    }
  | {
      end: number;
      hidden?: boolean;
      kind: 'set-column';
      operationId: string;
      start: number;
      width?: number;
    };

export interface XlsxWorksheetPropertyPatchResult {
  data: Uint8Array;
  patchBytes: number;
  patchCount: number;
}

function failure(
  message: string,
  part: string,
  request?: XlsxWorksheetPropertyPatch,
  featureClass = 'worksheet-property-xml',
): never {
  throw new XlsxWriteError('preservation-conflict', message, {
    featureClass,
    ...(request === undefined
      ? {}
      : {
          operationId: request.operationId,
          range:
            request.kind === 'set-row'
              ? String(request.row)
              : `${request.start}:${request.end}`,
        }),
    part,
  });
}

function directCollection(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  name: string,
): XlsxXmlTagToken | undefined {
  const prefix = root.name.slice(0, -xlsxXmlLocalName(root.name).length);
  return tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}${name}`,
  );
}

function replacement(
  text: string,
  token: XlsxXmlTagToken,
  replaced: ReadonlySet<string>,
  additions: string,
): string {
  const authored = token.attributes
    .filter((attribute) => !replaced.has(attribute.name))
    .map((attribute) => text.slice(attribute.start, attribute.end))
    .join('');
  return `<${token.name}${authored}${additions}${token.selfClosing ? '/>' : '>'}`;
}

function rowReplacement(
  text: string,
  token: XlsxXmlTagToken,
  request: Extract<XlsxWorksheetPropertyPatch, { kind: 'set-row' }>,
): string {
  const replaced = new Set<string>();
  let additions = '';
  if (request.height !== undefined) {
    replaced.add('customHeight');
    replaced.add('ht');
    additions += ` ht="${request.height}" customHeight="1"`;
  }
  if (request.hidden !== undefined) {
    replaced.add('hidden');
    additions += ` hidden="${request.hidden ? '1' : '0'}"`;
  }
  return replacement(text, token, replaced, additions);
}

function columnReplacement(
  text: string,
  token: XlsxXmlTagToken,
  request: Extract<XlsxWorksheetPropertyPatch, { kind: 'set-column' }>,
): string {
  const replaced = new Set<string>();
  let additions = '';
  if (request.width !== undefined) {
    replaced.add('customWidth');
    replaced.add('width');
    additions += ` width="${request.width}" customWidth="1"`;
  }
  if (request.hidden !== undefined) {
    replaced.add('hidden');
    additions += ` hidden="${request.hidden ? '1' : '0'}"`;
  }
  return replacement(text, token, replaced, additions);
}

function attributeValue(
  token: XlsxXmlTagToken,
  name: string,
): string | undefined {
  return token.attributes.find((attribute) => attribute.name === name)?.value;
}

export function patchXlsxWorksheetProperties(
  bytes: Uint8Array,
  requested: readonly XlsxWorksheetPropertyPatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxWorksheetPropertyPatchResult {
  if (requested.length > limits.maxPatchCount) {
    writeLimitFailure(
      'maxPatchCount',
      requested.length,
      limits.maxPatchCount,
      part,
    );
  }
  if (requested.length === 0) {
    return { data: bytes.slice(), patchBytes: 0, patchCount: 0 };
  }
  const decoded = decodeXlsxXml(bytes, part);
  const tokens = tokenizeXlsxXml(decoded.text, part);
  const root = tokens.find(
    (token) =>
      token.depth === 0 && xlsxXmlLocalName(token.name) === 'worksheet',
  );
  if (!root) failure('XLSX worksheet root cannot patch properties', part);
  const prefix = root.name.slice(0, -'worksheet'.length);
  const sheetData = directCollection(tokens, root, 'sheetData');
  const columns = directCollection(tokens, root, 'cols');
  const patches: TextPatch[] = [];
  const targets = new Set<string>();
  for (const request of requested) {
    const key =
      request.kind === 'set-row'
        ? `row:${request.row}`
        : `column:${request.start}:${request.end}`;
    if (targets.has(key)) {
      failure(
        'XLSX worksheet property patch targets must be unique',
        part,
        request,
      );
    }
    targets.add(key);
    const collection = request.kind === 'set-row' ? sheetData : columns;
    if (!collection) {
      failure(
        request.kind === 'set-row'
          ? 'XLSX target row has no safe explicit XML span'
          : 'XLSX target column range has no safe explicit XML span',
        part,
        request,
        request.kind === 'set-row' ? 'missing-row-span' : 'missing-column-span',
      );
    }
    const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(collection));
    const matches = tokens.filter((token) => {
      if (
        token.closing ||
        token.start < collection.end ||
        token.end > close.start ||
        token.depth !== collection.depth + 1
      ) {
        return false;
      }
      if (request.kind === 'set-row') {
        return (
          token.name === `${prefix}row` &&
          attributeValue(token, 'r') === String(request.row)
        );
      }
      return (
        token.name === `${prefix}col` &&
        attributeValue(token, 'min') === String(request.start) &&
        attributeValue(token, 'max') === String(request.end)
      );
    });
    if (matches.length !== 1) {
      failure(
        request.kind === 'set-row'
          ? 'XLSX target row has no unique safe XML span'
          : 'XLSX target column range has no unique safe XML span',
        part,
        request,
        request.kind === 'set-row' ? 'missing-row-span' : 'missing-column-span',
      );
    }
    const token = matches[0]!;
    patches.push({
      end: token.end,
      replacement:
        request.kind === 'set-row'
          ? rowReplacement(decoded.text, token, request)
          : columnReplacement(decoded.text, token, request),
      start: token.start,
    });
  }
  let patchBytes = 0;
  for (const patch of patches) {
    patchBytes += encodeXlsxXml({
      bom: false,
      encoding: decoded.encoding,
      text: patch.replacement,
    }).byteLength;
    if (patchBytes > limits.maxPatchBytes) {
      writeLimitFailure(
        'maxPatchBytes',
        patchBytes,
        limits.maxPatchBytes,
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
