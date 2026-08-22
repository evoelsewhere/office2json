import { decodeXmlEntities } from '../../../common/text/html';
import { parseXlsxCellReference } from '../internal/cell-reference';
import { XlsxWriteError } from './errors';
import { transformXlsxStructuralSourceFormula } from './formula-reference';
import { xlsxMatchingCloseToken } from './hyperlink-patch';
import { transformXlsxStructuralCell } from './structural-reference';
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

const X14_NAMESPACE =
  'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM_NAMESPACE = 'http://schemas.microsoft.com/office/excel/2006/main';
const SPARKLINE_EXTENSION_URI = '{05c60535-1f16-4fd2-b633-f4f36f0b64e0}';

interface TextPatch {
  end: number;
  replacement: string;
  start: number;
}

export interface XlsxSparklineStructurePatchResult {
  data: Uint8Array;
  patchBytes: number;
  patchCount: number;
}

function requestRange(request: XlsxWorksheetStructurePatch): string {
  return `${request.index}:${request.index + request.count - 1}`;
}

function failure(
  message: string,
  part: string,
  request: XlsxWorksheetStructurePatch,
  featureClass = 'sparkline-structure-xml',
): never {
  throw new XlsxWriteError('preservation-conflict', message, {
    featureClass,
    operationId: request.operationId,
    part,
    range: requestRange(request),
  });
}

function formulaFailure(
  message: string,
  part: string,
  request: XlsxWorksheetStructurePatch,
  featureClass: string,
): never {
  throw new XlsxWriteError('formula-rewrite-unsupported', message, {
    featureClass,
    operationId: request.operationId,
    part,
    range: requestRange(request),
  });
}

function namespaceContexts(
  tokens: readonly XlsxXmlTagToken[],
): ReadonlyMap<XlsxXmlTagToken, ReadonlyMap<string, string>> {
  const output = new Map<XlsxXmlTagToken, ReadonlyMap<string, string>>();
  const stack: Array<ReadonlyMap<string, string>> = [];
  for (const token of tokens) {
    const context = new Map(stack[token.depth - 1]);
    for (const attribute of token.attributes) {
      if (attribute.name === 'xmlns') context.set('', attribute.value);
      else if (attribute.name.startsWith('xmlns:')) {
        context.set(attribute.name.slice('xmlns:'.length), attribute.value);
      }
    }
    output.set(token, context);
    stack[token.depth] = context;
    stack.length = token.depth + 1;
  }
  return output;
}

function tokenNamespace(
  token: XlsxXmlTagToken,
  contexts: ReadonlyMap<XlsxXmlTagToken, ReadonlyMap<string, string>>,
): string | undefined {
  const prefix = token.name
    .slice(0, token.name.indexOf(':') + 1)
    .replace(':', '');
  return contexts.get(token)?.get(prefix);
}

function directChildren(
  tokens: readonly XlsxXmlTagToken[],
  parent: XlsxXmlTagToken,
  contexts: ReadonlyMap<XlsxXmlTagToken, ReadonlyMap<string, string>>,
  namespace: string,
  localName: string,
): XlsxXmlTagToken[] {
  const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(parent));
  return tokens.filter(
    (token) =>
      !token.closing &&
      token.depth === parent.depth + 1 &&
      token.start >= parent.end &&
      token.end <= close.start &&
      xlsxXmlLocalName(token.name) === localName &&
      tokenNamespace(token, contexts) === namespace,
  );
}

function onlyChild(
  tokens: readonly XlsxXmlTagToken[],
  parent: XlsxXmlTagToken,
  contexts: ReadonlyMap<XlsxXmlTagToken, ReadonlyMap<string, string>>,
  namespace: string,
  localName: string,
  part: string,
  request: XlsxWorksheetStructurePatch,
): XlsxXmlTagToken {
  const matches = directChildren(
    tokens,
    parent,
    contexts,
    namespace,
    localName,
  );
  if (matches.length !== 1) {
    failure('XLSX structural sparkline graph is invalid', part, request);
  }
  return matches[0]!;
}

function directText(
  text: string,
  tokens: readonly XlsxXmlTagToken[],
  open: XlsxXmlTagToken,
  part: string,
  request: XlsxWorksheetStructurePatch,
): { end: number; source: string; start: number } {
  if (open.selfClosing) {
    failure('XLSX structural sparkline text is invalid', part, request);
  }
  const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(open));
  if (
    tokens.some(
      (token) =>
        !token.closing && token.start >= open.end && token.end <= close.start,
    )
  ) {
    failure('XLSX structural sparkline text is invalid', part, request);
  }
  return {
    end: close.start,
    source: decodeXmlEntities(text.slice(open.end, close.start)),
    start: open.end,
  };
}

function escapeXmlText(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('\r', '&#13;');
}

function sparklinePatches(
  text: string,
  tokens: readonly XlsxXmlTagToken[],
  request: XlsxWorksheetStructurePatch,
  part: string,
  sheetName: string,
): TextPatch[] {
  const root = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === 0 &&
      xlsxXmlLocalName(token.name) === 'worksheet',
  );
  if (!root) {
    failure('XLSX worksheet root cannot patch sparklines', part, request);
  }
  const contexts = namespaceContexts(tokens);
  const worksheetNamespace = tokenNamespace(root, contexts);
  if (worksheetNamespace === undefined) {
    failure('XLSX worksheet namespace cannot patch sparklines', part, request);
  }
  const extensionLists = directChildren(
    tokens,
    root,
    contexts,
    worksheetNamespace,
    'extLst',
  );
  if (extensionLists.length === 0) return [];
  if (extensionLists.length !== 1) {
    failure('XLSX structural sparkline graph is invalid', part, request);
  }
  const extensions = directChildren(
    tokens,
    extensionLists[0]!,
    contexts,
    worksheetNamespace,
    'ext',
  ).filter(
    (token) =>
      token.attributes
        .find((attribute) => attribute.name === 'uri')
        ?.value.toLowerCase() === SPARKLINE_EXTENSION_URI,
  );
  if (extensions.length === 0) return [];
  if (extensions.length !== 1) {
    failure('XLSX structural sparkline graph is invalid', part, request);
  }
  const groupsContainer = onlyChild(
    tokens,
    extensions[0]!,
    contexts,
    X14_NAMESPACE,
    'sparklineGroups',
    part,
    request,
  );
  const groups = directChildren(
    tokens,
    groupsContainer,
    contexts,
    X14_NAMESPACE,
    'sparklineGroup',
  );
  const patches: TextPatch[] = [];
  for (const group of groups) {
    const entriesContainer = onlyChild(
      tokens,
      group,
      contexts,
      X14_NAMESPACE,
      'sparklines',
      part,
      request,
    );
    const entries = directChildren(
      tokens,
      entriesContainer,
      contexts,
      X14_NAMESPACE,
      'sparkline',
    );
    if (entries.length === 0) {
      failure('XLSX structural sparkline graph is invalid', part, request);
    }
    for (const entry of entries) {
      const formula = directText(
        text,
        tokens,
        onlyChild(tokens, entry, contexts, XM_NAMESPACE, 'f', part, request),
        part,
        request,
      );
      const location = directText(
        text,
        tokens,
        onlyChild(
          tokens,
          entry,
          contexts,
          XM_NAMESPACE,
          'sqref',
          part,
          request,
        ),
        part,
        request,
      );
      const parsedLocation = parseXlsxCellReference(location.source);
      if (
        !parsedLocation ||
        parsedLocation.absoluteColumn ||
        parsedLocation.absoluteRow
      ) {
        failure('XLSX structural sparkline location is invalid', part, request);
      }
      const transformedLocation = transformXlsxStructuralCell(
        parsedLocation.row,
        parsedLocation.column,
        request,
      );
      if (transformedLocation === null) {
        failure(
          'XLSX structural edit would delete a sparkline location',
          part,
          request,
          'sparkline-location-deletion',
        );
      }
      const transformedFormula = transformXlsxStructuralSourceFormula(
        formula.source,
        sheetName,
        sheetName,
        request,
      );
      if (transformedFormula.kind === 'unsupported') {
        formulaFailure(
          'XLSX structural sparkline formula is unsupported',
          part,
          request,
          'sparkline-formula-reference',
        );
      }
      if (transformedFormula.kind === 'deleted') {
        formulaFailure(
          'XLSX structural edit would delete a sparkline source',
          part,
          request,
          'sparkline-source-deletion',
        );
      }
      if (transformedFormula.expression !== formula.source) {
        patches.push({
          end: formula.end,
          replacement: escapeXmlText(transformedFormula.expression),
          start: formula.start,
        });
      }
      if (transformedLocation.address !== location.source) {
        patches.push({
          end: location.end,
          replacement: transformedLocation.address,
          start: location.start,
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
  sheetName: string,
): XlsxSparklineStructurePatchResult {
  const decoded = decodeXlsxXml(bytes, part);
  const tokens = tokenizeXlsxXml(decoded.text, part);
  const patches = sparklinePatches(
    decoded.text,
    tokens,
    request,
    part,
    sheetName,
  );
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

export function patchXlsxSparklineStructure(
  bytes: Uint8Array,
  requested: readonly XlsxWorksheetStructurePatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
  sheetName: string,
): XlsxSparklineStructurePatchResult {
  let data: Uint8Array = bytes.slice();
  let patchBytes = 0;
  let patchCount = 0;
  for (const request of requested) {
    const result = patchOne(data, request, limits, part, sheetName);
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
