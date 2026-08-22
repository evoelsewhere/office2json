import { decodeXmlEntities } from '../../../common/text/html';
import { XlsxWriteError } from './errors';
import {
  transformXlsxStructuralSourceFormula,
  xlsxStructuralSourceFormulaArea,
} from './formula-reference';
import { xlsxMatchingCloseToken } from './hyperlink-patch';
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

function chartNamespaceSupported(namespace: string): boolean {
  return (
    namespace === 'http://purl.oclc.org/ooxml/drawingml/chart' ||
    namespace === 'http://schemas.openxmlformats.org/drawingml/2006/chart'
  );
}

function chartReferenceElement(localName: string): boolean {
  return (
    localName === 'multiLvlStrRef' ||
    localName === 'numRef' ||
    localName === 'strRef'
  );
}

function chartSourceContainer(localName: string): boolean {
  return (
    localName === 'bubbleSize' ||
    localName === 'cat' ||
    localName === 'tx' ||
    localName === 'val' ||
    localName === 'xVal' ||
    localName === 'yVal'
  );
}

export interface XlsxChartStructurePatchResult {
  data: Uint8Array;
  patchBytes: number;
  patchCount: number;
}

function failure(
  code: 'formula-rewrite-unsupported' | 'preservation-conflict',
  message: string,
  part: string,
  request: XlsxWorksheetStructurePatch,
  featureClass: string,
): never {
  throw new XlsxWriteError(code, message, {
    featureClass,
    operationId: request.operationId,
    part,
    range: `${request.index}:${request.index + request.count - 1}`,
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

function parentToken(
  tokens: readonly XlsxXmlTagToken[],
  child: XlsxXmlTagToken,
): XlsxXmlTagToken | undefined {
  const childIndex = tokens.indexOf(child);
  return tokens
    .slice(0, childIndex)
    .reverse()
    .find((candidate) => candidate.depth === child.depth - 1);
}

function directText(
  text: string,
  tokens: readonly XlsxXmlTagToken[],
  open: XlsxXmlTagToken,
  part: string,
  request: XlsxWorksheetStructurePatch,
): { end: number; source: string; start: number } {
  if (open.selfClosing) {
    failure(
      'preservation-conflict',
      'XLSX structural chart formula is invalid',
      part,
      request,
      'chart-structure-xml',
    );
  }
  const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(open));
  if (
    tokens.some(
      (token) =>
        !token.closing && token.start >= open.end && token.end <= close.start,
    )
  ) {
    failure(
      'preservation-conflict',
      'XLSX structural chart formula is invalid',
      part,
      request,
      'chart-structure-xml',
    );
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

function chartPatches(
  text: string,
  tokens: readonly XlsxXmlTagToken[],
  request: XlsxWorksheetStructurePatch,
  part: string,
  sheetName: string,
): TextPatch[] {
  const root = tokens.find(
    (token) =>
      token.depth === 0 && xlsxXmlLocalName(token.name) === 'chartSpace',
  );
  if (!root) {
    failure(
      'preservation-conflict',
      'XLSX chart root cannot patch structure',
      part,
      request,
      'chart-structure-xml',
    );
  }
  const contexts = namespaceContexts(tokens);
  const chartNamespace = tokenNamespace(root, contexts);
  if (
    chartNamespace === undefined ||
    !chartNamespaceSupported(chartNamespace)
  ) {
    failure(
      'preservation-conflict',
      'XLSX chart namespace cannot patch structure',
      part,
      request,
      'chart-structure-xml',
    );
  }
  const formulas = tokens.filter((token) => {
    if (
      token.closing ||
      xlsxXmlLocalName(token.name) !== 'f' ||
      tokenNamespace(token, contexts) !== chartNamespace
    ) {
      return false;
    }
    const parent = parentToken(tokens, token);
    const grandparent =
      parent === undefined ? undefined : parentToken(tokens, parent);
    return (
      parent !== undefined &&
      grandparent !== undefined &&
      tokenNamespace(parent, contexts) === chartNamespace &&
      tokenNamespace(grandparent, contexts) === chartNamespace &&
      chartReferenceElement(xlsxXmlLocalName(parent.name)) &&
      chartSourceContainer(xlsxXmlLocalName(grandparent.name))
    );
  });
  const patches: TextPatch[] = [];
  for (const token of formulas) {
    const formula = directText(text, tokens, token, part, request);
    const transformed = transformXlsxStructuralSourceFormula(
      formula.source,
      sheetName,
      sheetName,
      request,
    );
    if (transformed.kind === 'unsupported') {
      failure(
        'formula-rewrite-unsupported',
        'XLSX structural chart formula is unsupported',
        part,
        request,
        'chart-formula-reference',
      );
    }
    if (transformed.kind === 'deleted') {
      failure(
        'formula-rewrite-unsupported',
        'XLSX structural edit would delete a chart source',
        part,
        request,
        'chart-source-deletion',
      );
    }
    if (transformed.kind === 'preserved') continue;
    if (
      xlsxStructuralSourceFormulaArea(formula.source) !==
      xlsxStructuralSourceFormulaArea(transformed.expression)
    ) {
      failure(
        'preservation-conflict',
        'XLSX structural chart cache cardinality would change',
        part,
        request,
        'chart-cache-cardinality',
      );
    }
    patches.push({
      end: formula.end,
      replacement: escapeXmlText(transformed.expression),
      start: formula.start,
    });
  }
  return patches;
}

function patchOne(
  bytes: Uint8Array,
  request: XlsxWorksheetStructurePatch,
  limits: ResolvedXlsxWriteLimits,
  part: string,
  sheetName: string,
): XlsxChartStructurePatchResult {
  const decoded = decodeXlsxXml(bytes, part);
  const tokens = tokenizeXlsxXml(decoded.text, part);
  const patches = chartPatches(decoded.text, tokens, request, part, sheetName);
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

export function patchXlsxChartStructure(
  bytes: Uint8Array,
  requested: readonly XlsxWorksheetStructurePatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
  sheetName: string,
): XlsxChartStructurePatchResult {
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
