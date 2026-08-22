import type { XlsxCell } from '../types';
import { XlsxWriteError } from './errors';
import type { ResolvedXlsxWriteLimits } from './types';
import { writeLimitFailure } from './write-limits';

export type XlsxXmlEncoding = 'utf-16be' | 'utf-16le' | 'utf-8';

export interface DecodedXlsxXml {
  bom: boolean;
  encoding: XlsxXmlEncoding;
  text: string;
}

export interface XlsxXmlAttributeSpan {
  end: number;
  name: string;
  start: number;
  value: string;
}

export interface XlsxXmlTagToken {
  attributes: XlsxXmlAttributeSpan[];
  closing: boolean;
  depth: number;
  end: number;
  name: string;
  selfClosing: boolean;
  start: number;
}

export interface XlsxWorksheetCellPatch {
  cell: XlsxCell;
  contentChanged?: boolean;
  operationId: string;
  xmlStyleIndex?: number;
}

export interface XlsxWorksheetPatchResult {
  data: Uint8Array;
  patchBytes: number;
  patchCount: number;
}

const ALLOWED_CELL_ATTRIBUTES = new Set(['cm', 'ph', 'r', 's', 't', 'vm']);
const INDEX_NOT_FOUND = -1;

function patchFailure(
  message: string,
  part: string,
  patch?: XlsxWorksheetCellPatch,
  featureClass?: string,
): never {
  throw new XlsxWriteError('preservation-conflict', message, {
    ...(patch === undefined
      ? {}
      : { cell: patch.cell.address, operationId: patch.operationId }),
    ...(featureClass === undefined ? {} : { featureClass }),
    part,
  });
}

function formulaFailure(
  message: string,
  part: string,
  patch: XlsxWorksheetCellPatch,
): never {
  throw new XlsxWriteError('formula-rewrite-unsupported', message, {
    cell: patch.cell.address,
    featureClass: 'formula-group',
    operationId: patch.operationId,
    part,
  });
}

export function decodeXlsxXml(bytes: Uint8Array, part: string): DecodedXlsxXml {
  let bom = false;
  let encoding: XlsxXmlEncoding;
  let offset = 0;
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    bom = true;
    encoding = 'utf-16le';
    offset = 2;
  } else if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    bom = true;
    encoding = 'utf-16be';
    offset = 2;
  } else if (bytes[0] === 0x3c && bytes[1] === 0x00) {
    encoding = 'utf-16le';
  } else if (bytes[0] === 0x00 && bytes[1] === 0x3c) {
    encoding = 'utf-16be';
  } else {
    encoding = 'utf-8';
    if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
      bom = true;
      offset = 3;
    }
  }
  try {
    return {
      bom,
      encoding,
      text: new TextDecoder(encoding, { fatal: true }).decode(
        bytes.subarray(offset),
      ),
    };
  } catch {
    patchFailure('XLSX worksheet XML encoding is invalid', part);
  }
}

function prepend(prefix: readonly number[], bytes: Uint8Array): Uint8Array {
  const output = new Uint8Array(prefix.length + bytes.byteLength);
  output.set(prefix);
  output.set(bytes, prefix.length);
  return output;
}

function encodeUtf16(text: string, littleEndian: boolean): Uint8Array {
  const bytes = new Uint8Array(text.length * 2);
  let offset = 0;
  const writeUnit = (code: number): void => {
    bytes[offset + (littleEndian ? 0 : 1)] = code & 0xff;
    bytes[offset + (littleEndian ? 1 : 0)] = code >>> 8;
    offset += 2;
  };
  for (const character of text) {
    writeUnit(character.charCodeAt(0));
    if (character.length === 2) writeUnit(character.charCodeAt(1));
  }
  return bytes;
}

export function encodeXlsxXml(value: DecodedXlsxXml): Uint8Array {
  if (value.encoding === 'utf-8') {
    const bytes = new TextEncoder().encode(value.text);
    return value.bom ? prepend([0xef, 0xbb, 0xbf], bytes) : bytes;
  }
  const bytes = encodeUtf16(value.text, value.encoding === 'utf-16le');
  if (!value.bom) return bytes;
  return value.encoding === 'utf-16le'
    ? prepend([0xff, 0xfe], bytes)
    : prepend([0xfe, 0xff], bytes);
}

function tokenEnd(text: string, start: number, part: string): number {
  const match = /^(?:(?:"[^"]*"|'[^']*'|[^'">])*)>/u.exec(text.slice(start));
  if (!match) patchFailure('XLSX worksheet XML token is unterminated', part);
  return start + match[0].length;
}

function markupTerminatorEnd(
  text: string,
  start: number,
  terminator: string,
  message: string,
  part: string,
): number {
  const end = text.indexOf(terminator, start);
  if (end === INDEX_NOT_FOUND) patchFailure(message, part);
  return end;
}

function skippedMarkupEnd(text: string, start: number, part: string): number {
  if (text.startsWith('<!--', start)) {
    return markupTerminatorEnd(
      text,
      start,
      '-->',
      'XLSX worksheet comment is unterminated',
      part,
    );
  }
  if (text.startsWith('<![CDATA[', start)) {
    return markupTerminatorEnd(
      text,
      start,
      ']]>',
      'XLSX worksheet CDATA is unterminated',
      part,
    );
  }
  if (text.startsWith('<?', start)) {
    return markupTerminatorEnd(
      text,
      start,
      '?>',
      'XLSX worksheet processing instruction is unterminated',
      part,
    );
  }
  if (text.startsWith('<!', start)) {
    patchFailure('XLSX worksheet declaration cannot be patched safely', part);
  }
  return start;
}

function skipWhitespace(text: string, start: number, end: number): number {
  const offset = text.slice(start, end).search(/\S/u);
  return start + offset;
}

function scanNameEnd(
  text: string,
  start: number,
  end: number,
  part: string,
): number {
  const offset = text.slice(start, end).search(/[^A-Za-z0-9_.:-]/u);
  if (offset === INDEX_NOT_FOUND) {
    patchFailure(
      'XLSX worksheet element or attribute name is unterminated',
      part,
    );
  }
  return start + offset;
}

function attributes(
  text: string,
  nameEnd: number,
  tagEnd: number,
  part: string,
): XlsxXmlAttributeSpan[] {
  const raw = text.slice(nameEnd, tagEnd);
  const output: XlsxXmlAttributeSpan[] = [];
  let cursor = 0;
  // Iteration count is the security bound; the character value is irrelevant.
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  for (const _character of raw) {
    const spanStart = cursor;
    cursor = skipWhitespace(raw, cursor, raw.length);
    if (raw[cursor] === '>' || raw[cursor] === '/') return output;
    const nameStart = cursor;
    cursor = scanNameEnd(raw, cursor, raw.length, part);
    if (cursor === nameStart) {
      patchFailure('XLSX worksheet attribute name is invalid', part);
    }
    const name = raw.slice(nameStart, cursor);
    cursor = skipWhitespace(raw, cursor, raw.length);
    if (raw[cursor] !== '=') {
      patchFailure('XLSX worksheet attribute assignment is invalid', part);
    }
    cursor += 1;
    cursor = skipWhitespace(raw, cursor, raw.length);
    const quote = raw[cursor];
    if (quote !== '"' && quote !== "'") {
      patchFailure('XLSX worksheet attribute quote is invalid', part);
    }
    cursor += 1;
    const valueStart = cursor;
    const valueEnd = raw.indexOf(quote, valueStart);
    cursor = valueEnd + 1;
    output.push({
      end: nameEnd + cursor,
      name,
      start: nameEnd + spanStart,
      value: raw.slice(valueStart, valueEnd),
    });
  }
  return output;
}

export function tokenizeXlsxXml(text: string, part: string): XlsxXmlTagToken[] {
  const tokens: XlsxXmlTagToken[] = [];
  const stack: string[] = [];
  let cursor = 0;
  // Iteration count is the security bound; the character value is irrelevant.
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  for (const _character of text) {
    const start = text.indexOf('<', cursor);
    if (start < 0) {
      cursor = text.length;
      continue;
    }
    const skipped = skippedMarkupEnd(text, start, part);
    if (skipped !== start) {
      cursor = skipped;
      continue;
    }
    const closing = text[start + 1] === '/';
    const nameStart = start + (closing ? 2 : 1);
    const first = text[nameStart];
    if (first === undefined || !/[A-Za-z_:]/u.test(first)) {
      patchFailure('XLSX worksheet element name is invalid', part);
    }
    const elementNameEnd = scanNameEnd(text, nameStart + 1, text.length, part);
    const name = text.slice(nameStart, elementNameEnd);
    const end = tokenEnd(text, elementNameEnd, part);
    const selfClosing =
      !closing && /\/\s*>$/u.test(text.slice(elementNameEnd, end));
    if (closing) {
      const expected = stack.pop();
      if (expected !== name) {
        patchFailure('XLSX worksheet element nesting is invalid', part);
      }
    }
    const depth = stack.length;
    tokens.push({
      attributes: closing ? [] : attributes(text, elementNameEnd, end, part),
      closing,
      depth,
      end,
      name,
      selfClosing,
      start,
    });
    if (!closing && !selfClosing) stack.push(name);
    cursor = end;
  }
  if (stack.length !== 0) {
    patchFailure('XLSX worksheet element is unclosed', part);
  }
  return tokens;
}

export function xlsxXmlLocalName(name: string): string {
  return name.slice(name.lastIndexOf(':') + 1);
}

function xlsxXmlCharacter(character: string): boolean {
  const code = character.codePointAt(0)!;
  return (
    code === 0x09 ||
    code === 0x0a ||
    code === 0x0d ||
    (code >= 0x20 && code <= 0xd7ff) ||
    (code >= 0xe000 && code <= 0xfffd) ||
    code >= 0x1_0000
  );
}

export function escapeXlsxCellText(
  value: string,
  patch: XlsxWorksheetCellPatch,
): string {
  let output = '';
  for (const character of value) {
    if (!xlsxXmlCharacter(character)) {
      throw new XlsxWriteError(
        'invalid-roundtrip-json',
        'XLSX cell content contains an invalid XML character',
        { cell: patch.cell.address, operationId: patch.operationId },
      );
    }
    output +=
      character === '&'
        ? '&amp;'
        : character === '<'
          ? '&lt;'
          : character === '>'
            ? '&gt;'
            : character === '\r'
              ? '&#13;'
              : character;
  }
  return output;
}

function serializedContent(
  patch: XlsxWorksheetCellPatch,
  prefix: string,
): {
  content: string;
  type?: string;
} {
  const qualified = (name: string): string => `${prefix}${name}`;
  const { content } = patch.cell;
  if (content.kind === 'blank') return { content: '' };
  if (content.kind === 'formula') {
    return {
      content: `<${qualified('f')}>${escapeXlsxCellText(content.formula.expression, patch)}</${qualified('f')}>`,
    };
  }
  const { value } = content;
  if (value.kind === 'text') {
    return {
      content: `<${qualified('is')}><${qualified('t')} xml:space="preserve">${escapeXlsxCellText(value.text, patch)}</${qualified('t')}></${qualified('is')}>`,
      type: 'inlineStr',
    };
  }
  if (value.kind === 'number') {
    return {
      content: `<${qualified('v')}>${String(value.value)}</${qualified('v')}>`,
    };
  }
  if (value.kind === 'boolean') {
    return {
      content: `<${qualified('v')}>${value.value ? '1' : '0'}</${qualified('v')}>`,
      type: 'b',
    };
  }
  if (value.kind === 'error') {
    return {
      content: `<${qualified('v')}>${escapeXlsxCellText(value.code, patch)}</${qualified('v')}>`,
      type: 'e',
    };
  }
  throw new XlsxWriteError(
    'unsupported-edit-operation',
    'XLSX worksheet patch does not support date values',
    {
      cell: patch.cell.address,
      featureClass: 'date-value',
      operationId: patch.operationId,
    },
  );
}

function cellReplacement(
  text: string,
  token: XlsxXmlTagToken,
  close: XlsxXmlTagToken,
  patch: XlsxWorksheetCellPatch,
  part: string,
): string {
  for (const attribute of token.attributes) {
    if (!ALLOWED_CELL_ATTRIBUTES.has(attribute.name)) {
      patchFailure(
        'XLSX target cell contains an unsupported attribute',
        part,
        patch,
        'cell-extension',
      );
    }
  }
  const contentChanged = patch.contentChanged !== false;
  const authored = token.attributes
    .filter(
      (attribute) =>
        (!contentChanged || attribute.name !== 't') &&
        (patch.xmlStyleIndex === undefined || attribute.name !== 's'),
    )
    .map((attribute) => text.slice(attribute.start, attribute.end))
    .join('');
  const style =
    patch.xmlStyleIndex === undefined ? '' : ` s="${patch.xmlStyleIndex}"`;
  if (!contentChanged) {
    return token.selfClosing
      ? `<${token.name}${authored}${style}/>`
      : `<${token.name}${authored}${style}>${text.slice(token.end, close.end)}`;
  }
  const prefix = token.name.slice(0, -xlsxXmlLocalName(token.name).length);
  const serialized = serializedContent(patch, prefix);
  const type = serialized.type === undefined ? '' : ` t="${serialized.type}"`;
  return serialized.content
    ? `<${token.name}${authored}${style}${type}>${serialized.content}</${token.name}>`
    : `<${token.name}${authored}${style}${type}/>`;
}

function matchingClose(
  tokens: readonly XlsxXmlTagToken[],
  index: number,
): number {
  const open = tokens[index]!;
  if (open.selfClosing) return index;
  const relative = tokens
    .slice(index + 1)
    .findIndex((token) => token.closing && token.name === open.name);
  return index + relative + 1;
}

function assertCellChildrenSafe(
  tokens: readonly XlsxXmlTagToken[],
  openIndex: number,
  closeIndex: number,
  patch: XlsxWorksheetCellPatch,
  part: string,
): void {
  if (patch.contentChanged === false) return;
  const directDepth = tokens[openIndex]!.depth + 1;
  const open = tokens[openIndex]!;
  const close = tokens[closeIndex]!;
  const directChildren = tokens.filter(
    (token) =>
      !token.closing &&
      token.depth === directDepth &&
      token.start >= open.end &&
      token.end <= close.start,
  );
  for (const token of directChildren) {
    const child = xlsxXmlLocalName(token.name);
    if (child !== 'f' && child !== 'is' && child !== 'v') {
      patchFailure(
        'XLSX target cell contains unsupported child content',
        part,
        patch,
        'cell-extension',
      );
    }
    if (child !== 'f' && token.attributes.length !== 0) {
      patchFailure(
        'XLSX target cell child contains unsupported attributes',
        part,
        patch,
        'cell-extension',
      );
    }
    if (token.attributes.some(isFormulaGroupAttribute)) {
      formulaFailure(
        'XLSX grouped formula cells cannot be patched independently',
        part,
        patch,
      );
    }
  }
}

function isFormulaGroupAttribute(attribute: XlsxXmlAttributeSpan): boolean {
  return (
    attribute.name === 'ref' ||
    attribute.name === 'si' ||
    attribute.name === 't'
  );
}

interface XlsxTextPatch {
  end: number;
  replacement: string;
  start: number;
}

export function patchXlsxWorksheetPartWithReport(
  bytes: Uint8Array,
  requested: readonly XlsxWorksheetCellPatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxWorksheetPatchResult {
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
  const patchesByAddress = new Map<string, XlsxWorksheetCellPatch>();
  for (const patch of requested) {
    if (patchesByAddress.has(patch.cell.address)) {
      patchFailure('XLSX worksheet patch cells must be unique', part, patch);
    }
    patchesByAddress.set(patch.cell.address, patch);
  }
  const decoded = decodeXlsxXml(bytes, part);
  const tokens = tokenizeXlsxXml(decoded.text, part);
  const root = tokens.find((token) => !token.closing);
  if (!root || xlsxXmlLocalName(root.name) !== 'worksheet') {
    patchFailure('XLSX worksheet root cannot be patched safely', part);
  }
  const qualifiedCell = `${root.name.slice(0, -'worksheet'.length)}c`;
  const found = new Set<string>();
  const textPatches: XlsxTextPatch[] = [];
  for (const [index, token] of tokens.entries()) {
    if (token.name !== qualifiedCell) continue;
    const reference = token.attributes.find(
      (attribute) => attribute.name === 'r',
    );
    if (!reference || reference.value.includes('&')) continue;
    const patch = patchesByAddress.get(reference.value);
    if (!patch) continue;
    if (found.has(reference.value)) {
      patchFailure('XLSX target cell reference is ambiguous', part, patch);
    }
    found.add(reference.value);
    const closeIndex = matchingClose(tokens, index);
    assertCellChildrenSafe(tokens, index, closeIndex, patch, part);
    textPatches.push({
      end: tokens[closeIndex]!.end,
      replacement: cellReplacement(
        decoded.text,
        token,
        tokens[closeIndex]!,
        patch,
        part,
      ),
      start: token.start,
    });
  }
  for (const patch of requested) {
    if (!found.has(patch.cell.address)) {
      patchFailure(
        'XLSX target cell has no safe explicit XML span',
        part,
        patch,
        'missing-cell-span',
      );
    }
  }
  let patchBytes = 0;
  for (const patch of textPatches) {
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
  textPatches.sort((left, right) => right.start - left.start);
  let output = decoded.text;
  for (const patch of textPatches) {
    output = `${output.slice(0, patch.start)}${patch.replacement}${output.slice(patch.end)}`;
  }
  const encoded = encodeXlsxXml({ ...decoded, text: output });
  if (encoded.byteLength > limits.maxGeneratedXmlBytes) {
    writeLimitFailure(
      'maxGeneratedXmlBytes',
      encoded.byteLength,
      limits.maxGeneratedXmlBytes,
      part,
    );
  }
  return {
    data: encoded,
    patchBytes,
    patchCount: textPatches.length,
  };
}

export function patchXlsxWorksheetPart(
  bytes: Uint8Array,
  requested: readonly XlsxWorksheetCellPatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
): Uint8Array {
  return patchXlsxWorksheetPartWithReport(bytes, requested, limits, part).data;
}
