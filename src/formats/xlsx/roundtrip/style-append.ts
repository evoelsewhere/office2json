import { xlsxBuiltinNumberFormatCode } from '../internal/number-format';
import type {
  XlsxAlignment,
  XlsxBorder,
  XlsxBorderSide,
  XlsxColor,
  XlsxFill,
  XlsxFont,
  XlsxProtection,
  XlsxStyle,
} from '../types';
import { XlsxWriteError } from './errors';
import type { ResolvedXlsxWriteLimits } from './types';
import {
  decodeXlsxXml,
  encodeXlsxXml,
  tokenizeXlsxXml,
  xlsxXmlLocalName,
  type XlsxXmlAttributeSpan,
  type XlsxXmlTagToken,
} from './worksheet-patch';
import { writeLimitFailure } from './write-limits';

interface TextPatch {
  end: number;
  replacement: string;
  start: number;
}

interface CollectionAppend {
  items: string[];
  name: 'borders' | 'cellXfs' | 'fills' | 'fonts' | 'numFmts';
}

export interface XlsxStyleAppendResult {
  cellXfIndexes: number[];
  data: Uint8Array;
  patchBytes: number;
  patchCount: number;
}

function failure(message: string, part: string): never {
  throw new XlsxWriteError('preservation-conflict', message, {
    featureClass: 'styles-part',
    part,
  });
}

function invalidCharacter(value: string): boolean {
  for (const character of value) {
    const code = character.codePointAt(0)!;
    if (
      code !== 0x09 &&
      code !== 0x0a &&
      code !== 0x0d &&
      (code < 0x20 ||
        (code > 0xd7ff && code < 0xe000) ||
        (code > 0xfffd && code < 0x1_0000))
    ) {
      return true;
    }
  }
  return false;
}

function attribute(value: string, part: string): string {
  if (invalidCharacter(value)) {
    throw new XlsxWriteError(
      'invalid-roundtrip-json',
      'XLSX style text contains an invalid XML character',
      { part },
    );
  }
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('"', '&quot;')
    .replaceAll('\r', '&#13;')
    .replaceAll('\n', '&#10;')
    .replaceAll('\t', '&#9;');
}

function q(prefix: string, name: string): string {
  return `${prefix}${name}`;
}

function attrs(
  values: Readonly<Record<string, string | number | undefined>>,
): string {
  return Object.entries(values)
    .filter(
      (entry): entry is [string, string | number] => entry[1] !== undefined,
    )
    .map(([name, value]) => ` ${name}="${String(value)}"`)
    .join('');
}

function color(value: XlsxColor, prefix: string): string {
  const name = q(prefix, 'color');
  if (value.kind === 'automatic') return `<${name} auto="1"/>`;
  if (value.kind === 'rgb') {
    return `<${name}${attrs({ rgb: value.argb, tint: value.tint })}/>`;
  }
  return `<${name}${attrs({
    [value.kind === 'theme' ? 'theme' : 'indexed']: value.index,
    tint: value.tint,
  })}/>`;
}

function property(
  name: string,
  value: string | number | boolean | undefined,
  prefix: string,
  part: string,
): string {
  if (value === undefined) return '';
  const tag = q(prefix, name);
  return value === true
    ? `<${tag}/>`
    : `<${tag} val="${attribute(String(value), part)}"/>`;
}

function font(
  value: XlsxFont | undefined,
  prefix: string,
  part: string,
): string {
  const body = value
    ? [
        property('b', value.bold, prefix, part),
        property('i', value.italic, prefix, part),
        property('strike', value.strike, prefix, part),
        property('outline', value.outline, prefix, part),
        property('shadow', value.shadow, prefix, part),
        property('condense', value.condense, prefix, part),
        property('extend', value.extend, prefix, part),
        property('name', value.name, prefix, part),
        property('sz', value.size, prefix, part),
        property('family', value.family, prefix, part),
        property('charset', value.charset, prefix, part),
        property('scheme', value.scheme, prefix, part),
        property(
          'u',
          value.underline?.replace('-accounting', 'Accounting'),
          prefix,
          part,
        ),
        property('vertAlign', value.verticalAlignment, prefix, part),
        value.color === undefined ? '' : color(value.color, prefix),
      ].join('')
    : '';
  const name = q(prefix, 'font');
  return body ? `<${name}>${body}</${name}>` : `<${name}/>`;
}

function fill(value: XlsxFill | undefined, prefix: string): string {
  const fillName = q(prefix, 'fill');
  if (!value) {
    return `<${fillName}><${q(prefix, 'patternFill')} patternType="none"/></${fillName}>`;
  }
  if (value.kind === 'pattern') {
    const pattern = q(prefix, 'patternFill');
    const colors: string[] = [];
    if (value.foregroundColor !== undefined) {
      colors.push(
        color(value.foregroundColor, prefix).replace(
          `<${q(prefix, 'color')}`,
          `<${q(prefix, 'fgColor')}`,
        ),
      );
    }
    if (value.backgroundColor !== undefined) {
      colors.push(
        color(value.backgroundColor, prefix).replace(
          `<${q(prefix, 'color')}`,
          `<${q(prefix, 'bgColor')}`,
        ),
      );
    }
    return `<${fillName}><${pattern} patternType="${value.pattern}">${colors.join('')}</${pattern}></${fillName}>`;
  }
  const gradient = q(prefix, 'gradientFill');
  const stops = value.stops
    .map((stop) => {
      const stopName = q(prefix, 'stop');
      return `<${stopName} position="${stop.position}">${color(stop.color, prefix)}</${stopName}>`;
    })
    .join('');
  return `<${fillName}><${gradient}${attrs({
    bottom: value.bottom,
    degree: value.angle,
    left: value.left,
    right: value.right,
    top: value.top,
    type: value.type === 'path' ? 'path' : undefined,
  })}>${stops}</${gradient}></${fillName}>`;
}

function borderSide(
  name: string,
  value: XlsxBorderSide | undefined,
  prefix: string,
): string {
  const tag = q(prefix, name);
  if (!value) return `<${tag}/>`;
  const child = value.color === undefined ? '' : color(value.color, prefix);
  return child
    ? `<${tag}${attrs({ style: value.style })}>${child}</${tag}>`
    : `<${tag}${attrs({ style: value.style })}/>`;
}

function border(value: XlsxBorder | undefined, prefix: string): string {
  const name = q(prefix, 'border');
  if (!value) return `<${name}/>`;
  const body = (
    [
      'left',
      'right',
      'top',
      'bottom',
      'diagonal',
      'vertical',
      'horizontal',
      'start',
      'end',
    ] as const
  )
    .map((side) => borderSide(side, value[side], prefix))
    .join('');
  return `<${name}${attrs({
    diagonalDown: value.diagonalDown ? 1 : undefined,
    diagonalUp: value.diagonalUp ? 1 : undefined,
    outline: value.outline === false ? 0 : undefined,
  })}>${body}</${name}>`;
}

function alignment(value: XlsxAlignment | undefined, prefix: string): string {
  const name = q(prefix, 'alignment');
  return `<${name}${
    value
      ? attrs({
          horizontal: value.horizontal,
          indent: value.indent,
          justifyLastLine: value.justifyLastLine ? 1 : undefined,
          readingOrder:
            value.readingOrder === 'left-to-right'
              ? 1
              : value.readingOrder === 'right-to-left'
                ? 2
                : undefined,
          relativeIndent: value.relativeIndent,
          shrinkToFit: value.shrinkToFit ? 1 : undefined,
          textRotation: value.textRotation,
          vertical: value.vertical,
          wrapText: value.wrapText ? 1 : undefined,
        })
      : ''
  }/>`;
}

function protection(value: XlsxProtection | undefined, prefix: string): string {
  const name = q(prefix, 'protection');
  return `<${name}${
    value
      ? attrs({
          hidden: value.hidden ? 1 : undefined,
          locked: value.locked === false ? 0 : undefined,
        })
      : ''
  }/>`;
}

function builtinNumberFormat(code: string): number | undefined {
  const ids = [
    0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 37,
    38, 39, 40, 45, 46, 47, 48, 49,
  ];
  return ids.find((id) => xlsxBuiltinNumberFormatCode(id) === code);
}

export function xlsxAppendedStyleRecordCount(
  styles: readonly XlsxStyle[],
): number {
  return styles.reduce(
    (total, style) =>
      total +
      4 +
      (style.numberFormat !== undefined &&
      builtinNumberFormat(style.numberFormat) === undefined
        ? 1
        : 0),
    0,
  );
}

function countAttribute(
  token: XlsxXmlTagToken,
  part: string,
): XlsxXmlAttributeSpan {
  const count = token.attributes.find(
    (candidate) => candidate.name === 'count',
  );
  if (!count)
    failure(
      `XLSX styles ${xlsxXmlLocalName(token.name)} count cannot be patched`,
      part,
    );
  return count;
}

function childCollection(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  name: CollectionAppend['name'],
): XlsxXmlTagToken | undefined {
  const prefix = root.name.slice(0, -xlsxXmlLocalName(root.name).length);
  return tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}${name}`,
  );
}

function requiredCollection(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  name: Exclude<CollectionAppend['name'], 'numFmts'>,
  part: string,
): XlsxXmlTagToken {
  const token = childCollection(tokens, root, name);
  if (!token) failure(`XLSX styles ${name} collection is missing`, part);
  return token;
}

function closingToken(
  tokens: readonly XlsxXmlTagToken[],
  open: XlsxXmlTagToken,
): XlsxXmlTagToken {
  const close = tokens.find(
    (token) =>
      token.closing && token.depth === open.depth && token.name === open.name,
  );
  return close!;
}

function replaceCount(
  text: string,
  token: XlsxXmlTagToken,
  count: number,
  part: string,
): TextPatch {
  const source = countAttribute(token, part);
  return {
    end: source.end,
    replacement: ` count="${count}"`,
    start: source.start,
  };
}

function appendCollection(
  text: string,
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  request: CollectionAppend,
  prefix: string,
  part: string,
): TextPatch[] {
  if (request.items.length === 0) return [];
  const token = childCollection(tokens, root, request.name);
  if (!token) {
    const fonts = requiredCollection(tokens, root, 'fonts', part);
    const name = q(prefix, request.name);
    return [
      {
        end: fonts.start,
        replacement: `<${name} count="${request.items.length}">${request.items.join('')}</${name}>`,
        start: fonts.start,
      },
    ];
  }
  const count = Number(countAttribute(token, part).value);
  const nextCount = count + request.items.length;
  if (token.selfClosing) {
    const raw = text.slice(token.start, token.end);
    const countSpan = countAttribute(token, part);
    const relativeStart = countSpan.start - token.start;
    const relativeEnd = countSpan.end - token.start;
    const opening =
      `${raw.slice(0, relativeStart)} count="${nextCount}"${raw.slice(relativeEnd)}`.replace(
        /\/\s*>$/u,
        '>',
      );
    return [
      {
        end: token.end,
        replacement: `${opening}${request.items.join('')}</${token.name}>`,
        start: token.start,
      },
    ];
  }
  const close = closingToken(tokens, token);
  return [
    replaceCount(text, token, nextCount, part),
    {
      end: close.start,
      replacement: request.items.join(''),
      start: close.start,
    },
  ];
}

function maximumCustomNumberFormatId(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
): number {
  const collection = childCollection(tokens, root, 'numFmts');
  if (!collection || collection.selfClosing) return 163;
  const close = closingToken(tokens, collection);
  let maximum = 163;
  for (const token of tokens) {
    if (
      token.closing ||
      token.depth !== collection.depth + 1 ||
      token.start < collection.end ||
      token.end > close.start ||
      xlsxXmlLocalName(token.name) !== 'numFmt'
    ) {
      continue;
    }
    const value = token.attributes.find(
      (candidate) => candidate.name === 'numFmtId',
    )?.value;
    if (value !== undefined) maximum = Math.max(maximum, Number(value));
  }
  return maximum;
}

export function appendXlsxStylesPart(
  bytes: Uint8Array,
  appended: readonly XlsxStyle[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxStyleAppendResult {
  const decoded = decodeXlsxXml(bytes, part);
  const tokens = tokenizeXlsxXml(decoded.text, part);
  const root = tokens.find(
    (token) =>
      token.depth === 0 && xlsxXmlLocalName(token.name) === 'styleSheet',
  );
  if (!root) failure('XLSX styles root cannot be patched', part);
  const prefix = root.name.slice(0, -'styleSheet'.length);
  const cellXfs = requiredCollection(tokens, root, 'cellXfs', part);
  const sourceCellXfCount = Number(countAttribute(cellXfs, part).value);
  const cellXfIndexes = appended.map(
    (_style, index) => sourceCellXfCount + index,
  );
  let customId = maximumCustomNumberFormatId(tokens, root);
  const numberFormats: string[] = [];
  const fonts: string[] = [];
  const fills: string[] = [];
  const borders: string[] = [];
  const xfs: string[] = [];
  const fontCount = Number(
    countAttribute(requiredCollection(tokens, root, 'fonts', part), part).value,
  );
  const fillCount = Number(
    countAttribute(requiredCollection(tokens, root, 'fills', part), part).value,
  );
  const borderCount = Number(
    countAttribute(requiredCollection(tokens, root, 'borders', part), part)
      .value,
  );
  for (const [index, style] of appended.entries()) {
    fonts.push(font(style.font, prefix, part));
    fills.push(fill(style.fill, prefix));
    borders.push(border(style.border, prefix));
    let numberFormatId = 0;
    if (style.numberFormat !== undefined) {
      const builtin = builtinNumberFormat(style.numberFormat);
      if (builtin !== undefined) {
        numberFormatId = builtin;
      } else {
        if (customId >= 0xffff_ffff) {
          throw new XlsxWriteError(
            'identifier-allocation-failed',
            'XLSX custom number-format IDs are exhausted',
            { featureClass: 'number-format', part },
          );
        }
        customId += 1;
        numberFormatId = customId;
        const name = q(prefix, 'numFmt');
        numberFormats.push(
          `<${name} numFmtId="${numberFormatId}" formatCode="${attribute(style.numberFormat, part)}"/>`,
        );
      }
    }
    const xf = q(prefix, 'xf');
    xfs.push(
      `<${xf}${attrs({
        applyAlignment: 1,
        applyBorder: 1,
        applyFill: 1,
        applyFont: 1,
        applyNumberFormat: 1,
        applyProtection: 1,
        borderId: borderCount + index,
        fillId: fillCount + index,
        fontId: fontCount + index,
        numFmtId: numberFormatId,
        xfId: 0,
      })}>${alignment(style.alignment, prefix)}${protection(style.protection, prefix)}</${xf}>`,
    );
  }
  const requests: CollectionAppend[] = [
    { items: numberFormats, name: 'numFmts' },
    { items: fonts, name: 'fonts' },
    { items: fills, name: 'fills' },
    { items: borders, name: 'borders' },
    { items: xfs, name: 'cellXfs' },
  ];
  const patches = requests.flatMap((request) =>
    appendCollection(decoded.text, tokens, root, request, prefix, part),
  );
  patches.sort((left, right) => right.start - left.start);
  let output = decoded.text;
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
  return { cellXfIndexes, data, patchBytes, patchCount: patches.length };
}
