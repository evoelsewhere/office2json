import { XlsxParseError } from '../errors';
import type { XlsxStyle } from '../types';
import { parseXlsxStyleBorder } from './style-border';
import { parseXlsxStyleFill } from './style-fill';
import { parseXlsxStyleFont } from './style-font';
import { parseXlsxXfFormatting } from './style-formatting';

type XmlRecord = Record<string, unknown>;

function structureFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-structure',
    message,
    part,
    severity: 'error',
  });
}

function valueFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function child(
  root: XmlRecord,
  prefix: string,
  localName: 'border' | 'fill' | 'font' | 'numFmt',
  part: string,
): XmlRecord | undefined {
  const value = root[prefix ? `${prefix}:${localName}` : localName];
  if (value === undefined) return undefined;
  const values: unknown[] = Array.isArray(value) ? value : [value];
  if (values.length !== 1) {
    structureFailure(
      `Differential style ${localName} element is duplicated`,
      part,
    );
  }
  const parsed = record(values[0]);
  if (!parsed) {
    structureFailure(
      `Differential style ${localName} element is invalid`,
      part,
    );
  }
  return parsed;
}

function attributes(node: XmlRecord, context: string, part: string): XmlRecord {
  if (node.attrs === undefined) return {};
  const attrs = record(node.attrs);
  if (!attrs) structureFailure(`${context} attributes are invalid`, part);
  return attrs;
}

function unsignedInteger(
  value: unknown,
  message: string,
  part: string,
): number {
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (parsed > 0xffff_ffff) valueFailure(message, part);
  return parsed;
}

function numberFormat(
  node: XmlRecord | undefined,
  part: string,
): string | undefined {
  if (!node) return undefined;
  const attrs = attributes(node, 'Differential number format', part);
  unsignedInteger(
    attrs.numFmtId,
    'Differential number-format ID is invalid',
    part,
  );
  if (typeof attrs.formatCode !== 'string' || attrs.formatCode.length === 0) {
    valueFailure('Differential number-format code is invalid', part);
  }
  return attrs.formatCode === 'General' ? undefined : attrs.formatCode;
}

function semanticFill(fill: XlsxStyle['fill']): XlsxStyle['fill'] {
  if (
    fill?.kind === 'pattern' &&
    fill.pattern === 'none' &&
    fill.foregroundColor === undefined &&
    fill.backgroundColor === undefined
  ) {
    return undefined;
  }
  return fill;
}

export function parseXlsxDifferentialStyle(
  value: unknown,
  prefix: string,
  part: string,
): XlsxStyle {
  const root = record(value);
  if (!root) structureFailure('Differential style element is invalid', part);
  const borderNode = child(root, prefix, 'border', part);
  const fillNode = child(root, prefix, 'fill', part);
  const fontNode = child(root, prefix, 'font', part);
  const numFmtNode = child(root, prefix, 'numFmt', part);
  const formatting = parseXlsxXfFormatting(root, prefix, part);
  const border = borderNode
    ? parseXlsxStyleBorder(borderNode, prefix, part)
    : undefined;
  const fill = semanticFill(
    fillNode ? parseXlsxStyleFill(fillNode, prefix, part) : undefined,
  );
  const font = fontNode
    ? parseXlsxStyleFont(fontNode, prefix, part)
    : undefined;
  const formatCode = numberFormat(numFmtNode, part);
  return Object.freeze({
    ...formatting,
    ...(border === undefined || Object.keys(border).length === 0
      ? {}
      : { border }),
    ...(fill === undefined ? {} : { fill }),
    ...(font === undefined || Object.keys(font).length === 0 ? {} : { font }),
    ...(formatCode === undefined ? {} : { numberFormat: formatCode }),
  });
}
