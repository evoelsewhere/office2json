import { XlsxParseError } from '../errors';
import type { XlsxBorder, XlsxBorderSide, XlsxBorderStyle } from '../types';
import { parseXlsxStyleColor } from './style-color';

type XmlRecord = Record<string, unknown>;
type BorderSideName =
  | 'bottom'
  | 'diagonal'
  | 'end'
  | 'horizontal'
  | 'left'
  | 'right'
  | 'start'
  | 'top'
  | 'vertical';

function borderStyles(): readonly (XlsxBorderStyle | 'none')[] {
  return [
    'none',
    'thin',
    'medium',
    'dashed',
    'dotted',
    'thick',
    'double',
    'hair',
    'mediumDashed',
    'dashDot',
    'mediumDashDot',
    'dashDotDot',
    'mediumDashDotDot',
    'slantDashDot',
  ];
}

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

function attributes(node: XmlRecord, context: string, part: string): XmlRecord {
  if (node.attrs === undefined) return {};
  const attrs = record(node.attrs);
  if (!attrs) structureFailure(`${context} attributes are invalid`, part);
  return attrs;
}

function optionalBoolean(
  value: unknown,
  message: string,
  part: string,
): boolean | undefined {
  if (value === undefined) return undefined;
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  valueFailure(message, part);
}

function child(
  root: XmlRecord,
  prefix: string,
  localName: string,
  context: string,
  part: string,
): XmlRecord | undefined {
  const value = root[prefix ? `${prefix}:${localName}` : localName];
  if (value === undefined) return undefined;
  const values: unknown[] = Array.isArray(value) ? value : [value];
  if (values.length !== 1) {
    structureFailure(`${context} ${localName} element is duplicated`, part);
  }
  const parsed = record(values[0]);
  if (!parsed) {
    structureFailure(`${context} ${localName} element is invalid`, part);
  }
  return parsed;
}

function side(
  root: XmlRecord,
  prefix: string,
  name: BorderSideName,
  part: string,
): XlsxBorderSide | undefined {
  const node = child(root, prefix, name, 'Border', part);
  if (!node) return undefined;
  const sourceStyle = attributes(node, `Border ${name}`, part).style;
  if (
    sourceStyle !== undefined &&
    !borderStyles().includes(sourceStyle as XlsxBorderStyle | 'none')
  ) {
    valueFailure(`Border ${name} style is invalid`, part);
  }
  const style = sourceStyle === 'none' ? undefined : sourceStyle;
  const color = parseXlsxStyleColor(
    child(node, prefix, 'color', `Border ${name}`, part),
    part,
    `Border ${name}`,
  );
  if (style === undefined && color === undefined) return undefined;
  return Object.freeze({
    ...(color === undefined ? {} : { color }),
    ...(style === undefined ? {} : { style: style as XlsxBorderStyle }),
  });
}

export function parseXlsxStyleBorder(
  value: unknown,
  prefix: string,
  part: string,
): XlsxBorder {
  const root = record(value);
  if (!root) structureFailure('Border element is invalid', part);
  const attrs = attributes(root, 'Border', part);
  const diagonalDown = optionalBoolean(
    attrs.diagonalDown,
    'Border diagonalDown value is invalid',
    part,
  );
  const diagonalUp = optionalBoolean(
    attrs.diagonalUp,
    'Border diagonalUp value is invalid',
    part,
  );
  const outline = optionalBoolean(
    attrs.outline,
    'Border outline value is invalid',
    part,
  );
  const bottom = side(root, prefix, 'bottom', part);
  const diagonal = side(root, prefix, 'diagonal', part);
  const end = side(root, prefix, 'end', part);
  const horizontal = side(root, prefix, 'horizontal', part);
  const left = side(root, prefix, 'left', part);
  const right = side(root, prefix, 'right', part);
  const start = side(root, prefix, 'start', part);
  const top = side(root, prefix, 'top', part);
  const vertical = side(root, prefix, 'vertical', part);

  return Object.freeze({
    ...(bottom === undefined ? {} : { bottom }),
    ...(diagonal === undefined ? {} : { diagonal }),
    ...(diagonalDown ? { diagonalDown: true } : {}),
    ...(diagonalUp ? { diagonalUp: true } : {}),
    ...(end === undefined ? {} : { end }),
    ...(horizontal === undefined ? {} : { horizontal }),
    ...(left === undefined ? {} : { left }),
    ...(outline === false ? { outline: false } : {}),
    ...(right === undefined ? {} : { right }),
    ...(start === undefined ? {} : { start }),
    ...(top === undefined ? {} : { top }),
    ...(vertical === undefined ? {} : { vertical }),
  });
}
