import { XlsxParseError } from '../errors';
import type { XlsxAlignment, XlsxProtection } from '../types';

type XmlRecord = Record<string, unknown>;

export interface XlsxXfFormatting {
  alignment?: XlsxAlignment;
  protection?: XlsxProtection;
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

function child(
  root: XmlRecord,
  prefix: string,
  localName: 'alignment' | 'protection',
  part: string,
): XmlRecord | undefined {
  const value = root[prefix ? `${prefix}:${localName}` : localName];
  if (value === undefined) return undefined;
  const values: unknown[] = Array.isArray(value) ? value : [value];
  if (values.length !== 1) {
    structureFailure(`XF ${localName} element is duplicated`, part);
  }
  const parsed = record(values[0]);
  if (!parsed) structureFailure(`XF ${localName} element is invalid`, part);
  return parsed;
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

function unsignedInteger(
  value: unknown,
  maximum: number,
  message: string,
  part: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (parsed > maximum) valueFailure(message, part);
  return parsed;
}

function signedInteger(
  value: unknown,
  minimum: number,
  maximum: number,
  message: string,
  part: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !/^[+-]?(?:0|[1-9]\d*)$/u.test(value)) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (parsed < minimum || parsed > maximum) valueFailure(message, part);
  return parsed;
}

function horizontal(value: unknown, part: string): XlsxAlignment['horizontal'] {
  if (value === undefined || value === 'general') return undefined;
  if (
    value === 'left' ||
    value === 'center' ||
    value === 'right' ||
    value === 'fill' ||
    value === 'justify' ||
    value === 'centerContinuous' ||
    value === 'distributed'
  ) {
    return value;
  }
  valueFailure('Alignment horizontal value is invalid', part);
}

function vertical(value: unknown, part: string): XlsxAlignment['vertical'] {
  if (value === undefined || value === 'bottom') return undefined;
  if (
    value === 'top' ||
    value === 'center' ||
    value === 'justify' ||
    value === 'distributed'
  ) {
    return value;
  }
  valueFailure('Alignment vertical value is invalid', part);
}

function textRotation(value: unknown, part: string): number | undefined {
  const parsed = unsignedInteger(
    value,
    255,
    'Alignment textRotation value is invalid',
    part,
  );
  if (parsed !== undefined && parsed > 180 && parsed !== 255) {
    valueFailure('Alignment textRotation value is invalid', part);
  }
  return parsed === 0 ? undefined : parsed;
}

function readingOrder(
  value: unknown,
  part: string,
): XlsxAlignment['readingOrder'] {
  const parsed = unsignedInteger(
    value,
    2,
    'Alignment readingOrder value is invalid',
    part,
  );
  if (parsed === undefined || parsed === 0) return undefined;
  return parsed === 1 ? 'left-to-right' : 'right-to-left';
}

function parseAlignment(
  node: XmlRecord,
  part: string,
): XlsxAlignment | undefined {
  const attrs = attributes(node, 'Alignment', part);
  const horizontalValue = horizontal(attrs.horizontal, part);
  const indent = unsignedInteger(
    attrs.indent,
    250,
    'Alignment indent value is invalid',
    part,
  );
  const justifyLastLine = optionalBoolean(
    attrs.justifyLastLine,
    'Alignment justifyLastLine value is invalid',
    part,
  );
  const order = readingOrder(attrs.readingOrder, part);
  const relativeIndent = signedInteger(
    attrs.relativeIndent,
    -15,
    15,
    'Alignment relativeIndent value is invalid',
    part,
  );
  const shrinkToFit = optionalBoolean(
    attrs.shrinkToFit,
    'Alignment shrinkToFit value is invalid',
    part,
  );
  const rotation = textRotation(attrs.textRotation, part);
  const verticalValue = vertical(attrs.vertical, part);
  const wrapText = optionalBoolean(
    attrs.wrapText,
    'Alignment wrapText value is invalid',
    part,
  );
  const normalized: XlsxAlignment = {
    ...(horizontalValue === undefined ? {} : { horizontal: horizontalValue }),
    ...(indent === undefined || indent === 0 ? {} : { indent }),
    ...(justifyLastLine ? { justifyLastLine: true } : {}),
    ...(order === undefined ? {} : { readingOrder: order }),
    ...(relativeIndent === undefined || relativeIndent === 0
      ? {}
      : { relativeIndent }),
    ...(shrinkToFit ? { shrinkToFit: true } : {}),
    ...(rotation === undefined ? {} : { textRotation: rotation }),
    ...(verticalValue === undefined ? {} : { vertical: verticalValue }),
    ...(wrapText ? { wrapText: true } : {}),
  };
  return Object.keys(normalized).length === 0
    ? undefined
    : Object.freeze(normalized);
}

function parseProtection(
  node: XmlRecord,
  part: string,
): XlsxProtection | undefined {
  const attrs = attributes(node, 'Protection', part);
  const hidden = optionalBoolean(
    attrs.hidden,
    'Protection hidden value is invalid',
    part,
  );
  const locked = optionalBoolean(
    attrs.locked,
    'Protection locked value is invalid',
    part,
  );
  const normalized: XlsxProtection = {
    ...(hidden ? { hidden: true } : {}),
    ...(locked === false ? { locked: false } : {}),
  };
  return Object.keys(normalized).length === 0
    ? undefined
    : Object.freeze(normalized);
}

export function parseXlsxXfFormatting(
  value: unknown,
  prefix: string,
  part: string,
): XlsxXfFormatting {
  const root = record(value);
  if (!root) structureFailure('XF element is invalid', part);
  const alignmentNode = child(root, prefix, 'alignment', part);
  const protectionNode = child(root, prefix, 'protection', part);
  const alignment = alignmentNode
    ? parseAlignment(alignmentNode, part)
    : undefined;
  const protection = protectionNode
    ? parseProtection(protectionNode, part)
    : undefined;
  return Object.freeze({
    ...(alignment === undefined ? {} : { alignment }),
    ...(protection === undefined ? {} : { protection }),
  });
}
