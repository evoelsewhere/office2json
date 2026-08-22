import { decodeXmlEntities } from '../../../common/text/html';
import { XlsxParseError } from '../errors';
import type { XlsxFont } from '../types';
import { parseXlsxStyleColor } from './style-color';

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
  localName: string,
  part: string,
): XmlRecord | undefined {
  const value = root[prefix ? `${prefix}:${localName}` : localName];
  if (value === undefined) return undefined;
  if (Array.isArray(value)) {
    if (value.length !== 1) {
      structureFailure(`Font ${localName} element is duplicated`, part);
    }
    const item = record(value[0]);
    if (!item) structureFailure(`Font ${localName} element is invalid`, part);
    return item;
  }
  const item = record(value);
  if (!item) structureFailure(`Font ${localName} element is invalid`, part);
  return item;
}

function attributes(
  node: XmlRecord,
  localName: string,
  part: string,
): XmlRecord {
  if (node.attrs === undefined) return {};
  const attrs = record(node.attrs);
  if (!attrs)
    structureFailure(`Font ${localName} attributes are invalid`, part);
  return attrs;
}

function booleanProperty(
  root: XmlRecord,
  prefix: string,
  localName: string,
  part: string,
): boolean {
  const node = child(root, prefix, localName, part);
  if (!node) return false;
  const value = attributes(node, localName, part).val;
  if (value === undefined || value === '1' || value === 'true') return true;
  if (value === '0' || value === 'false') return false;
  valueFailure(`Font ${localName} value is invalid`, part);
}

function stringProperty(
  root: XmlRecord,
  prefix: string,
  localName: string,
  part: string,
): string | undefined {
  const node = child(root, prefix, localName, part);
  if (!node) return undefined;
  const value = attributes(node, localName, part).val;
  if (typeof value !== 'string' || value.length === 0) {
    valueFailure(`Font ${localName} value is invalid`, part);
  }
  return decodeXmlEntities(value);
}

function unsignedProperty(
  root: XmlRecord,
  prefix: string,
  localName: string,
  maximum: number,
  part: string,
): number | undefined {
  const value = stringProperty(root, prefix, localName, part);
  if (value === undefined) return undefined;
  if (!/^(?:0|[1-9]\d*)$/u.test(value)) {
    valueFailure(`Font ${localName} value is invalid`, part);
  }
  const parsed = Number(value);
  if (parsed > maximum)
    valueFailure(`Font ${localName} value is invalid`, part);
  return parsed;
}

function sizeProperty(
  root: XmlRecord,
  prefix: string,
  part: string,
): number | undefined {
  const value = stringProperty(root, prefix, 'sz', part);
  if (value === undefined) return undefined;
  if (!/^[+]?(?:\d+(?:\.\d*)?|\.\d+)$/u.test(value)) {
    valueFailure('Font sz value is invalid', part);
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0 || parsed > 409) {
    valueFailure('Font sz value is invalid', part);
  }
  return parsed;
}

function enumProperty<T extends string>(
  root: XmlRecord,
  prefix: string,
  localName: string,
  values: readonly T[],
  part: string,
): T | undefined {
  const value = stringProperty(root, prefix, localName, part);
  if (value === undefined) return undefined;
  if (!values.includes(value as T)) {
    valueFailure(`Font ${localName} value is invalid`, part);
  }
  return value as T;
}

function underlineProperty(
  root: XmlRecord,
  prefix: string,
  part: string,
):
  | 'double'
  | 'doubleAccounting'
  | 'none'
  | 'single'
  | 'singleAccounting'
  | undefined {
  const node = child(root, prefix, 'u', part);
  if (!node) return undefined;
  const value = attributes(node, 'u', part).val;
  if (value === undefined) return 'single';
  if (
    value === 'single' ||
    value === 'double' ||
    value === 'singleAccounting' ||
    value === 'doubleAccounting' ||
    value === 'none'
  ) {
    return value;
  }
  valueFailure('Font u value is invalid', part);
}

export function parseXlsxStyleFont(
  value: unknown,
  prefix: string,
  part: string,
): XlsxFont {
  const root = record(value);
  if (!root) structureFailure('Font element is invalid', part);
  const bold = booleanProperty(root, prefix, 'b', part);
  const italic = booleanProperty(root, prefix, 'i', part);
  const strike = booleanProperty(root, prefix, 'strike', part);
  const outline = booleanProperty(root, prefix, 'outline', part);
  const shadow = booleanProperty(root, prefix, 'shadow', part);
  const condense = booleanProperty(root, prefix, 'condense', part);
  const extend = booleanProperty(root, prefix, 'extend', part);
  const name = stringProperty(root, prefix, 'name', part);
  const size = sizeProperty(root, prefix, part);
  const family = unsignedProperty(root, prefix, 'family', 5, part);
  const charset = unsignedProperty(root, prefix, 'charset', 255, part);
  const scheme = enumProperty(
    root,
    prefix,
    'scheme',
    ['major', 'minor', 'none'] as const,
    part,
  );
  const underline = underlineProperty(root, prefix, part);
  const verticalAlignment = enumProperty(
    root,
    prefix,
    'vertAlign',
    ['baseline', 'superscript', 'subscript'] as const,
    part,
  );
  const colorNode = child(root, prefix, 'color', part);
  const color = parseXlsxStyleColor(colorNode, part, 'Font');

  return Object.freeze({
    ...(bold ? { bold: true } : {}),
    ...(charset === undefined ? {} : { charset }),
    ...(color === undefined ? {} : { color }),
    ...(condense ? { condense: true } : {}),
    ...(extend ? { extend: true } : {}),
    ...(family === undefined ? {} : { family }),
    ...(italic ? { italic: true } : {}),
    ...(name === undefined ? {} : { name }),
    ...(outline ? { outline: true } : {}),
    ...(scheme === undefined || scheme === 'none' ? {} : { scheme }),
    ...(shadow ? { shadow: true } : {}),
    ...(size === undefined ? {} : { size }),
    ...(strike ? { strike: true } : {}),
    ...(underline === undefined || underline === 'none'
      ? {}
      : {
          underline:
            underline === 'singleAccounting'
              ? 'single-accounting'
              : underline === 'doubleAccounting'
                ? 'double-accounting'
                : underline,
        }),
    ...(verticalAlignment === undefined || verticalAlignment === 'baseline'
      ? {}
      : { verticalAlignment }),
  });
}
