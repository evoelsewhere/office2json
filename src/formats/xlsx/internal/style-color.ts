import { XlsxParseError } from '../errors';
import type { XlsxColor } from '../types';

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

function optionalTint(
  value: unknown,
  message: string,
  part: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (
    typeof value !== 'string' ||
    !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/u.test(value)
  ) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < -1 || parsed > 1) {
    valueFailure(message, part);
  }
  return parsed;
}

function rgb(
  value: unknown,
  message: string,
  part: string,
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !/^[0-9A-Fa-f]{8}$/u.test(value)) {
    valueFailure(message, part);
  }
  return value.toUpperCase();
}

export function parseXlsxStyleColor(
  value: unknown,
  part: string,
  context: string,
): XlsxColor | undefined {
  if (value === undefined) return undefined;
  const node = record(value);
  if (!node) structureFailure(`${context} color element is invalid`, part);
  const attrs = node.attrs === undefined ? {} : record(node.attrs);
  if (!attrs) structureFailure(`${context} color attributes are invalid`, part);

  const argb = rgb(attrs.rgb, `${context} color RGB is invalid`, part);
  const theme = unsignedInteger(
    attrs.theme,
    11,
    `${context} theme-color index is invalid`,
    part,
  );
  const indexed = unsignedInteger(
    attrs.indexed,
    65,
    `${context} indexed-color index is invalid`,
    part,
  );
  const automatic = optionalBoolean(
    attrs.auto,
    `${context} automatic-color flag is invalid`,
    part,
  );
  const tint = optionalTint(
    attrs.tint,
    `${context} color tint is invalid`,
    part,
  );

  const selectorCount =
    (argb === undefined ? 0 : 1) +
    (theme === undefined ? 0 : 1) +
    (indexed === undefined ? 0 : 1) +
    (automatic === true ? 1 : 0);
  if (selectorCount > 1) {
    structureFailure(`${context} color has multiple selectors`, part);
  }
  if (selectorCount === 0) {
    if (tint !== undefined) {
      valueFailure(`${context} color tint has no base color`, part);
    }
    return undefined;
  }
  if (automatic === true) {
    if (tint !== undefined) {
      valueFailure(`${context} automatic color cannot have a tint`, part);
    }
    return Object.freeze({ kind: 'automatic' });
  }

  const normalizedTint = tint === undefined || tint === 0 ? {} : { tint };
  if (argb !== undefined) {
    return Object.freeze({ argb, kind: 'rgb', ...normalizedTint });
  }
  if (theme !== undefined) {
    return Object.freeze({ index: theme, kind: 'theme', ...normalizedTint });
  }
  return Object.freeze({
    index: indexed!,
    kind: 'indexed',
    ...normalizedTint,
  });
}
