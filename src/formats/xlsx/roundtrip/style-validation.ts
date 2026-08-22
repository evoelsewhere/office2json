import type {
  XlsxAlignment,
  XlsxBorder,
  XlsxBorderSide,
  XlsxColor,
  XlsxFill,
  XlsxFont,
  XlsxGradientStop,
  XlsxProtection,
  XlsxStyle,
} from '../types';
import { XlsxWriteError } from './errors';

type Plain = Record<string, unknown>;

const STYLE_KEYS = [
  'alignment',
  'border',
  'checkbox',
  'fill',
  'font',
  'numberFormat',
  'protection',
] as const;
const ALIGNMENT_KEYS = [
  'horizontal',
  'indent',
  'justifyLastLine',
  'readingOrder',
  'relativeIndent',
  'shrinkToFit',
  'textRotation',
  'vertical',
  'wrapText',
] as const;
const BORDER_KEYS = [
  'bottom',
  'diagonal',
  'diagonalDown',
  'diagonalUp',
  'end',
  'horizontal',
  'left',
  'outline',
  'right',
  'start',
  'top',
  'vertical',
] as const;
const BORDER_SIDE_KEYS = ['color', 'style'] as const;
const FONT_KEYS = [
  'bold',
  'charset',
  'color',
  'condense',
  'extend',
  'family',
  'italic',
  'name',
  'outline',
  'scheme',
  'shadow',
  'size',
  'strike',
  'underline',
  'verticalAlignment',
] as const;
const PROTECTION_KEYS = ['hidden', 'locked'] as const;
const COLOR_KINDS = ['automatic', 'indexed', 'rgb', 'theme'] as const;
const BORDER_STYLES = [
  'dashDot',
  'dashDotDot',
  'dashed',
  'dotted',
  'double',
  'hair',
  'medium',
  'mediumDashDot',
  'mediumDashDotDot',
  'mediumDashed',
  'slantDashDot',
  'thick',
  'thin',
] as const;
const PATTERN_TYPES = [
  'darkDown',
  'darkGray',
  'darkGrid',
  'darkHorizontal',
  'darkTrellis',
  'darkUp',
  'darkVertical',
  'gray0625',
  'gray125',
  'lightDown',
  'lightGray',
  'lightGrid',
  'lightHorizontal',
  'lightTrellis',
  'lightUp',
  'lightVertical',
  'mediumGray',
  'none',
  'solid',
] as const;

function invalid(message: string, operationId: string): never {
  throw new XlsxWriteError('invalid-roundtrip-json', message, { operationId });
}

function plain(value: unknown, message: string, id: string): Plain {
  if (
    value === null ||
    value === undefined ||
    Array.isArray(value) ||
    Object.getPrototypeOf(value) !== Object.prototype
  ) {
    invalid(message, id);
  }
  return value as Plain;
}

function exact(
  value: Plain,
  allowed: readonly string[],
  message: string,
  id: string,
): void {
  if (Object.keys(value).some((key) => !allowed.includes(key))) {
    invalid(message, id);
  }
}

function enumValue<T extends string>(
  value: unknown,
  allowed: readonly T[],
  message: string,
  id: string,
): T {
  if (!allowed.includes(value as T)) {
    invalid(message, id);
  }
  return value as T;
}

function normalizedTrue(value: unknown, message: string, id: string): true {
  if (value !== true) invalid(message, id);
  return true;
}

function integer(
  value: unknown,
  minimum: number,
  maximum: number,
  message: string,
  id: string,
): number {
  if (
    typeof value !== 'number' ||
    !Number.isInteger(value) ||
    value < minimum ||
    value > maximum
  ) {
    invalid(message, id);
  }
  return value;
}

function finite(
  value: unknown,
  minimum: number,
  maximum: number,
  message: string,
  id: string,
): number {
  if (
    typeof value !== 'number' ||
    !Number.isFinite(value) ||
    value < minimum ||
    value > maximum
  ) {
    invalid(message, id);
  }
  return Object.is(value, -0) ? 0 : value;
}

function tint(value: unknown, message: string, id: string): number | undefined {
  if (value === undefined) return undefined;
  const result = finite(value, -1, 1, message, id);
  if (result === 0) invalid(message, id);
  return result;
}

function color(value: unknown, id: string): XlsxColor {
  const record = plain(value, 'XLSX style color shape is invalid', id);
  const kind = enumValue(
    record.kind,
    COLOR_KINDS,
    'XLSX style color kind is invalid',
    id,
  );
  if (kind === 'automatic') {
    if (Object.keys(record).length !== 1) {
      invalid('XLSX automatic style color shape is invalid', id);
    }
    return { kind };
  }
  if (kind === 'rgb') {
    exact(
      record,
      ['argb', 'kind', 'tint'],
      'XLSX RGB style color shape is invalid',
      id,
    );
    if (
      typeof record.argb !== 'string' ||
      !/^[0-9A-F]{8}$/u.test(record.argb)
    ) {
      invalid('XLSX RGB style color value is invalid', id);
    }
    const normalizedTint = tint(
      record.tint,
      'XLSX RGB style color tint is invalid',
      id,
    );
    return {
      argb: record.argb,
      kind,
      ...(normalizedTint === undefined ? {} : { tint: normalizedTint }),
    };
  }
  exact(
    record,
    ['index', 'kind', 'tint'],
    'XLSX indexed style color shape is invalid',
    id,
  );
  const index = integer(
    record.index,
    0,
    kind === 'theme' ? 11 : 65,
    'XLSX style color index is invalid',
    id,
  );
  const normalizedTint = tint(
    record.tint,
    'XLSX style color tint is invalid',
    id,
  );
  return {
    index,
    kind,
    ...(normalizedTint === undefined ? {} : { tint: normalizedTint }),
  };
}

function alignment(value: unknown, id: string): XlsxAlignment {
  const record = plain(value, 'XLSX style alignment shape is invalid', id);
  exact(record, ALIGNMENT_KEYS, 'XLSX style alignment shape is invalid', id);
  const result: XlsxAlignment = {};
  if (record.horizontal !== undefined) {
    result.horizontal = enumValue(
      record.horizontal,
      [
        'center',
        'centerContinuous',
        'distributed',
        'fill',
        'justify',
        'left',
        'right',
      ] as const,
      'XLSX style horizontal alignment is invalid',
      id,
    );
  }
  if (record.indent !== undefined) {
    result.indent = integer(
      record.indent,
      1,
      250,
      'XLSX style alignment indent is invalid',
      id,
    );
  }
  if (record.justifyLastLine !== undefined) {
    result.justifyLastLine = normalizedTrue(
      record.justifyLastLine,
      'XLSX style justify-last-line flag is invalid',
      id,
    );
  }
  if (record.readingOrder !== undefined) {
    result.readingOrder = enumValue(
      record.readingOrder,
      ['left-to-right', 'right-to-left'] as const,
      'XLSX style reading order is invalid',
      id,
    );
  }
  if (record.relativeIndent !== undefined) {
    const relative = integer(
      record.relativeIndent,
      -15,
      15,
      'XLSX style relative indent is invalid',
      id,
    );
    if (relative === 0) invalid('XLSX style relative indent is invalid', id);
    result.relativeIndent = relative;
  }
  if (record.shrinkToFit !== undefined) {
    result.shrinkToFit = normalizedTrue(
      record.shrinkToFit,
      'XLSX style shrink-to-fit flag is invalid',
      id,
    );
  }
  if (record.textRotation !== undefined) {
    const rotation = integer(
      record.textRotation,
      1,
      255,
      'XLSX style text rotation is invalid',
      id,
    );
    if (rotation > 180 && rotation !== 255)
      invalid('XLSX style text rotation is invalid', id);
    result.textRotation = rotation;
  }
  if (record.vertical !== undefined) {
    result.vertical = enumValue(
      record.vertical,
      ['center', 'distributed', 'justify', 'top'] as const,
      'XLSX style vertical alignment is invalid',
      id,
    );
  }
  if (record.wrapText !== undefined) {
    result.wrapText = normalizedTrue(
      record.wrapText,
      'XLSX style wrap-text flag is invalid',
      id,
    );
  }
  if (Object.keys(result).length === 0)
    invalid('XLSX style alignment must be normalized', id);
  return result;
}

function borderSide(value: unknown, id: string): XlsxBorderSide {
  const record = plain(value, 'XLSX style border-side shape is invalid', id);
  exact(
    record,
    BORDER_SIDE_KEYS,
    'XLSX style border-side shape is invalid',
    id,
  );
  const result: XlsxBorderSide = {};
  if (record.color !== undefined) result.color = color(record.color, id);
  if (record.style !== undefined) {
    result.style = enumValue(
      record.style,
      BORDER_STYLES,
      'XLSX style border-side kind is invalid',
      id,
    );
  }
  if (Object.keys(result).length === 0)
    invalid('XLSX style border side must be normalized', id);
  return result;
}

function border(value: unknown, id: string): XlsxBorder {
  const record = plain(value, 'XLSX style border shape is invalid', id);
  exact(record, BORDER_KEYS, 'XLSX style border shape is invalid', id);
  const result: XlsxBorder = {};
  for (const name of [
    'bottom',
    'diagonal',
    'end',
    'horizontal',
    'left',
    'right',
    'start',
    'top',
    'vertical',
  ] as const) {
    if (record[name] !== undefined) result[name] = borderSide(record[name], id);
  }
  for (const name of ['diagonalDown', 'diagonalUp'] as const) {
    if (record[name] !== undefined)
      result[name] = normalizedTrue(
        record[name],
        `XLSX style border ${name} flag is invalid`,
        id,
      );
  }
  if (record.outline !== undefined) {
    if (record.outline !== false)
      invalid('XLSX style border outline flag is invalid', id);
    result.outline = false;
  }
  if (Object.keys(result).length === 0)
    invalid('XLSX style border must be normalized', id);
  return result;
}

function gradientStop(value: unknown, id: string): XlsxGradientStop {
  const record = plain(value, 'XLSX gradient stop shape is invalid', id);
  exact(
    record,
    ['color', 'position'],
    'XLSX gradient stop shape is invalid',
    id,
  );
  if (!Object.hasOwn(record, 'color') || !Object.hasOwn(record, 'position')) {
    invalid('XLSX gradient stop shape is invalid', id);
  }
  return {
    color: color(record.color, id),
    position: finite(
      record.position,
      0,
      1,
      'XLSX gradient stop position is invalid',
      id,
    ),
  };
}

function fill(value: unknown, id: string): XlsxFill {
  const record = plain(value, 'XLSX style fill shape is invalid', id);
  if (record.kind === 'pattern') {
    exact(
      record,
      ['backgroundColor', 'foregroundColor', 'kind', 'pattern'],
      'XLSX pattern fill shape is invalid',
      id,
    );
    return {
      ...(record.backgroundColor === undefined
        ? {}
        : { backgroundColor: color(record.backgroundColor, id) }),
      ...(record.foregroundColor === undefined
        ? {}
        : { foregroundColor: color(record.foregroundColor, id) }),
      kind: 'pattern',
      pattern: enumValue(
        record.pattern,
        PATTERN_TYPES,
        'XLSX pattern fill type is invalid',
        id,
      ),
    };
  }
  if (record.kind !== 'gradient')
    invalid('XLSX style fill kind is invalid', id);
  exact(
    record,
    ['angle', 'bottom', 'kind', 'left', 'right', 'stops', 'top', 'type'],
    'XLSX gradient fill shape is invalid',
    id,
  );
  const type = enumValue(
    record.type,
    ['linear', 'path'] as const,
    'XLSX gradient fill type is invalid',
    id,
  );
  if (!Array.isArray(record.stops) || record.stops.length < 2) {
    invalid('XLSX gradient fill stops are invalid', id);
  }
  const stops = record.stops.map((stop) => gradientStop(stop, id));
  for (let index = 1; index < stops.length; index += 1) {
    if (stops[index]!.position <= stops[index - 1]!.position) {
      invalid('XLSX gradient stop positions are out of order', id);
    }
  }
  const result: Extract<XlsxFill, { kind: 'gradient' }> = {
    kind: 'gradient',
    stops,
    type,
  };
  if (record.angle !== undefined) {
    if (type !== 'linear') invalid('XLSX path gradient angle is invalid', id);
    result.angle = finite(
      record.angle,
      Number.MIN_VALUE,
      360,
      'XLSX gradient angle is invalid',
      id,
    );
  }
  for (const name of ['bottom', 'left', 'right', 'top'] as const) {
    if (record[name] !== undefined) {
      if (type !== 'path')
        invalid('XLSX linear gradient path bounds are invalid', id);
      result[name] = finite(
        record[name],
        Number.MIN_VALUE,
        1,
        'XLSX gradient path bound is invalid',
        id,
      );
    }
  }
  return result;
}

function font(value: unknown, id: string): XlsxFont {
  const record = plain(value, 'XLSX style font shape is invalid', id);
  exact(record, FONT_KEYS, 'XLSX style font shape is invalid', id);
  const result: XlsxFont = {};
  for (const name of [
    'bold',
    'condense',
    'extend',
    'italic',
    'outline',
    'shadow',
    'strike',
  ] as const) {
    if (record[name] !== undefined)
      result[name] = normalizedTrue(
        record[name],
        `XLSX style font ${name} flag is invalid`,
        id,
      );
  }
  if (record.charset !== undefined)
    result.charset = integer(
      record.charset,
      0,
      255,
      'XLSX style font charset is invalid',
      id,
    );
  if (record.color !== undefined) result.color = color(record.color, id);
  if (record.family !== undefined)
    result.family = integer(
      record.family,
      0,
      5,
      'XLSX style font family is invalid',
      id,
    );
  if (record.name !== undefined) {
    if (typeof record.name !== 'string' || record.name.length === 0)
      invalid('XLSX style font name is invalid', id);
    result.name = record.name;
  }
  if (record.scheme !== undefined)
    result.scheme = enumValue(
      record.scheme,
      ['major', 'minor'] as const,
      'XLSX style font scheme is invalid',
      id,
    );
  if (record.size !== undefined)
    result.size = finite(
      record.size,
      Number.MIN_VALUE,
      409,
      'XLSX style font size is invalid',
      id,
    );
  if (record.underline !== undefined)
    result.underline = enumValue(
      record.underline,
      ['double', 'double-accounting', 'single', 'single-accounting'] as const,
      'XLSX style font underline is invalid',
      id,
    );
  if (record.verticalAlignment !== undefined)
    result.verticalAlignment = enumValue(
      record.verticalAlignment,
      ['subscript', 'superscript'] as const,
      'XLSX style font vertical alignment is invalid',
      id,
    );
  if (Object.keys(result).length === 0)
    invalid('XLSX style font must be normalized', id);
  return result;
}

function protection(value: unknown, id: string): XlsxProtection {
  const record = plain(value, 'XLSX style protection shape is invalid', id);
  exact(record, PROTECTION_KEYS, 'XLSX style protection shape is invalid', id);
  const result: XlsxProtection = {};
  if (record.hidden !== undefined)
    result.hidden = normalizedTrue(
      record.hidden,
      'XLSX style protection hidden flag is invalid',
      id,
    );
  if (record.locked !== undefined) {
    if (record.locked !== false)
      invalid('XLSX style protection locked flag is invalid', id);
    result.locked = false;
  }
  if (Object.keys(result).length === 0)
    invalid('XLSX style protection must be normalized', id);
  return result;
}

export function validateXlsxOperationStyle(
  value: unknown,
  id: string,
): XlsxStyle {
  const record = plain(value, 'XLSX set-cell-style style shape is invalid', id);
  exact(record, STYLE_KEYS, 'XLSX set-cell-style style shape is invalid', id);
  const result: XlsxStyle = {};
  if (record.alignment !== undefined)
    result.alignment = alignment(record.alignment, id);
  if (record.border !== undefined) result.border = border(record.border, id);
  if (record.checkbox !== undefined)
    result.checkbox = normalizedTrue(
      record.checkbox,
      'XLSX checkbox style flag is invalid',
      id,
    );
  if (record.fill !== undefined) result.fill = fill(record.fill, id);
  if (record.font !== undefined) result.font = font(record.font, id);
  if (record.numberFormat !== undefined) {
    if (
      typeof record.numberFormat !== 'string' ||
      record.numberFormat.length === 0
    ) {
      invalid('XLSX style number format is invalid', id);
    }
    result.numberFormat = record.numberFormat;
  }
  if (record.protection !== undefined)
    result.protection = protection(record.protection, id);
  return result;
}
