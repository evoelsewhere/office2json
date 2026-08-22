import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { parseXlsxStyleBorder } from '../../src/formats/xlsx/internal/style-border';
import type { XlsxBorderStyle } from '../../src/formats/xlsx/types';

const PART = 'xl/styles.xml';

function borderSide(
  style?: unknown,
  colorAttrs?: Record<string, unknown>,
): unknown {
  return {
    ...(style === undefined ? {} : { attrs: { style } }),
    ...(colorAttrs === undefined ? {} : { color: { attrs: colorAttrs } }),
  };
}

function capture(value: unknown, prefix = ''): XlsxParseError {
  try {
    parseXlsxStyleBorder(value, prefix, PART);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected style border parsing to fail');
}

const BORDER_STYLES: readonly XlsxBorderStyle[] = [
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

describe('XLSX style borders', () => {
  it('normalizes every border side, color, and root flag into portable JSON', () => {
    const result = parseXlsxStyleBorder(
      {
        attrs: { diagonalDown: '1', diagonalUp: 'true', outline: 'false' },
        bottom: borderSide('dashed', { indexed: '1' }),
        diagonal: borderSide('dashDot', { auto: 'true' }),
        end: borderSide('dotted', { theme: '2' }),
        horizontal: borderSide('double', { rgb: 'FF010203' }),
        left: borderSide('thin', { rgb: 'ffabcdef' }),
        right: borderSide('medium', { theme: '4', tint: '.25' }),
        start: borderSide('hair', { indexed: '64' }),
        top: borderSide('thick'),
        vertical: borderSide(undefined, { rgb: 'FFFFFFFF' }),
      },
      '',
      PART,
    );

    expect(result).toEqual({
      bottom: { color: { index: 1, kind: 'indexed' }, style: 'dashed' },
      diagonal: { color: { kind: 'automatic' }, style: 'dashDot' },
      diagonalDown: true,
      diagonalUp: true,
      end: { color: { index: 2, kind: 'theme' }, style: 'dotted' },
      horizontal: {
        color: { argb: 'FF010203', kind: 'rgb' },
        style: 'double',
      },
      left: { color: { argb: 'FFABCDEF', kind: 'rgb' }, style: 'thin' },
      outline: false,
      right: {
        color: { index: 4, kind: 'theme', tint: 0.25 },
        style: 'medium',
      },
      start: { color: { index: 64, kind: 'indexed' }, style: 'hair' },
      top: { style: 'thick' },
      vertical: { color: { argb: 'FFFFFFFF', kind: 'rgb' } },
    });
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.left)).toBe(true);
    expect(Object.isFrozen(result.left?.color)).toBe(true);
    expect('style' in result.vertical!).toBe(false);
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });

  it('parses prefixed sides and applies border defaults', () => {
    expect(
      parseXlsxStyleBorder(
        {
          attrs: { diagonalDown: '0', diagonalUp: 'false', outline: 'true' },
          's:left': borderSide('none'),
          's:right': {},
          's:top': [borderSide('thin')],
        },
        's',
        PART,
      ),
    ).toEqual({ top: { style: 'thin' } });
  });

  it.each(BORDER_STYLES)('accepts border style %s', (style) => {
    expect(parseXlsxStyleBorder({ left: borderSide(style) }, '', PART)).toEqual(
      { left: { style } },
    );
  });

  it.each([
    ['diagonalDown', undefined, false],
    ['diagonalDown', '0', false],
    ['diagonalDown', 'false', false],
    ['diagonalDown', '1', true],
    ['diagonalDown', 'true', true],
    ['diagonalUp', '1', true],
    ['outline', undefined, false],
    ['outline', '1', false],
    ['outline', 'true', false],
    ['outline', '0', true],
    ['outline', 'false', true],
  ] as const)('normalizes root flag %s=%s', (name, source, present) => {
    const attrs = source === undefined ? {} : { [name]: source };
    const result = parseXlsxStyleBorder({ attrs }, '', PART);
    expect(name in result).toBe(present);
  });

  it.each([
    [null, 'invalid-document-structure', 'Border element is invalid'],
    [
      { attrs: 'bad' },
      'invalid-document-structure',
      'Border attributes are invalid',
    ],
    [
      { attrs: { diagonalDown: 'yes' } },
      'invalid-document-value',
      'Border diagonalDown value is invalid',
    ],
    [
      { attrs: { diagonalUp: 'yes' } },
      'invalid-document-value',
      'Border diagonalUp value is invalid',
    ],
    [
      { attrs: { outline: 'yes' } },
      'invalid-document-value',
      'Border outline value is invalid',
    ],
    [
      { left: [] },
      'invalid-document-structure',
      'Border left element is duplicated',
    ],
    [
      { left: [{}, {}] },
      'invalid-document-structure',
      'Border left element is duplicated',
    ],
    [
      { left: ['bad'] },
      'invalid-document-structure',
      'Border left element is invalid',
    ],
    [
      { left: 'bad' },
      'invalid-document-structure',
      'Border left element is invalid',
    ],
    [
      { left: { attrs: 'bad' } },
      'invalid-document-structure',
      'Border left attributes are invalid',
    ],
    [
      { left: borderSide('triple') },
      'invalid-document-value',
      'Border left style is invalid',
    ],
    [
      { left: borderSide(1) },
      'invalid-document-value',
      'Border left style is invalid',
    ],
    [
      { left: { color: [{}, {}] } },
      'invalid-document-structure',
      'Border left color element is duplicated',
    ],
    [
      { left: { color: { attrs: { theme: '12' } } } },
      'invalid-document-value',
      'Border left theme-color index is invalid',
    ],
    [
      { right: 'bad' },
      'invalid-document-structure',
      'Border right element is invalid',
    ],
    [
      { top: 'bad' },
      'invalid-document-structure',
      'Border top element is invalid',
    ],
    [
      { bottom: 'bad' },
      'invalid-document-structure',
      'Border bottom element is invalid',
    ],
    [
      { diagonal: 'bad' },
      'invalid-document-structure',
      'Border diagonal element is invalid',
    ],
    [
      { vertical: 'bad' },
      'invalid-document-structure',
      'Border vertical element is invalid',
    ],
    [
      { horizontal: 'bad' },
      'invalid-document-structure',
      'Border horizontal element is invalid',
    ],
    [
      { start: 'bad' },
      'invalid-document-structure',
      'Border start element is invalid',
    ],
    [
      { end: 'bad' },
      'invalid-document-structure',
      'Border end element is invalid',
    ],
  ] as const)('rejects invalid border %#', (source, code, message) => {
    expect(capture(source).diagnostic).toEqual({
      code,
      message,
      part: PART,
      severity: 'error',
    });
  });
});
