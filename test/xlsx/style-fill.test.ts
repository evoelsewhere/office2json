import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { parseXlsxStyleFill } from '../../src/formats/xlsx/internal/style-fill';
import type { XlsxPatternType } from '../../src/formats/xlsx/types';

const PART = 'xl/styles.xml';

function color(attrs: Record<string, unknown> = {}): unknown {
  return { attrs };
}

function stop(
  position: unknown,
  colorValue: unknown = color({ rgb: 'FF000000' }),
) {
  return { attrs: { position }, color: colorValue };
}

function gradient(
  stops: unknown[],
  attrs: Record<string, unknown> = {},
): unknown {
  return { gradientFill: { attrs, stop: stops } };
}

function capture(value: unknown, prefix = ''): XlsxParseError {
  try {
    parseXlsxStyleFill(value, prefix, PART);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected style fill parsing to fail');
}

const PATTERNS: readonly XlsxPatternType[] = [
  'none',
  'solid',
  'mediumGray',
  'darkGray',
  'lightGray',
  'darkHorizontal',
  'darkVertical',
  'darkDown',
  'darkUp',
  'darkGrid',
  'darkTrellis',
  'lightHorizontal',
  'lightVertical',
  'lightDown',
  'lightUp',
  'lightGrid',
  'lightTrellis',
  'gray125',
  'gray0625',
];

describe('XLSX style fills', () => {
  it.each([
    [{}, { kind: 'pattern', pattern: 'none' }],
    [{ patternFill: {} }, { kind: 'pattern', pattern: 'none' }],
    [
      {
        patternFill: {
          attrs: { patternType: 'solid' },
          bgColor: color({ indexed: '64' }),
          fgColor: color({ rgb: 'ffabcdef' }),
        },
      },
      {
        backgroundColor: { index: 64, kind: 'indexed' },
        foregroundColor: { argb: 'FFABCDEF', kind: 'rgb' },
        kind: 'pattern',
        pattern: 'solid',
      },
    ],
  ] as const)('normalizes pattern fill %#', (source, expected) => {
    const result = parseXlsxStyleFill(source, '', PART);
    expect(result).toEqual(expected);
    expect(Object.isFrozen(result)).toBe(true);
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });

  it.each(PATTERNS)('accepts pattern type %s', (patternType) => {
    expect(
      parseXlsxStyleFill({ patternFill: { attrs: { patternType } } }, '', PART),
    ).toEqual({ kind: 'pattern', pattern: patternType });
  });

  it('parses prefixed pattern children', () => {
    expect(
      parseXlsxStyleFill(
        {
          's:patternFill': {
            attrs: { patternType: 'solid' },
            's:fgColor': color({ theme: '2', tint: '-.25' }),
          },
        },
        's',
        PART,
      ),
    ).toEqual({
      foregroundColor: { index: 2, kind: 'theme', tint: -0.25 },
      kind: 'pattern',
      pattern: 'solid',
    });
  });

  it.each([
    ['0', undefined],
    ['45', 45],
    ['1.', 1],
    ['1.25', 1.25],
    ['.25', 0.25],
    ['00.5', 0.5],
    ['360', 360],
  ] as const)('normalizes linear gradient degree %s', (degree, expected) => {
    const result = parseXlsxStyleFill(
      gradient([stop('0'), stop('1', color({ theme: '4', tint: '.5' }))], {
        degree,
      }),
      '',
      PART,
    );
    expect(result).toEqual({
      ...(expected === undefined ? {} : { angle: expected }),
      kind: 'gradient',
      stops: [
        { color: { argb: 'FF000000', kind: 'rgb' }, position: 0 },
        { color: { index: 4, kind: 'theme', tint: 0.5 }, position: 1 },
      ],
      type: 'linear',
    });
    expect(Object.isFrozen(result)).toBe(true);
    if (result.kind === 'gradient') {
      expect(Object.isFrozen(result.stops)).toBe(true);
      expect(Object.isFrozen(result.stops[0])).toBe(true);
    }
  });

  it('normalizes prefixed path gradient bounds and ordered stops', () => {
    const result = parseXlsxStyleFill(
      {
        's:gradientFill': {
          attrs: {
            bottom: '1',
            left: '.1',
            right: '.2',
            top: '.3',
            type: 'path',
          },
          's:stop': [
            {
              attrs: { position: '0' },
              's:color': color({ auto: 'true' }),
            },
            {
              attrs: { position: '.5' },
              's:color': color({ indexed: '1' }),
            },
            {
              attrs: { position: '1' },
              's:color': color({ rgb: 'FFFFFFFF' }),
            },
          ],
        },
      },
      's',
      PART,
    );

    expect(result).toEqual({
      bottom: 1,
      kind: 'gradient',
      left: 0.1,
      right: 0.2,
      stops: [
        { color: { kind: 'automatic' }, position: 0 },
        { color: { index: 1, kind: 'indexed' }, position: 0.5 },
        { color: { argb: 'FFFFFFFF', kind: 'rgb' }, position: 1 },
      ],
      top: 0.3,
      type: 'path',
    });
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });

  it.each([
    [null, 'invalid-document-structure', 'Fill element is invalid'],
    [
      { patternFill: [] },
      'invalid-document-structure',
      'Fill patternFill element is duplicated',
    ],
    [
      { patternFill: [{}, {}] },
      'invalid-document-structure',
      'Fill patternFill element is duplicated',
    ],
    [
      { patternFill: ['bad'] },
      'invalid-document-structure',
      'Fill patternFill element is invalid',
    ],
    [
      { patternFill: {}, gradientFill: {} },
      'invalid-document-structure',
      'Fill has both pattern and gradient definitions',
    ],
    [
      { patternFill: { attrs: 'bad' } },
      'invalid-document-structure',
      'Pattern fill attributes are invalid',
    ],
    [
      { patternFill: { attrs: { patternType: '' } } },
      'invalid-document-value',
      'Pattern fill type is invalid',
    ],
    [
      { patternFill: { attrs: { patternType: 'dots' } } },
      'invalid-document-value',
      'Pattern fill type is invalid',
    ],
    [
      { patternFill: { attrs: { patternType: 1 } } },
      'invalid-document-value',
      'Pattern fill type is invalid',
    ],
    [
      { patternFill: { fgColor: [{}, {}] } },
      'invalid-document-structure',
      'Pattern fill fgColor element is duplicated',
    ],
    [
      { patternFill: { bgColor: 'bad' } },
      'invalid-document-structure',
      'Pattern fill bgColor element is invalid',
    ],
    [
      { patternFill: { fgColor: color({ theme: '12' }) } },
      'invalid-document-value',
      'Pattern foreground theme-color index is invalid',
    ],
    [
      { patternFill: { bgColor: color({ theme: '12' }) } },
      'invalid-document-value',
      'Pattern background theme-color index is invalid',
    ],
    [
      gradient([stop('0'), stop('1')], { type: 'radial' }),
      'invalid-document-value',
      'Gradient fill type is invalid',
    ],
    [
      { gradientFill: { attrs: 'bad' } },
      'invalid-document-structure',
      'Gradient fill attributes are invalid',
    ],
    [
      { gradientFill: 'bad' },
      'invalid-document-structure',
      'Fill gradientFill element is invalid',
    ],
    [
      { gradientFill: {} },
      'invalid-document-structure',
      'Gradient fill requires at least two stops',
    ],
    [
      gradient([stop('0')]),
      'invalid-document-structure',
      'Gradient fill requires at least two stops',
    ],
    [
      { gradientFill: { stop: stop('0') } },
      'invalid-document-structure',
      'Gradient fill requires at least two stops',
    ],
    [
      gradient([stop('0'), 'bad']),
      'invalid-document-structure',
      'Gradient stop element is invalid',
    ],
    [
      gradient([stop(undefined), stop('1')]),
      'invalid-document-value',
      'Gradient stop position is missing',
    ],
    [
      gradient([{ attrs: 'bad', color: color() }, stop('1')]),
      'invalid-document-structure',
      'Gradient stop attributes are invalid',
    ],
    [
      gradient([stop(0), stop('1')]),
      'invalid-document-value',
      'Gradient stop position is invalid',
    ],
    [
      gradient([stop(' 0'), stop('1')]),
      'invalid-document-value',
      'Gradient stop position is invalid',
    ],
    [
      gradient([stop('0 '), stop('1')]),
      'invalid-document-value',
      'Gradient stop position is invalid',
    ],
    [
      gradient([stop('-.1'), stop('1')]),
      'invalid-document-value',
      'Gradient stop position is invalid',
    ],
    [
      gradient([stop('0'), stop('1.1')]),
      'invalid-document-value',
      'Gradient stop position is invalid',
    ],
    [
      gradient([stop('0'), stop('0')]),
      'invalid-document-value',
      'Gradient stop positions are out of order',
    ],
    [
      gradient([stop('.5'), stop('.25')]),
      'invalid-document-value',
      'Gradient stop positions are out of order',
    ],
    [
      gradient([{ attrs: { position: '0' } }, stop('1')]),
      'invalid-document-value',
      'Gradient stop color is missing',
    ],
    [
      gradient([{ attrs: { position: '0' }, color: [{}, {}] }, stop('1')]),
      'invalid-document-structure',
      'Gradient stop color element is duplicated',
    ],
    [
      gradient([stop('0', color({ theme: '12' })), stop('1')]),
      'invalid-document-value',
      'Gradient stop theme-color index is invalid',
    ],
    [
      gradient([stop('0'), stop('1')], { degree: 1 }),
      'invalid-document-value',
      'Gradient fill degree is invalid',
    ],
    [
      gradient([stop('0'), stop('1')], { degree: ' 1' }),
      'invalid-document-value',
      'Gradient fill degree is invalid',
    ],
    [
      gradient([stop('0'), stop('1')], { degree: '1 ' }),
      'invalid-document-value',
      'Gradient fill degree is invalid',
    ],
    [
      gradient([stop('0'), stop('1')], { degree: '-1' }),
      'invalid-document-value',
      'Gradient fill degree is invalid',
    ],
    [
      gradient([stop('0'), stop('1')], { degree: '360.1' }),
      'invalid-document-value',
      'Gradient fill degree is invalid',
    ],
    [
      gradient([stop('0'), stop('1')], { left: '.1' }),
      'invalid-document-value',
      'Linear gradient fill has path bounds',
    ],
    [
      gradient([stop('0'), stop('1')], { right: '.1' }),
      'invalid-document-value',
      'Linear gradient fill has path bounds',
    ],
    [
      gradient([stop('0'), stop('1')], { top: '.1' }),
      'invalid-document-value',
      'Linear gradient fill has path bounds',
    ],
    [
      gradient([stop('0'), stop('1')], { bottom: '.1' }),
      'invalid-document-value',
      'Linear gradient fill has path bounds',
    ],
    [
      gradient([stop('0'), stop('1')], { type: 'path', degree: '1' }),
      'invalid-document-value',
      'Path gradient fill has a degree',
    ],
    [
      gradient([stop('0'), stop('1')], { type: 'path', left: '-.1' }),
      'invalid-document-value',
      'Gradient fill left bound is invalid',
    ],
    [
      gradient([stop('0'), stop('1')], { type: 'path', right: '1.1' }),
      'invalid-document-value',
      'Gradient fill right bound is invalid',
    ],
    [
      gradient([stop('0'), stop('1')], { type: 'path', top: 'x' }),
      'invalid-document-value',
      'Gradient fill top bound is invalid',
    ],
    [
      gradient([stop('0'), stop('1')], { type: 'path', bottom: 'NaN' }),
      'invalid-document-value',
      'Gradient fill bottom bound is invalid',
    ],
  ] as const)('rejects invalid fill %#', (source, code, message) => {
    expect(capture(source).diagnostic).toEqual({
      code,
      message,
      part: PART,
      severity: 'error',
    });
  });
});
