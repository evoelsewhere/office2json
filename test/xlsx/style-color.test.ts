import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { parseXlsxStyleColor } from '../../src/formats/xlsx/internal/style-color';

const PART = 'xl/styles.xml';
const CONTEXT = 'Font';

function color(attrs: Record<string, unknown>): unknown {
  return { attrs };
}

function capture(value: unknown): XlsxParseError {
  try {
    parseXlsxStyleColor(value, PART, CONTEXT);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected style color parsing to fail');
}

describe('XLSX style colors', () => {
  it.each([
    [undefined, undefined],
    [{}, undefined],
    [color({ auto: '0' }), undefined],
    [color({ auto: 'false' }), undefined],
    [color({ auto: '1' }), { kind: 'automatic' }],
    [color({ auto: 'true' }), { kind: 'automatic' }],
    [color({ rgb: 'ff00aBcD' }), { argb: 'FF00ABCD', kind: 'rgb' }],
    [color({ rgb: 'FF000000', tint: '0' }), { argb: 'FF000000', kind: 'rgb' }],
    [
      color({ theme: '0', tint: '+.25' }),
      { index: 0, kind: 'theme', tint: 0.25 },
    ],
    [
      color({ theme: '11', tint: '-1' }),
      { index: 11, kind: 'theme', tint: -1 },
    ],
    [color({ theme: '1', tint: '1.' }), { index: 1, kind: 'theme', tint: 1 }],
    [
      color({ theme: '2', tint: '00.5' }),
      { index: 2, kind: 'theme', tint: 0.5 },
    ],
    [
      color({ indexed: '0', tint: '1.0' }),
      { index: 0, kind: 'indexed', tint: 1 },
    ],
    [color({ indexed: '65', tint: '-0' }), { index: 65, kind: 'indexed' }],
  ] as const)('normalizes portable color %#', (source, expected) => {
    const result = parseXlsxStyleColor(source, PART, CONTEXT);
    expect(result).toEqual(expected);
    if (result) {
      expect(Object.isFrozen(result)).toBe(true);
      expect(JSON.parse(JSON.stringify(result))).toEqual(result);
    }
  });

  it.each([
    [null, 'invalid-document-structure', 'Font color element is invalid'],
    ['color', 'invalid-document-structure', 'Font color element is invalid'],
    [
      { attrs: 'bad' },
      'invalid-document-structure',
      'Font color attributes are invalid',
    ],
    [
      color({ rgb: 'FF000000', theme: '1' }),
      'invalid-document-structure',
      'Font color has multiple selectors',
    ],
    [
      color({ indexed: '1', auto: 'true' }),
      'invalid-document-structure',
      'Font color has multiple selectors',
    ],
    [
      color({ rgb: 'FFFFFF' }),
      'invalid-document-value',
      'Font color RGB is invalid',
    ],
    [
      color({ rgb: 'GG000000' }),
      'invalid-document-value',
      'Font color RGB is invalid',
    ],
    [
      color({ rgb: '0FF000000' }),
      'invalid-document-value',
      'Font color RGB is invalid',
    ],
    [
      color({ rgb: 'FF0000000' }),
      'invalid-document-value',
      'Font color RGB is invalid',
    ],
    [
      color({ theme: '-1' }),
      'invalid-document-value',
      'Font theme-color index is invalid',
    ],
    [
      color({ theme: '01' }),
      'invalid-document-value',
      'Font theme-color index is invalid',
    ],
    [
      color({ theme: '12' }),
      'invalid-document-value',
      'Font theme-color index is invalid',
    ],
    [
      color({ indexed: '66' }),
      'invalid-document-value',
      'Font indexed-color index is invalid',
    ],
    [
      color({ auto: 'yes' }),
      'invalid-document-value',
      'Font automatic-color flag is invalid',
    ],
    [
      color({ tint: '.5' }),
      'invalid-document-value',
      'Font color tint has no base color',
    ],
    [
      color({ auto: 'true', tint: '.5' }),
      'invalid-document-value',
      'Font automatic color cannot have a tint',
    ],
    [
      color({ rgb: 'FF000000', tint: 0.5 }),
      'invalid-document-value',
      'Font color tint is invalid',
    ],
    [
      color({ rgb: 'FF000000', tint: 'x.5' }),
      'invalid-document-value',
      'Font color tint is invalid',
    ],
    [
      color({ rgb: 'FF000000', tint: '.5x' }),
      'invalid-document-value',
      'Font color tint is invalid',
    ],
    [
      color({ rgb: 'FF000000', tint: ' .5' }),
      'invalid-document-value',
      'Font color tint is invalid',
    ],
    [
      color({ rgb: 'FF000000', tint: '.5 ' }),
      'invalid-document-value',
      'Font color tint is invalid',
    ],
    [
      color({ rgb: 'FF000000', tint: 'NaN' }),
      'invalid-document-value',
      'Font color tint is invalid',
    ],
    [
      color({ rgb: 'FF000000', tint: '1.1' }),
      'invalid-document-value',
      'Font color tint is invalid',
    ],
    [
      color({ rgb: 'FF000000', tint: '-1.1' }),
      'invalid-document-value',
      'Font color tint is invalid',
    ],
  ] as const)('rejects invalid style color %#', (source, code, message) => {
    expect(capture(source).diagnostic).toEqual({
      code,
      message,
      part: PART,
      severity: 'error',
    });
  });

  it('uses the caller-owned context in diagnostics', () => {
    expect(capture({ attrs: { indexed: '66' } }).diagnostic.message).toBe(
      'Font indexed-color index is invalid',
    );
    expect(
      captureWithContext({ attrs: { indexed: '66' } }, 'Border left').diagnostic
        .message,
    ).toBe('Border left indexed-color index is invalid');
  });
});

function captureWithContext(value: unknown, context: string): XlsxParseError {
  try {
    parseXlsxStyleColor(value, PART, context);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected contextual style color parsing to fail');
}
