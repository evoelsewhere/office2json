import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { parseXlsxScalarCellValue } from '../../src/formats/xlsx/internal/cell-value';
import type { XlsxSharedStringTable } from '../../src/formats/xlsx/internal/shared-strings';

const PART = 'xl/worksheets/sheet1.xml';
const CELL = 'B7';
const EMPTY_STRINGS: XlsxSharedStringTable = Object.freeze({
  part: null,
  values: Object.freeze([]),
});
const SHARED_STRINGS: XlsxSharedStringTable = Object.freeze({
  part: 'xl/sharedStrings.xml',
  values: Object.freeze([
    Object.freeze({ text: 'plain' }),
    Object.freeze({
      runs: Object.freeze([
        Object.freeze({ text: 'Rich' }),
        Object.freeze({ text: ' text' }),
      ]),
      text: 'Rich text',
    }),
  ]),
});

type CellType = Parameters<typeof parseXlsxScalarCellValue>[0];

function parse(
  type: CellType,
  value: string,
  strings: XlsxSharedStringTable = EMPTY_STRINGS,
) {
  return parseXlsxScalarCellValue(type, value, strings, PART, CELL);
}

function captureParseError(
  type: CellType,
  value: string,
  strings: XlsxSharedStringTable = EMPTY_STRINGS,
): XlsxParseError {
  try {
    parse(type, value, strings);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected cell value parsing to fail');
}

describe('XLSX scalar cell values', () => {
  it.each([
    ['0', 0],
    ['-0', 0],
    ['+0', 0],
    ['1', 1],
    ['+1', 1],
    ['12', 12],
    ['1.', 1],
    ['.5', 0.5],
    ['.55', 0.55],
    ['-.5', -0.5],
    ['1e3', 1000],
    ['-2.5E-2', -0.025],
    ['1.7976931348623157e308', Number.MAX_VALUE],
  ] as const)('parses finite decimal number %s', (source, expected) => {
    const result = parse('n', source);
    expect(result).toEqual({ kind: 'number', value: expected });
    if (source === '-0') {
      expect(Object.is((result as { value: number }).value, -0)).toBe(false);
    }
  });

  it.each([
    '',
    ' ',
    '\t1',
    '1\n',
    'NaN',
    'Infinity',
    '-Infinity',
    '0x10',
    '1,2',
    '1_0',
    '.',
    '+',
    '1e',
    '1e+',
    '--1',
  ])('rejects invalid number lexical form %#', (source) => {
    expect(captureParseError('n', source).diagnostic).toEqual({
      cell: CELL,
      code: 'invalid-document-value',
      message: 'Cell number is invalid',
      part: PART,
      severity: 'error',
    });
  });

  it.each(['1e309', '-1e309'])('rejects non-finite decimal %s', (source) => {
    expect(captureParseError('n', source).diagnostic).toEqual({
      cell: CELL,
      code: 'invalid-document-value',
      message: 'Cell number is outside the finite range',
      part: PART,
      severity: 'error',
    });
  });

  it('parses canonical booleans', () => {
    expect(parse('b', '0')).toEqual({ kind: 'boolean', value: false });
    expect(parse('b', '1')).toEqual({ kind: 'boolean', value: true });
  });

  it.each(['', '00', '01', '-1', '2', 'false', 'true', ' 1'])(
    'rejects invalid boolean lexical form %#',
    (source) => {
      expect(captureParseError('b', source).diagnostic).toEqual({
        cell: CELL,
        code: 'invalid-document-value',
        message: 'Cell boolean is invalid',
        part: PART,
        severity: 'error',
      });
    },
  );

  it.each([
    '#BLOCKED!',
    '#BUSY!',
    '#CALC!',
    '#CONNECT!',
    '#DIV/0!',
    '#FIELD!',
    '#GETTING_DATA',
    '#N/A',
    '#NAME?',
    '#NULL!',
    '#NUM!',
    '#REF!',
    '#SPILL!',
    '#UNKNOWN!',
    '#VALUE!',
  ])('preserves recognized error code %s', (code) => {
    expect(parse('e', code)).toEqual({ code, kind: 'error' });
  });

  it.each(['', '#DIV/0', '#div/0!', '#ERROR!', ' #N/A'])(
    'rejects invalid error code %#',
    (source) => {
      expect(captureParseError('e', source).diagnostic).toEqual({
        cell: CELL,
        code: 'invalid-document-value',
        message: 'Cell error code is invalid',
        part: PART,
        severity: 'error',
      });
    },
  );

  it.each(['', ' ', ' A & B ', '東京', '🙂'])(
    'preserves explicit text %#',
    (source) => {
      expect(parse('str', source)).toEqual({ kind: 'text', text: source });
    },
  );

  it('resolves plain and rich shared strings without exposing table arrays', () => {
    expect(parse('s', '0', SHARED_STRINGS)).toEqual({
      kind: 'text',
      text: 'plain',
    });
    const rich = parse('s', '1', SHARED_STRINGS);
    expect(rich).toEqual({
      kind: 'text',
      runs: [{ text: 'Rich' }, { text: ' text' }],
      text: 'Rich text',
    });
    expect((rich as { runs: unknown }).runs).not.toBe(
      SHARED_STRINGS.values[1]?.runs,
    );
  });

  it.each(['', '-1', '+1', '00', '01', '1.0', ' 0', '0 ', '1e0'])(
    'rejects invalid shared-string index %#',
    (source) => {
      expect(captureParseError('s', source, SHARED_STRINGS).diagnostic).toEqual(
        {
          cell: CELL,
          code: 'invalid-document-value',
          message: 'Cell shared-string index is invalid',
          part: PART,
          severity: 'error',
        },
      );
    },
  );

  it.each([
    ['0', EMPTY_STRINGS],
    ['2', SHARED_STRINGS],
    ['9007199254740992', SHARED_STRINGS],
  ] as const)(
    'rejects shared-string index %s outside the table',
    (source, table) => {
      expect(captureParseError('s', source, table).diagnostic).toEqual({
        cell: CELL,
        code: 'invalid-document-value',
        message: 'Cell shared-string index is out of range',
        part: PART,
        severity: 'error',
      });
    },
  );

  it.each([
    ['2024-02-29', 'date'],
    ['2000-02-29', 'date'],
    ['0001-01-01', 'date'],
    ['2010-01-01', 'date'],
    ['12024-12-31', 'date'],
    ['99999999999999999999999999992000-02-29', 'date'],
    ['2024-01-31T23:59:59', 'date-time'],
    ['2024-01-31T23:59:59.123456789Z', 'date-time'],
    ['2024-01-31T23:59:59+14:00', 'date-time'],
    ['2024-01-31T23:59:59-05:30', 'date-time'],
    ['2024-01-31T23:59:59+09:45', 'date-time'],
    ['2024-01-31T23:59:59+10:45', 'date-time'],
    ['2024-01-31T23:59:59+13:59', 'date-time'],
    ['00:00:00', 'time'],
    ['23:59:59.123Z', 'time'],
    ['12:34:56+00:00', 'time'],
    ['P1Y', 'duration'],
    ['P2M', 'duration'],
    ['P3D', 'duration'],
    ['PT4H', 'duration'],
    ['PT5M', 'duration'],
    ['PT0S', 'duration'],
    ['P1Y2M3DT4H5M6.7S', 'duration'],
    ['P12Y23M34DT45H56M67.89S', 'duration'],
    ['-PT1H', 'duration'],
  ] as const)('preserves direct ISO value %s as %s', (source, precision) => {
    expect(parse('d', source)).toEqual({
      kind: 'date',
      normalized: source,
      precision,
      source: { kind: 'iso', value: source },
    });
  });

  it.each([
    '',
    'x2024-01-01',
    '0000-01-01',
    '9007199254740993-02-29',
    '2024-00-01',
    '2024-13-01',
    '2024-01-00',
    '2024-01-32',
    '2023-02-29',
    '1900-02-29',
    '2000-02-30',
    '2024-04-31',
    '2024-06-31',
    '2024-09-31',
    '2024-11-31',
    '2024-01-01t00:00:00',
    '2024-01-01T',
    '2024-01-01T24:00:00',
    '2024-01-01T23:60:00',
    '2024-01-01T23:59:60',
    '2024-01-01T23:59:59+14:01',
    '2024-01-01T23:59:59+15:00',
    '2024-01-01T23:59:59+00:60',
    '2024-01-01T23:59:59+99:00',
    '24:00:00',
    '12:60:00',
    '12:00:60',
    '12:00:00z',
    '12:00:00+1:00',
    'P',
    'PT',
    'P1W',
    'P1.5D',
    'PT1.5H',
    'PT1.5M',
    'PT.5S',
    '+PT1H',
  ])('rejects invalid direct ISO value %#', (source) => {
    expect(captureParseError('d', source).diagnostic).toEqual({
      cell: CELL,
      code: 'invalid-document-value',
      message: 'Cell ISO date is invalid',
      part: PART,
      severity: 'error',
    });
  });

  it('produces portable JSON for every scalar discriminant', () => {
    const values = [
      parse('str', 'text'),
      parse('n', '1.25'),
      parse('b', '1'),
      parse('e', '#N/A'),
      parse('d', '2024-01-02'),
    ];

    expect(JSON.parse(JSON.stringify(values))).toEqual(values);
  });
});
