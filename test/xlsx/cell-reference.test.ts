import fc from 'fast-check';
import { describe, expect, it } from 'vitest';

import {
  parseXlsxCellReference,
  parseXlsxRangeReference,
  xlsxColumnName,
} from '../../src/formats/xlsx/internal/cell-reference';
import {
  XLSX_MAX_COLUMNS,
  XLSX_MAX_ROWS,
} from '../../src/formats/xlsx/internal/resource-limits';

describe('XLSX A1 references', () => {
  it.each([
    [1, 'A'],
    [24, 'X'],
    [26, 'Z'],
    [27, 'AA'],
    [52, 'AZ'],
    [702, 'ZZ'],
    [703, 'AAA'],
    [XLSX_MAX_COLUMNS, 'XFD'],
  ])('formats column %i as %s', (column, name) => {
    expect(xlsxColumnName(column)).toBe(name);
  });

  it.each([
    Number.NEGATIVE_INFINITY,
    -1,
    0,
    1.5,
    XLSX_MAX_COLUMNS + 1,
    Number.POSITIVE_INFINITY,
    Number.NaN,
  ])('rejects invalid column index %s', (column) => {
    expect(xlsxColumnName(column)).toBeUndefined();
  });

  it.each([
    ['A1', 'A1', 1, 1, false, false],
    ['xFd1048576', 'XFD1048576', 16_384, 1_048_576, false, false],
    ['$B3', 'B3', 2, 3, true, false],
    ['C$4', 'C4', 3, 4, false, true],
    ['$D$5', 'D5', 4, 5, true, true],
  ] as const)(
    'parses cell reference %s',
    (value, address, column, row, absoluteColumn, absoluteRow) => {
      expect(parseXlsxCellReference(value)).toEqual({
        absoluteColumn,
        absoluteRow,
        address,
        column,
        row,
      });
    },
  );

  it.each([
    undefined,
    null,
    1,
    '',
    ' A1',
    'A1 ',
    'A',
    '1',
    'A0',
    'A01',
    'A+1',
    'A1.0',
    'XFE1',
    'A1048577',
    'A9007199254740992',
    'A:A',
    '1:1',
    'A1:B2',
    '$$A1',
    'A$$1',
    'Ä1',
  ])('rejects invalid cell reference %s', (value) => {
    expect(parseXlsxCellReference(value)).toBeUndefined();
  });

  it.each([
    [
      'A1:B2',
      {
        end: { column: 2, row: 2 },
        reference: 'A1:B2',
        start: { column: 1, row: 1 },
      },
    ],
    [
      '$b$2:$C$4',
      {
        end: { column: 3, row: 4 },
        reference: 'B2:C4',
        start: { column: 2, row: 2 },
      },
    ],
    [
      'D5',
      {
        end: { column: 4, row: 5 },
        reference: 'D5',
        start: { column: 4, row: 5 },
      },
    ],
  ] as const)('parses range reference %s', (value, expected) => {
    expect(parseXlsxRangeReference(value)).toEqual(expected);
  });

  it.each([
    undefined,
    null,
    1,
    '',
    ':',
    'A1:',
    ':B2',
    'A1:B2:C3',
    'B1:A1',
    'A2:A1',
    'XFE1:XFE2',
    'A0:A1',
    'A1 B2',
  ])('rejects invalid range reference %s', (value) => {
    expect(parseXlsxRangeReference(value)).toBeUndefined();
  });

  it('round-trips every generated in-grid coordinate', () => {
    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: XLSX_MAX_COLUMNS }),
        fc.integer({ min: 1, max: XLSX_MAX_ROWS }),
        fc.boolean(),
        fc.boolean(),
        (column, row, absoluteColumn, absoluteRow) => {
          const name = xlsxColumnName(column)!;
          const value = `${absoluteColumn ? '$' : ''}${name}${
            absoluteRow ? '$' : ''
          }${row}`;
          expect(parseXlsxCellReference(value)).toEqual({
            absoluteColumn,
            absoluteRow,
            address: `${name}${row}`,
            column,
            row,
          });
        },
      ),
      { numRuns: 500, seed: 2_026_081_702 },
    );
  });
});
