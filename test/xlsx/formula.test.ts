import { describe, expect, it } from 'vitest';

import { translateXlsxSharedFormula } from '../../src/formats/xlsx/internal/formula';

describe('XLSX shared-formula translation', () => {
  it('shifts relative A1 cells and ranges from the master coordinate', () => {
    expect(
      translateXlsxSharedFormula(
        'A1+B2:C3',
        { column: 2, row: 2 },
        {
          column: 4,
          row: 5,
        },
      ),
    ).toBe('C4+D5:E6');
  });

  it('preserves absolute and shifts mixed references independently', () => {
    expect(
      translateXlsxSharedFormula(
        '$A$1+A$1+$A1',
        { column: 1, row: 1 },
        { column: 3, row: 4 },
      ),
    ).toBe('$A$1+C$1+$A4');
  });

  it('preserves strings and escaped quotes without treating their text as references', () => {
    expect(
      translateXlsxSharedFormula(
        'IF(A1="A1""B2",B2,C3)',
        { column: 1, row: 1 },
        { column: 2, row: 2 },
      ),
    ).toBe('IF(B2="A1""B2",C3,D4)');
  });

  it('preserves quoted, unquoted, 3D, and external qualifiers while shifting their references', () => {
    expect(
      translateXlsxSharedFormula(
        "'A1':Sheet2!A1+[Book.xlsx]A1!B2+Data!C3",
        { column: 1, row: 1 },
        { column: 2, row: 3 },
      ),
    ).toBe("'A1':Sheet2!B3+[Book.xlsx]A1!C4+Data!D5");
  });

  it('preserves apostrophe escapes in quoted sheet names', () => {
    expect(
      translateXlsxSharedFormula(
        "'Owner''s A1'!A1",
        { column: 1, row: 1 },
        { column: 2, row: 2 },
      ),
    ).toBe("'Owner''s A1'!B2");
  });

  it('preserves structured-reference brackets and cell-like column names', () => {
    expect(
      translateXlsxSharedFormula(
        'Table1[[#This Row],[A1]]+A1',
        { column: 1, row: 1 },
        { column: 2, row: 2 },
      ),
    ).toBe('Table1[[#This Row],[A1]]+B2');
  });

  it('does not translate identifiers, function names, or cell-like sheet names', () => {
    expect(
      translateXlsxSharedFormula(
        'LOG10(A1)+name.A1+A1foo+A1!B2+A1:B2!C3',
        { column: 1, row: 1 },
        { column: 2, row: 2 },
      ),
    ).toBe('LOG10(B2)+name.A1+A1foo+A1!C3+A1:B2!D4');
  });

  it.each([
    ['A1!B2', 'A1!C3'],
    ['A1:B2!C3', 'A1:B2!D4'],
    ['A1[Column]+B2', 'A1[Column]+C3'],
  ])(
    'keeps cell-like qualifier syntax opaque in %s',
    (expression, expected) => {
      expect(
        translateXlsxSharedFormula(
          expression,
          { column: 1, row: 1 },
          { column: 2, row: 2 },
        ),
      ).toBe(expected);
    },
  );

  it('returns the source text exactly for the master coordinate', () => {
    expect(
      translateXlsxSharedFormula(
        'a1+"unterminated',
        { column: 1, row: 1 },
        {
          column: 1,
          row: 1,
        },
      ),
    ).toBe('a1+"unterminated');
  });

  it.each([
    [
      { column: 0, row: 1 },
      { column: 1, row: 1 },
    ],
    [
      { column: 1, row: 0 },
      { column: 1, row: 1 },
    ],
    [
      { column: 16_385, row: 1 },
      { column: 1, row: 1 },
    ],
    [
      { column: 1, row: 1_048_577 },
      { column: 1, row: 1 },
    ],
    [
      { column: 1.5, row: 1 },
      { column: 1, row: 1 },
    ],
    [
      { column: 1, row: 1 },
      { column: Number.NaN, row: 1 },
    ],
  ] as const)(
    'rejects invalid source or target coordinate %#',
    (source, target) => {
      expect(translateXlsxSharedFormula('A1', source, target)).toBeUndefined();
    },
  );

  it('accepts exact worksheet coordinate maxima', () => {
    expect(
      translateXlsxSharedFormula(
        'XFD1048576',
        { column: 16_384, row: 1_048_576 },
        { column: 16_384, row: 1_048_576 },
      ),
    ).toBe('XFD1048576');
  });

  it.each([
    { column: 16_385, row: 1 },
    { column: 1, row: 1_048_577 },
  ])('rejects identical coordinates above a hard maximum %#', (coordinate) => {
    expect(
      translateXlsxSharedFormula('$A$1', coordinate, coordinate),
    ).toBeUndefined();
  });

  it('allows a translated reference to land exactly on both grid maxima', () => {
    expect(
      translateXlsxSharedFormula(
        'XFC1048575',
        { column: 1, row: 1 },
        { column: 2, row: 2 },
      ),
    ).toBe('XFD1048576');
  });

  it.each([
    ['A1', { column: 2, row: 1 }, { column: 1, row: 1 }],
    ['XFD1', { column: 1, row: 1 }, { column: 2, row: 1 }],
    ['A1', { column: 1, row: 2 }, { column: 1, row: 1 }],
    ['A1048576', { column: 1, row: 1 }, { column: 1, row: 2 }],
  ] as const)(
    'rejects translated references outside the worksheet %#',
    (expression, source, target) => {
      expect(
        translateXlsxSharedFormula(expression, source, target),
      ).toBeUndefined();
    },
  );

  it.each(['"A1', "'Sheet!A1", 'Table[A1', 'Table]A1'])(
    'rejects unterminated opaque token %#',
    (expression) => {
      expect(
        translateXlsxSharedFormula(
          expression,
          { column: 1, row: 1 },
          { column: 2, row: 2 },
        ),
      ).toBeUndefined();
    },
  );

  it('preserves non-reference and out-of-grid identifier tokens', () => {
    expect(
      translateXlsxSharedFormula(
        '1+TRUE+XFE1+A1048577',
        { column: 1, row: 1 },
        { column: 2, row: 2 },
      ),
    ).toBe('1+TRUE+XFE1+A1048577');
  });
});
