import { describe, expect, it } from 'vitest';

import { transformXlsxStructuralSourceFormula } from '../../src/formats/xlsx/roundtrip/formula-reference';

describe('XLSX structural source formulas', () => {
  it('transforms local cells and ranges while preserving absolute markers', () => {
    expect(
      transformXlsxStructuralSourceFormula('$A$1:B3', 'Data', 'Data', {
        count: 2,
        index: 2,
        kind: 'insert-rows',
      }),
    ).toEqual({ expression: '$A$1:B5', kind: 'transformed' });
    expect(
      transformXlsxStructuralSourceFormula('B2', 'Data', 'Data', {
        count: 1,
        index: 2,
        kind: 'insert-columns',
      }),
    ).toEqual({ expression: 'C2', kind: 'transformed' });
  });

  it('matches quoted and case-insensitive sheet qualifiers', () => {
    expect(
      transformXlsxStructuralSourceFormula(
        "'Owner''s Data'!$A1:$A3",
        'Other',
        "OWNER'S DATA",
        { count: 1, index: 2, kind: 'delete-rows' },
      ),
    ).toEqual({
      expression: "'Owner''s Data'!$A1:$A2",
      kind: 'transformed',
    });
    expect(
      transformXlsxStructuralSourceFormula('Data!A1:A3', 'Other', 'data', {
        count: 1,
        index: 2,
        kind: 'insert-rows',
      }),
    ).toEqual({ expression: 'Data!A1:A4', kind: 'transformed' });
    expect(
      transformXlsxStructuralSourceFormula('Data!AA1:AB3', 'Other', 'data', {
        count: 1,
        index: 27,
        kind: 'insert-columns',
      }),
    ).toEqual({ expression: 'Data!AB1:AC3', kind: 'transformed' });
    expect(
      transformXlsxStructuralSourceFormula('Data!B2', 'Other', 'data', {
        count: 1,
        index: 2,
        kind: 'insert-columns',
      }),
    ).toEqual({ expression: 'Data!C2', kind: 'transformed' });
  });

  it('preserves a valid reference to a different worksheet exactly', () => {
    expect(
      transformXlsxStructuralSourceFormula('Other!a1:a3', 'Data', 'Data', {
        count: 1,
        index: 1,
        kind: 'insert-rows',
      }),
    ).toEqual({ expression: 'Other!a1:a3', kind: 'preserved' });
  });

  it('classifies deleted and unchanged local ranges independently', () => {
    expect(
      transformXlsxStructuralSourceFormula('A2', 'Data', 'Data', {
        count: 1,
        index: 2,
        kind: 'delete-rows',
      }),
    ).toEqual({ kind: 'deleted' });
    expect(
      transformXlsxStructuralSourceFormula('A1', 'Data', 'Data', {
        count: 1,
        index: 3,
        kind: 'insert-rows',
      }),
    ).toEqual({ expression: 'A1', kind: 'preserved' });
  });

  it('accepts exact grid maxima and rejects structural overflow', () => {
    expect(
      transformXlsxStructuralSourceFormula('XFC1048575', 'Data', 'Data', {
        count: 1,
        index: 16_383,
        kind: 'insert-columns',
      }),
    ).toEqual({ expression: 'XFD1048575', kind: 'transformed' });
    expect(
      transformXlsxStructuralSourceFormula('XFD1', 'Data', 'Data', {
        count: 1,
        index: 16_384,
        kind: 'insert-columns',
      }),
    ).toEqual({ kind: 'unsupported' });
    expect(
      transformXlsxStructuralSourceFormula('A1048576', 'Data', 'Data', {
        count: 1,
        index: 1_048_576,
        kind: 'insert-rows',
      }),
    ).toEqual({ kind: 'unsupported' });
    expect(
      transformXlsxStructuralSourceFormula('A1048576', 'Data', 'Data', {
        count: 1,
        index: 2,
        kind: 'insert-columns',
      }),
    ).toEqual({ expression: 'A1048576', kind: 'preserved' });
  });

  it.each([
    'SUM(A1:A3)',
    '[Book.xlsx]Data!A1:A3',
    'Start:End!A1:A3',
    'Table1[Column]',
    'A3:A1',
    'C1:A1',
    'XFE1',
    'A1048577',
    'NamedRange',
    'A1,B2',
    'Data!A1junk',
    'Data!Other!A1',
    'Data Sheet!A1',
    '!A1',
    "'Data!A1",
    "'Data' A1",
    'A1:',
    ':A1',
    'A1:B2:C3',
  ])('rejects unsupported source syntax %s', (expression) => {
    expect(
      transformXlsxStructuralSourceFormula(expression, 'Data', 'Data', {
        count: 1,
        index: 1,
        kind: 'insert-rows',
      }),
    ).toEqual({ kind: 'unsupported' });
  });
});
