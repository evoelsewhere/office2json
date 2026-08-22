import { describe, expect, it } from 'vitest';

import {
  xlsxBuiltinNumberFormatCode,
  xlsxNumberFormatDatePrecision,
} from '../../src/formats/xlsx/internal/number-format';

describe('XLSX number format metadata', () => {
  it.each([
    [0, 'General'],
    [1, '0'],
    [2, '0.00'],
    [3, '#,##0'],
    [4, '#,##0.00'],
    [9, '0%'],
    [10, '0.00%'],
    [11, '0.00E+00'],
    [12, '# ?/?'],
    [13, '# ??/??'],
    [14, 'mm-dd-yy'],
    [15, 'd-mmm-yy'],
    [16, 'd-mmm'],
    [17, 'mmm-yy'],
    [18, 'h:mm AM/PM'],
    [19, 'h:mm:ss AM/PM'],
    [20, 'h:mm'],
    [21, 'h:mm:ss'],
    [22, 'm/d/yy h:mm'],
    [37, '#,##0 ;(#,##0)'],
    [38, '#,##0 ;[Red](#,##0)'],
    [39, '#,##0.00;(#,##0.00)'],
    [40, '#,##0.00;[Red](#,##0.00)'],
    [45, 'mm:ss'],
    [46, '[h]:mm:ss'],
    [47, 'mmss.0'],
    [48, '##0.0E+0'],
    [49, '@'],
  ] as const)('resolves built-in format %s', (id, code) => {
    expect(xlsxBuiltinNumberFormatCode(id)).toBe(code);
  });

  it.each([-1, 1.5, 164, Number.NaN, Infinity])(
    'does not invent an unavailable built-in format for %s',
    (id) => {
      expect(xlsxBuiltinNumberFormatCode(id)).toBeUndefined();
    },
  );

  it.each([
    ['yyyy-mm-dd', 1, 'date'],
    ['mmm d, yyyy', 1, 'date'],
    ['m', 1, 'date'],
    ['y', 1, 'date'],
    ['d', 1, 'date'],
    ['h:mm', 0.5, 'time'],
    ['mm:ss.000', 0.5, 'time'],
    ['h:mm AM/PM', 0.5, 'time'],
    ['h:mm A/P', 0.5, 'time'],
    ['0 AM/PM', 0.5, 'time'],
    ['AM/PM 0', 0.5, 'time'],
    ['0 am/pm', 0.5, 'time'],
    ['0 A/P', 0.5, 'time'],
    ['A/P 0', 0.5, 'time'],
    ['yyyy-mm-dd"T"hh:mm:ss', 1.5, 'date-time'],
    ['[h]:mm:ss', 1.5, 'duration'],
    ['[hh]:mm:ss', 1.5, 'duration'],
    ['[Red][h]:mm:ss', 1.5, 'duration'],
    ['[mm]:ss', 1.5, 'duration'],
    ['[ss].000', 1.5, 'duration'],
    ['[$-409]mmm d, yyyy', 1, 'date'],
    ['[Red]yyyy-mm-dd', 1, 'date'],
  ] as const)('classifies %s as %s', (code, value, precision) => {
    expect(xlsxNumberFormatDatePrecision(code, value)).toBe(precision);
  });

  it.each([
    ['0.00', 1],
    ['General', 1],
    ['0.00 "days"', 1],
    ['0.00\\d', 1],
    ['0.00_d', 1],
    ['0.00*d', 1],
    ['[Blue]0.00', 1],
    ['[xh]0.00', 1],
    ['[hx]0.00', 1],
    ['[hm]0.00', 1],
  ] as const)('does not classify non-date format %s', (code, value) => {
    expect(xlsxNumberFormatDatePrecision(code, value)).toBeUndefined();
  });

  it('selects positive, negative, and zero numeric sections', () => {
    const code = 'yyyy-mm-dd;h:mm;[h]:mm';
    expect(xlsxNumberFormatDatePrecision(code, 1)).toBe('date');
    expect(xlsxNumberFormatDatePrecision(code, -1)).toBe('time');
    expect(xlsxNumberFormatDatePrecision(code, 0)).toBe('duration');

    const two = 'yyyy-mm-dd;h:mm';
    expect(xlsxNumberFormatDatePrecision(two, 1)).toBe('date');
    expect(xlsxNumberFormatDatePrecision(two, 0)).toBe('date');
    expect(xlsxNumberFormatDatePrecision(two, -1)).toBe('time');

    const four = 'yyyy-mm-dd;h:mm;[h]:mm;@';
    expect(xlsxNumberFormatDatePrecision(four, 1)).toBe('date');
  });

  it.each(['0\\;yyyy', '0_;yyyy', '0*;yyyy', '0";"yyyy', '[Foo;Bar]yyyy'])(
    'does not split an opaque semicolon in %s',
    (code) => {
      expect(xlsxNumberFormatDatePrecision(code, 1)).toBe('date');
    },
  );

  it.each([
    'yyyy-mm-dd;h:mm;0;@;extra',
    'yyyy-mm-dd;[>=1]h:mm',
    'yyyy-mm-dd;[oops',
    'yyyy-mm-dd;"oops',
    'yyyy-mm-dd;h:mm;0;@;',
    ']yyyy-mm-dd',
    'yyyy-mm-dd\\',
    'yyyy-mm-dd_',
    'yyyy-mm-dd*',
  ])('rejects ambiguous or malformed section grammar %#', (code) => {
    expect(xlsxNumberFormatDatePrecision(code, 1)).toBeUndefined();
  });

  it.each([Number.NaN, Infinity, -Infinity])(
    'does not classify a format for non-finite value %s',
    (value) => {
      expect(
        xlsxNumberFormatDatePrecision('yyyy-mm-dd', value),
      ).toBeUndefined();
    },
  );
});
