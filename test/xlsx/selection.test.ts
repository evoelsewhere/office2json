import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import type { XlsxSelection, XlsxSheet } from '../../src/formats/xlsx/types';
import {
  parseXlsxSelectionRange,
  resolveXlsxSelection,
  xlsxSelectionCandidateWindow,
  xlsxSelectionIncludesCell,
  xlsxSelectionIncludesRow,
} from '../../src/formats/xlsx/internal/selection';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from '../../src/formats/xlsx/internal/resource-limits';

const SHEETS: XlsxSheet[] = [
  {
    columns: [],
    comments: [],
    conditionalFormattings: [],
    dataValidations: [],
    drawings: [],
    hyperlinks: [],
    index: 0,
    kind: 'worksheet',
    mergedRanges: [],
    name: 'Data',
    payload: 'full-sheet',
    protectedRanges: [],
    rows: [],
    state: 'visible',
    tables: [],
    views: [],
  },
  {
    index: 1,
    kind: 'chart-sheet',
    name: 'Chart',
    payload: 'full-sheet',
    state: 'hidden',
  },
  {
    columns: [],
    comments: [],
    conditionalFormattings: [],
    dataValidations: [],
    drawings: [],
    hyperlinks: [],
    index: 2,
    kind: 'worksheet',
    mergedRanges: [],
    name: 'Résumé',
    payload: 'full-sheet',
    protectedRanges: [],
    rows: [],
    state: 'very-hidden',
    tables: [],
    views: [],
  },
];

function limits(
  overrides: Partial<ResolvedXlsxResourceLimits> = {},
): ResolvedXlsxResourceLimits {
  return { ...defaultXlsxResourceLimits(), ...overrides };
}

function resolve(
  selection: XlsxSelection | undefined,
  overrides: Partial<ResolvedXlsxResourceLimits> = {},
) {
  return resolveXlsxSelection(selection, SHEETS, limits(overrides));
}

function captureSelectionError(selection: unknown): XlsxParseError {
  try {
    resolveXlsxSelection(
      selection as XlsxSelection,
      SHEETS,
      defaultXlsxResourceLimits(),
    );
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected selection resolution to fail');
}

describe('XLSX selection', () => {
  it.each([
    [
      'A1',
      {
        end: { column: 1, row: 1 },
        reference: 'A1',
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
      'AA12',
      {
        end: { column: 27, row: 12 },
        reference: 'AA12',
        start: { column: 27, row: 12 },
      },
    ],
    [
      'A1:AA2',
      {
        end: { column: 27, row: 2 },
        reference: 'A1:AA2',
        start: { column: 1, row: 1 },
      },
    ],
    [
      'A:C',
      {
        end: { column: 3, row: 1_048_576 },
        reference: 'A:C',
        start: { column: 1, row: 1 },
      },
    ],
    [
      'A:A',
      {
        end: { column: 1, row: 1_048_576 },
        reference: 'A:A',
        start: { column: 1, row: 1 },
      },
    ],
    [
      'AA:AB',
      {
        end: { column: 28, row: 1_048_576 },
        reference: 'AA:AB',
        start: { column: 27, row: 1 },
      },
    ],
    [
      '$B:$XFD',
      {
        end: { column: 16_384, row: 1_048_576 },
        reference: 'B:XFD',
        start: { column: 2, row: 1 },
      },
    ],
    [
      '1:3',
      {
        end: { column: 16_384, row: 3 },
        reference: '1:3',
        start: { column: 1, row: 1 },
      },
    ],
    [
      '$2:$1048576',
      {
        end: { column: 16_384, row: 1_048_576 },
        reference: '2:1048576',
        start: { column: 1, row: 2 },
      },
    ],
  ] as const)('normalizes selection range %s', (source, expected) => {
    expect(parseXlsxSelectionRange(source)).toEqual(expected);
  });

  it.each([
    undefined,
    null,
    1,
    '',
    'A',
    '1',
    'A1:',
    ':B2',
    'A1:B2:C3',
    'C:A',
    '3:1',
    'XFE:XFE',
    '1048577:1048577',
    'Sheet1!A1',
    '[Book.xlsx]Sheet1!A1',
    'A1 B2',
    '!A:C',
    'A:C!',
    '!1:3',
    '1:3!',
    'A'.repeat(65),
  ])('rejects invalid selection range %#', (source) => {
    expect(parseXlsxSelectionRange(source)).toBeUndefined();
  });

  it('selects every sheet only when selection is absent', () => {
    expect(resolve(undefined)).toEqual([
      { kind: 'full-sheet' },
      { kind: 'full-sheet' },
      { kind: 'full-sheet' },
    ]);
    expect(resolve({})).toEqual([
      { kind: 'not-selected' },
      { kind: 'not-selected' },
      { kind: 'not-selected' },
    ]);
  });

  it('matches whole sheet names case-insensitively without changing order', () => {
    expect(resolve({ sheetNames: ['résumé', 'CHART'] })).toEqual([
      { kind: 'not-selected' },
      { kind: 'full-sheet' },
      { kind: 'full-sheet' },
    ]);
  });

  it('normalizes, sorts, and deduplicates worksheet ranges', () => {
    const result = resolve({
      ranges: {
        data: ['D2:D4', 'A1:B2', '$A$1:$B$2', '2:2'],
      },
    });

    expect(result[0]).toEqual({
      endRowPrefix: [2, 2, 4],
      kind: 'selected-ranges',
      ranges: [
        {
          end: { column: 2, row: 2 },
          reference: 'A1:B2',
          start: { column: 1, row: 1 },
        },
        {
          end: { column: 16_384, row: 2 },
          reference: '2:2',
          start: { column: 1, row: 2 },
        },
        {
          end: { column: 4, row: 4 },
          reference: 'D2:D4',
          start: { column: 4, row: 2 },
        },
      ],
    });
    expect(result.slice(1)).toEqual([
      { kind: 'not-selected' },
      { kind: 'not-selected' },
    ]);
  });

  it('sorts ranges by every coordinate deterministically', () => {
    const result = resolve({
      ranges: { Data: ['A2:A5', 'A1:B3', 'A1:A4', 'A1:A3'] },
    })[0];

    expect(result).toMatchObject({
      kind: 'selected-ranges',
      ranges: [
        { reference: 'A1:A3' },
        { reference: 'A1:B3' },
        { reference: 'A1:A4' },
        { reference: 'A2:A5' },
      ],
    });
  });

  it('lets a full-sheet selection win after validating ranges', () => {
    expect(
      resolve({
        ranges: { Data: ['A1:B2'] },
        sheetNames: ['DATA'],
      })[0],
    ).toEqual({ kind: 'full-sheet' });
  });

  it('indexes overlapping row and column membership deterministically', () => {
    const selection = resolve({
      ranges: { Data: ['A1:B2', 'D2:D4', 'B4:C5', '10:10'] },
    })[0]!;

    expect(xlsxSelectionIncludesCell(selection, 1, 1)).toBe(true);
    expect(xlsxSelectionIncludesCell(selection, 2, 2)).toBe(true);
    expect(xlsxSelectionIncludesCell(selection, 2, 3)).toBe(false);
    expect(xlsxSelectionIncludesCell(selection, 2, 4)).toBe(true);
    expect(xlsxSelectionIncludesCell(selection, 3, 4)).toBe(true);
    expect(xlsxSelectionIncludesCell(selection, 4, 2)).toBe(true);
    expect(xlsxSelectionIncludesCell(selection, 5, 3)).toBe(true);
    expect(xlsxSelectionIncludesCell(selection, 6, 3)).toBe(false);
    expect(xlsxSelectionIncludesCell(selection, 10, 16_384)).toBe(true);
    expect(xlsxSelectionCandidateWindow(selection, 0)).toBeNull();
    expect(xlsxSelectionCandidateWindow(selection, 1)).toEqual({
      first: 0,
      last: 0,
    });
    expect(xlsxSelectionCandidateWindow(selection, 3)).toEqual({
      first: 1,
      last: 1,
    });
    expect(xlsxSelectionCandidateWindow(selection, 5)).toEqual({
      first: 2,
      last: 2,
    });
    expect(xlsxSelectionCandidateWindow(selection, 6)).toBeNull();
    expect(xlsxSelectionCandidateWindow(selection, 10)).toEqual({
      first: 3,
      last: 3,
    });
    expect(xlsxSelectionIncludesRow(selection, 1)).toBe(true);
    expect(xlsxSelectionIncludesRow(selection, 3)).toBe(true);
    expect(xlsxSelectionIncludesRow(selection, 5)).toBe(true);
    expect(xlsxSelectionIncludesRow(selection, 6)).toBe(false);
    expect(xlsxSelectionIncludesRow(selection, 10)).toBe(true);
  });

  it('keeps an earlier long-lived range in the candidate window', () => {
    const selection = resolve({
      ranges: { Data: ['Z1:Z10', 'A2:A2'] },
    })[0]!;

    expect(xlsxSelectionCandidateWindow(selection, 3)).toEqual({
      first: 0,
      last: 1,
    });
    expect(xlsxSelectionIncludesCell(selection, 3, 1)).toBe(false);
    expect(xlsxSelectionIncludesCell(selection, 3, 26)).toBe(true);
  });

  it('finds a prefix lower bound after the middle of a larger index', () => {
    const selection = resolve({
      ranges: {
        Data: [
          'A1:A1',
          'A2:A2',
          'A3:A3',
          'A4:A4',
          'A5:A5',
          'Z6:Z100',
          'A7:A7',
          'A8:A8',
        ],
      },
    })[0]!;

    expect(xlsxSelectionCandidateWindow(selection, 50)).toEqual({
      first: 5,
      last: 7,
    });
    expect(xlsxSelectionIncludesCell(selection, 50, 26)).toBe(true);
  });

  it('finds the final prefix entry when every earlier range expired', () => {
    const selection = resolve({
      ranges: { Data: ['A1:A1', 'B1:B1', 'C1:C1', 'D1:D1', 'Z2:Z100'] },
    })[0]!;

    expect(xlsxSelectionCandidateWindow(selection, 50)).toEqual({
      first: 4,
      last: 4,
    });
    expect(xlsxSelectionIncludesCell(selection, 50, 26)).toBe(true);
  });

  it('handles full and absent membership without consulting an index', () => {
    expect(xlsxSelectionCandidateWindow({ kind: 'full-sheet' }, 1)).toBeNull();
    expect(
      xlsxSelectionCandidateWindow({ kind: 'not-selected' }, 1),
    ).toBeNull();
    expect(xlsxSelectionIncludesCell({ kind: 'full-sheet' }, 1, 1)).toBe(true);
    expect(xlsxSelectionIncludesRow({ kind: 'full-sheet' }, 1)).toBe(true);
    expect(xlsxSelectionIncludesCell({ kind: 'not-selected' }, 1, 1)).toBe(
      false,
    );
    expect(xlsxSelectionIncludesRow({ kind: 'not-selected' }, 1)).toBe(false);
  });

  it.each([
    [null, 'XLSX selection must be an object'],
    [[], 'XLSX selection must be an object'],
    [{ sheetNames: 'Data' }, 'XLSX selection sheetNames must be an array'],
    [{ sheetNames: [1] }, 'XLSX selection contains an invalid sheet name'],
    [{ sheetNames: [''] }, 'XLSX selection contains an invalid sheet name'],
    [
      { sheetNames: ['x'.repeat(31)] },
      'XLSX selection references an unknown sheet',
    ],
    [
      { sheetNames: ['Data', 'DATA'] },
      'XLSX selection contains duplicate sheet names',
    ],
    [{ sheetNames: [' Data'] }, 'XLSX selection references an unknown sheet'],
    [{ sheetNames: ['Missing'] }, 'XLSX selection references an unknown sheet'],
    [{ ranges: [] }, 'XLSX selection ranges must be a record of arrays'],
    [
      { ranges: { Data: [] } },
      'XLSX selection ranges must be non-empty arrays',
    ],
    [
      { ranges: { Data: 'A1' } },
      'XLSX selection ranges must be non-empty arrays',
    ],
    [
      { ranges: { '': ['A1'] } },
      'XLSX selection contains an invalid range sheet name',
    ],
    [
      { ranges: { ['x'.repeat(32)]: ['A1'] } },
      'XLSX selection contains an invalid range sheet name',
    ],
    [
      { ranges: { Data: ['A1'], data: ['B2'] } },
      'XLSX selection contains duplicate range sheet names',
    ],
    [
      { ranges: { Missing: ['A1'] } },
      'XLSX selection ranges reference an unknown sheet',
    ],
    [
      { ranges: { Chart: ['A1'] } },
      'XLSX selection ranges require a worksheet',
    ],
    [{ ranges: { Data: [1] } }, 'XLSX selection contains an invalid range'],
    [
      { ranges: { Data: ['Sheet1!A1'] } },
      'XLSX selection contains an invalid range',
    ],
  ] as const)('reports stable invalid selection %#', (selection, message) => {
    expect(captureSelectionError(selection).diagnostic).toEqual({
      code: 'invalid-selection',
      message,
      severity: 'error',
    });
  });

  it('accepts maxRangeAreas exactly and rejects every supplied area one over', () => {
    expect(
      resolve(
        { ranges: { Data: ['A1', 'A1'], Résumé: ['B2'] } },
        { maxRangeAreas: 3 },
      ),
    ).toMatchObject([
      { kind: 'selected-ranges', ranges: [{ reference: 'A1' }] },
      { kind: 'not-selected' },
      { kind: 'selected-ranges', ranges: [{ reference: 'B2' }] },
    ]);
    let error: unknown;
    try {
      resolve(
        { ranges: { Data: ['A1', 'A1'], Résumé: ['B2'] } },
        { maxRangeAreas: 2 },
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(XlsxResourceLimitError);
    expect(error).toMatchObject({
      actual: 3,
      limit: 2,
      limitName: 'maxRangeAreas',
      name: 'XlsxResourceLimitError',
    } satisfies Partial<XlsxResourceLimitError>);
  });
});
