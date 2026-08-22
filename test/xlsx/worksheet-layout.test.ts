import { describe, expect, it } from 'vitest';

import {
  normalizeXlsxColumnRanges,
  type XlsxAuthoredColumnRange,
  xlsxMergedRangesOverlap,
} from '../../src/formats/xlsx/internal/worksheet-layout';
import type { XlsxRange } from '../../src/formats/xlsx/types';

function merged(
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number,
): XlsxRange {
  return {
    end: { column: endColumn, row: endRow },
    reference: `${startRow},${startColumn}:${endRow},${endColumn}`,
    start: { column: startColumn, row: startRow },
  };
}

function bruteOverlap(ranges: readonly XlsxRange[]): boolean {
  for (let left = 0; left < ranges.length; left += 1) {
    for (let right = left + 1; right < ranges.length; right += 1) {
      const a = ranges[left]!;
      const b = ranges[right]!;
      if (
        a.start.row <= b.end.row &&
        a.end.row >= b.start.row &&
        a.start.column <= b.end.column &&
        a.end.column >= b.start.column
      ) {
        return true;
      }
    }
  }
  return false;
}

function bruteColumns(
  values: readonly XlsxAuthoredColumnRange[],
): Array<Record<string, unknown> | undefined> {
  const output: Array<Record<string, unknown> | undefined> = [];
  for (const value of values) {
    for (let column = value.start; column <= value.end; column += 1) {
      output[column] = {
        ...(value.collapsed === undefined
          ? {}
          : { collapsed: value.collapsed }),
        ...(value.hidden === undefined ? {} : { hidden: value.hidden }),
        ...(value.outlineLevel === undefined
          ? {}
          : { outlineLevel: value.outlineLevel }),
        ...(value.style === undefined ? {} : { style: value.style }),
        ...(value.width === undefined ? {} : { width: value.width }),
      };
    }
  }
  return output;
}

function expandColumns(
  values: ReturnType<typeof normalizeXlsxColumnRanges>,
): Array<Record<string, unknown> | undefined> {
  const output: Array<Record<string, unknown> | undefined> = [];
  for (const value of values) {
    const properties = {
      ...(value.collapsed === undefined ? {} : { collapsed: value.collapsed }),
      ...(value.hidden === undefined ? {} : { hidden: value.hidden }),
      ...(value.outlineLevel === undefined
        ? {}
        : { outlineLevel: value.outlineLevel }),
      ...(value.style === undefined ? {} : { style: value.style }),
      ...(value.width === undefined ? {} : { width: value.width }),
    };
    for (let column = value.start; column <= value.end; column += 1) {
      output[column] = properties;
    }
  }
  return output;
}

describe('XLSX worksheet layout normalization', () => {
  it('applies last-authored column precedence and coalesces equal neighbors', () => {
    const result = normalizeXlsxColumnRanges([
      { end: 6, hidden: true, order: 0, start: 1, width: 10 },
      { end: 4, order: 1, start: 3, style: 2, width: 20 },
      { collapsed: true, end: 5, order: 2, outlineLevel: 3, start: 4 },
      { end: 8, hidden: true, order: 3, start: 7, width: 10 },
    ]);

    expect(result).toStrictEqual([
      { end: 2, hidden: true, start: 1, width: 10 },
      { end: 3, start: 3, style: 2, width: 20 },
      { collapsed: true, end: 5, outlineLevel: 3, start: 4 },
      { end: 8, hidden: true, start: 6, width: 10 },
    ]);
  });

  it('returns empty and disjoint column ranges without phantom gaps', () => {
    expect(normalizeXlsxColumnRanges([])).toStrictEqual([]);
    expect(
      normalizeXlsxColumnRanges([
        { end: 1, order: 0, start: 1, width: 1 },
        { end: 4, order: 1, start: 4, width: 1 },
      ]),
    ).toStrictEqual([
      { end: 1, start: 1, width: 1 },
      { end: 4, start: 4, width: 1 },
    ]);
  });

  it('preserves false and zero column properties without undefined keys', () => {
    expect(
      normalizeXlsxColumnRanges([
        {
          collapsed: false,
          end: 2,
          hidden: false,
          order: 0,
          outlineLevel: 0,
          start: 1,
          style: 0,
          width: 0,
        },
      ]),
    ).toStrictEqual([
      {
        collapsed: false,
        end: 2,
        hidden: false,
        outlineLevel: 0,
        start: 1,
        style: 0,
        width: 0,
      },
    ]);
  });

  it.each([
    [{ collapsed: true }, { collapsed: false }],
    [{ hidden: true }, { hidden: false }],
    [{ outlineLevel: 1 }, { outlineLevel: 2 }],
    [{ style: 0 }, { style: 1 }],
    [{ width: 1 }, { width: 2 }],
  ] as const)(
    'does not coalesce columns differing by property %#',
    (left, right) => {
      expect(
        normalizeXlsxColumnRanges([
          { end: 1, order: 0, start: 1, ...left },
          { end: 2, order: 1, start: 2, ...right },
        ]),
      ).toHaveLength(2);
    },
  );

  it('uses range end as a deterministic heap tie-breaker', () => {
    expect(
      normalizeXlsxColumnRanges([
        { end: 2, order: 1, start: 1, width: 1 },
        { end: 3, order: 1, start: 1, width: 2 },
      ]),
    ).toStrictEqual([{ end: 3, start: 1, width: 2 }]);
  });

  it('matches a brute-force column oracle across generated overlays', () => {
    let state = 0x12345678;
    const random = (maximum: number) => {
      state = (state * 1_664_525 + 1_013_904_223) >>> 0;
      return state % maximum;
    };
    for (let sample = 0; sample < 128; sample += 1) {
      const values: XlsxAuthoredColumnRange[] = [];
      const count = 1 + random(12);
      for (let order = 0; order < count; order += 1) {
        const start = 1 + random(12);
        const end = start + random(13 - start);
        values.push({
          end,
          ...(random(2) === 0 ? {} : { hidden: random(2) === 1 }),
          order,
          start,
          ...(random(2) === 0 ? {} : { style: random(4) }),
          ...(random(2) === 0 ? {} : { width: random(20) / 2 }),
        });
      }
      expect(expandColumns(normalizeXlsxColumnRanges(values))).toStrictEqual(
        bruteColumns(values),
      );
    }
  });

  it.each([
    [[], false],
    [[merged(1, 1, 2, 2)], false],
    [[merged(1, 1, 2, 2), merged(3, 1, 4, 2)], false],
    [[merged(3, 1, 4, 2), merged(1, 1, 2, 2)], false],
    [[merged(1, 1, 2, 2), merged(1, 3, 2, 4)], false],
    [[merged(1, 1, 2, 2), merged(2, 2, 3, 3)], true],
    [[merged(1, 1, 4, 1), merged(2, 2, 3, 2)], false],
    [[merged(1, 1, 4, 3), merged(2, 2, 3, 2)], true],
  ] as const)('detects merged-range overlap %#', (ranges, expected) => {
    expect(xlsxMergedRangesOverlap(ranges)).toBe(expected);
  });

  it('matches a brute-force merged-range oracle across generated rectangles', () => {
    let state = 0x9e3779b9;
    const random = (maximum: number) => {
      state = (state * 1_103_515_245 + 12_345) >>> 0;
      return state % maximum;
    };
    for (let sample = 0; sample < 256; sample += 1) {
      const ranges: XlsxRange[] = [];
      const count = random(10);
      for (let index = 0; index < count; index += 1) {
        const startRow = 1 + random(8);
        const startColumn = 1 + random(8);
        ranges.push(
          merged(
            startRow,
            startColumn,
            startRow + random(9 - startRow),
            startColumn + random(9 - startColumn),
          ),
        );
      }
      expect(xlsxMergedRangesOverlap(ranges)).toBe(bruteOverlap(ranges));
    }
  });

  it.each([
    [[merged(1, 31, 2, 34), merged(1, 32, 2, 32)], true],
    [[merged(1, 31, 2, 32), merged(1, 33, 2, 34)], false],
    [[merged(1, 1, 2, 64), merged(2, 64, 3, 64)], true],
    [[merged(1, 1, 2, 64), merged(3, 1, 4, 64)], false],
    [[merged(3, 1, 4, 64), merged(1, 1, 2, 64)], false],
    [[merged(1, 32, 2, 33), merged(1, 31, 2, 31)], false],
    [[merged(1, 31, 2, 64), merged(1, 33, 2, 33)], true],
    [[merged(1, 33, 2, 64), merged(1, 65, 2, 65)], false],
  ] as const)(
    'handles merged ranges across bitset words %#',
    (ranges, expected) => {
      expect(xlsxMergedRangesOverlap(ranges)).toBe(expected);
    },
  );

  it('matches a brute-force oracle for every rectangle pair on a 4x4 grid', () => {
    const rectangles: XlsxRange[] = [];
    for (let startRow = 1; startRow <= 4; startRow += 1) {
      for (let startColumn = 1; startColumn <= 4; startColumn += 1) {
        for (let endRow = startRow; endRow <= 4; endRow += 1) {
          for (let endColumn = startColumn; endColumn <= 4; endColumn += 1) {
            rectangles.push(merged(startRow, startColumn, endRow, endColumn));
          }
        }
      }
    }
    for (let left = 0; left < rectangles.length; left += 1) {
      for (let right = left + 1; right < rectangles.length; right += 1) {
        const pair = [rectangles[left]!, rectangles[right]!];
        expect(xlsxMergedRangesOverlap(pair)).toBe(bruteOverlap(pair));
      }
    }
  });
});
