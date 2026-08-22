import { describe, expect, it } from 'vitest';

import {
  transformXlsxStructuralCell,
  transformXlsxStructuralDrawingAnchor,
  transformXlsxStructuralPageBreak,
  transformXlsxStructuralRange,
  transformXlsxStructuralViewSelection,
  transformXlsxStructuralVisualCell,
  type XlsxStructuralReferenceOperation,
} from '../../src/formats/xlsx/roundtrip/structural-reference';
import type { XlsxRange } from '../../src/formats/xlsx/types';

function range(reference: string, start: number, end: number): XlsxRange {
  return {
    end: { column: 1, row: end },
    reference,
    start: { column: 1, row: start },
  };
}

describe('XLSX structural reference transforms', () => {
  it.each([
    [1, 'A1'],
    [2, 'A4'],
    [3, 'A5'],
    [4, 'A6'],
  ] as const)('inserts rows at coordinate %s', (row, address) => {
    expect(
      transformXlsxStructuralCell(row, 1, {
        count: 2,
        index: 2,
        kind: 'insert-rows',
      }),
    ).toEqual({ address, column: 1, row: Number(address.slice(1)) });
  });

  it.each([
    [1, 'A1'],
    [2, null],
    [3, null],
    [4, 'A2'],
    [5, 'A3'],
  ] as const)('deletes rows at coordinate %s', (row, address) => {
    expect(
      transformXlsxStructuralCell(row, 1, {
        count: 2,
        index: 2,
        kind: 'delete-rows',
      }),
    ).toEqual(
      address === null
        ? null
        : { address, column: 1, row: Number(address.slice(1)) },
    );
  });

  it.each([
    [1, 1, 'A1'],
    [2, 2, null],
    [1, 2, 'A1'],
    [2, 3, null],
    [1, 4, 'A1:A2'],
    [4, 5, 'A2:A3'],
    [3, 4, 'A2'],
    [3, 5, 'A2:A3'],
    [5, 6, 'A3:A4'],
  ] as const)(
    'clips deleted row interval %s:%s',
    (start, end, expectedReference) => {
      expect(
        transformXlsxStructuralRange(range(`A${start}:A${end}`, start, end), {
          count: 2,
          index: 2,
          kind: 'delete-rows',
        })?.reference ?? null,
      ).toBe(expectedReference);
    },
  );

  it('expands spanning ranges and transforms columns symmetrically', () => {
    expect(
      transformXlsxStructuralCell(2, 3, {
        count: 2,
        index: 2,
        kind: 'insert-rows',
      }),
    ).toEqual({ address: 'C4', column: 3, row: 4 });
    expect(
      transformXlsxStructuralRange(range('A1:A2', 1, 2), {
        count: 2,
        index: 4,
        kind: 'delete-rows',
      })?.reference,
    ).toBe('A1:A2');
    expect(
      transformXlsxStructuralRange(range('A1:A1', 1, 1), {
        count: 2,
        index: 3,
        kind: 'insert-rows',
      })?.reference,
    ).toBe('A1');
    expect(
      transformXlsxStructuralRange(range('A1:A4', 1, 4), {
        count: 2,
        index: 2,
        kind: 'insert-rows',
      }),
    ).toEqual({
      end: { column: 1, row: 6 },
      reference: 'A1:A6',
      start: { column: 1, row: 1 },
    });
    const source: XlsxRange = {
      end: { column: 5, row: 2 },
      reference: 'B2:E2',
      start: { column: 2, row: 2 },
    };
    const operations: XlsxStructuralReferenceOperation[] = [
      { count: 2, index: 3, kind: 'insert-columns' },
      { count: 2, index: 2, kind: 'delete-columns' },
    ];
    const inserted = transformXlsxStructuralRange(source, operations[0]!)!;
    expect(inserted.reference).toBe('B2:G2');
    expect(transformXlsxStructuralRange(inserted, operations[1]!)).toEqual({
      end: { column: 5, row: 2 },
      reference: 'B2:E2',
      start: { column: 2, row: 2 },
    });
    expect(
      transformXlsxStructuralCell(2, 3, {
        count: 1,
        index: 2,
        kind: 'insert-columns',
      }),
    ).toEqual({ address: 'D2', column: 4, row: 2 });
    expect(
      transformXlsxStructuralCell(2, 2, {
        count: 1,
        index: 2,
        kind: 'delete-columns',
      }),
    ).toBeNull();
  });

  it('transforms page-break positions and zero-based extents by axis', () => {
    const source = {
      end: 2,
      manual: true,
      pivot: false,
      position: 3,
      start: 0,
    };
    expect(
      transformXlsxStructuralPageBreak(source, 'row', {
        count: 2,
        index: 2,
        kind: 'insert-rows',
      }),
    ).toEqual({ ...source, position: 5 });
    expect(
      transformXlsxStructuralPageBreak(source, 'row', {
        count: 1,
        index: 3,
        kind: 'delete-rows',
      }),
    ).toBeNull();
    expect(
      transformXlsxStructuralPageBreak(source, 'row', {
        count: 2,
        index: 2,
        kind: 'insert-columns',
      }),
    ).toEqual({ ...source, end: 4 });
    expect(
      transformXlsxStructuralPageBreak(source, 'row', {
        count: 2,
        index: 2,
        kind: 'delete-columns',
      }),
    ).toEqual({ ...source, end: 0 });
    expect(
      transformXlsxStructuralPageBreak({ ...source, end: 2, start: 2 }, 'row', {
        count: 1,
        index: 3,
        kind: 'delete-columns',
      }),
    ).toBeNull();
    expect(
      transformXlsxStructuralPageBreak(source, 'column', {
        count: 1,
        index: 2,
        kind: 'insert-columns',
      }),
    ).toEqual({ ...source, position: 4 });
  });

  it('keeps full-grid page-break extents and terminal positions bounded', () => {
    const fullRowBreak = {
      end: 16_383,
      manual: false,
      pivot: true,
      position: 1_048_576,
      start: 0,
    };
    expect(
      transformXlsxStructuralPageBreak(fullRowBreak, 'row', {
        count: 1,
        index: 1,
        kind: 'insert-columns',
      }),
    ).toEqual(fullRowBreak);
    expect(
      transformXlsxStructuralPageBreak(fullRowBreak, 'row', {
        count: 1,
        index: 1,
        kind: 'insert-rows',
      }),
    ).toEqual(fullRowBreak);
    expect(
      transformXlsxStructuralPageBreak(
        { ...fullRowBreak, end: 16_382, start: 16_381 },
        'row',
        { count: 2, index: 16_383, kind: 'insert-columns' },
      ),
    ).toEqual({ ...fullRowBreak, end: 16_383, start: 16_381 });
    expect(
      transformXlsxStructuralPageBreak(
        { ...fullRowBreak, end: 16_383, start: 16_383 },
        'row',
        { count: 1, index: 16_384, kind: 'insert-columns' },
      ),
    ).toBeNull();
  });

  it('transforms visual anchors and remaps worksheet selections', () => {
    expect(
      transformXlsxStructuralVisualCell('B2', {
        count: 1,
        index: 2,
        kind: 'insert-rows',
      }),
    ).toBe('B3');
    expect(
      transformXlsxStructuralVisualCell('B2', {
        count: 1,
        index: 2,
        kind: 'delete-rows',
      }),
    ).toBe('B2');
    expect(
      transformXlsxStructuralVisualCell('B3', {
        count: 2,
        index: 2,
        kind: 'delete-rows',
      }),
    ).toBe('B2');
    expect(
      transformXlsxStructuralVisualCell('C5', {
        count: 2,
        index: 2,
        kind: 'delete-columns',
      }),
    ).toBe('B5');
    const selection = {
      activeCell: 'B3',
      activeCellId: 1,
      pane: 'top-left' as const,
      ranges: [
        range('A1', 1, 1),
        {
          end: { column: 2, row: 4 },
          reference: 'B2:B4',
          start: { column: 2, row: 2 },
        },
      ],
    };
    expect(
      transformXlsxStructuralViewSelection(selection, {
        count: 1,
        index: 1,
        kind: 'delete-rows',
      }),
    ).toEqual({
      activeCell: 'B2',
      activeCellId: 0,
      pane: 'top-left',
      ranges: [
        {
          end: { column: 2, row: 3 },
          reference: 'B1:B3',
          start: { column: 2, row: 1 },
        },
      ],
    });
    expect(
      transformXlsxStructuralViewSelection(
        { ...selection, activeCell: 'B2', ranges: [selection.ranges[1]!] },
        { count: 1, index: 2, kind: 'delete-rows' },
      ),
    ).toMatchObject({ activeCell: 'B2', activeCellId: 0 });
    expect(
      transformXlsxStructuralViewSelection(
        { ...selection, ranges: [selection.ranges[0]!] },
        { count: 1, index: 1, kind: 'delete-rows' },
      ),
    ).toBeNull();
    expect(
      transformXlsxStructuralViewSelection(
        {
          pane: 'top-left',
          ranges: [
            {
              end: { column: 1, row: 2 },
              reference: 'A1:A2',
              start: { column: 1, row: 1 },
            },
          ],
        },
        { count: 1, index: 3, kind: 'insert-rows' },
      ),
    ).toEqual({
      pane: 'top-left',
      ranges: [
        {
          end: { column: 1, row: 2 },
          reference: 'A1:A2',
          start: { column: 1, row: 1 },
        },
      ],
    });
  });

  it('transforms drawing anchors according to their edit mode', () => {
    const marker = {
      column: 2,
      columnOffset: 10,
      row: 2,
      rowOffset: 20,
    };
    expect(
      transformXlsxStructuralDrawingAnchor(
        { from: marker, kind: 'one-cell' },
        { count: 1, index: 2, kind: 'insert-rows' },
      ),
    ).toEqual({ from: { ...marker, row: 3 }, kind: 'one-cell' });
    expect(
      transformXlsxStructuralDrawingAnchor(
        { from: marker, kind: 'one-cell' },
        { count: 1, index: 2, kind: 'delete-columns' },
      ),
    ).toEqual({ from: marker, kind: 'one-cell' });
    expect(
      transformXlsxStructuralDrawingAnchor(
        { from: { ...marker, column: 5, row: 3 }, kind: 'one-cell' },
        { count: 2, index: 2, kind: 'delete-rows' },
      ),
    ).toMatchObject({ from: { column: 5, row: 2 } });
    expect(
      transformXlsxStructuralDrawingAnchor(
        { from: { ...marker, column: 3, row: 5 }, kind: 'one-cell' },
        { count: 2, index: 2, kind: 'delete-columns' },
      ),
    ).toMatchObject({ from: { column: 2, row: 5 } });
    const twoCell = {
      from: marker,
      kind: 'two-cell' as const,
      to: { ...marker, column: 4, row: 4 },
    };
    expect(
      transformXlsxStructuralDrawingAnchor(twoCell, {
        count: 1,
        index: 3,
        kind: 'insert-rows',
      }),
    ).toMatchObject({ from: marker, to: { column: 4, row: 5 } });
    expect(
      transformXlsxStructuralDrawingAnchor(twoCell, {
        count: 2,
        index: 2,
        kind: 'delete-rows',
      }),
    ).toMatchObject({ from: { row: 2 }, to: { row: 2 } });
    expect(
      transformXlsxStructuralDrawingAnchor(
        { ...twoCell, editAs: 'one-cell' },
        { count: 1, index: 3, kind: 'insert-rows' },
      ),
    ).toEqual({ ...twoCell, editAs: 'one-cell' });
    expect(
      transformXlsxStructuralDrawingAnchor(
        { ...twoCell, editAs: 'one-cell' },
        { count: 1, index: 2, kind: 'insert-columns' },
      ),
    ).toMatchObject({
      from: { column: 3 },
      to: { column: 5 },
    });
    const absolute = { ...twoCell, editAs: 'absolute' as const };
    expect(
      transformXlsxStructuralDrawingAnchor(absolute, {
        count: 1,
        index: 1,
        kind: 'insert-rows',
      }),
    ).toBe(absolute);
    const absoluteAnchor = { kind: 'absolute' as const };
    expect(
      transformXlsxStructuralDrawingAnchor(absoluteAnchor, {
        count: 1,
        index: 1,
        kind: 'insert-rows',
      }),
    ).toBe(absoluteAnchor);
    expect(
      transformXlsxStructuralDrawingAnchor(
        {
          ...twoCell,
          editAs: 'one-cell',
          from: { ...marker, column: 16_384 },
          to: { ...marker, column: 16_384 },
        },
        { count: 1, index: 16_384, kind: 'insert-columns' },
      ),
    ).toBeNull();
    expect(
      transformXlsxStructuralDrawingAnchor(
        {
          from: { ...marker, column: 16_384 },
          kind: 'one-cell',
        },
        { count: 1, index: 16_384, kind: 'insert-columns' },
      ),
    ).toBeNull();
    expect(
      transformXlsxStructuralDrawingAnchor(
        {
          from: { ...marker, row: 1_048_576 },
          kind: 'one-cell',
        },
        { count: 1, index: 1_048_576, kind: 'insert-rows' },
      ),
    ).toBeNull();
    expect(
      transformXlsxStructuralDrawingAnchor(
        {
          from: { ...marker, row: 1_048_576 },
          kind: 'one-cell',
        },
        { count: 1, index: 1, kind: 'insert-columns' },
      ),
    ).toMatchObject({ from: { row: 1_048_576 } });
    expect(
      transformXlsxStructuralDrawingAnchor(
        {
          from: { ...marker, column: 16_384 },
          kind: 'one-cell',
        },
        { count: 1, index: 1, kind: 'insert-rows' },
      ),
    ).toMatchObject({ from: { column: 16_384 } });
    expect(
      transformXlsxStructuralDrawingAnchor(
        {
          ...twoCell,
          editAs: 'one-cell',
          from: { ...marker, column: 16_383 },
          to: { ...marker, column: 16_384 },
        },
        { count: 1, index: 16_383, kind: 'insert-columns' },
      ),
    ).toBeNull();
    expect(
      transformXlsxStructuralDrawingAnchor(
        {
          ...twoCell,
          from: { ...marker, row: 1_048_575 },
          to: { ...marker, row: 1_048_576 },
        },
        { count: 1, index: 1_048_575, kind: 'insert-rows' },
      ),
    ).toBeNull();
    expect(
      transformXlsxStructuralDrawingAnchor(
        {
          ...twoCell,
          from: { ...marker, column: 16_383 },
          to: { ...marker, column: 16_384 },
        },
        { count: 1, index: 16_383, kind: 'insert-columns' },
      ),
    ).toBeNull();
    expect(
      transformXlsxStructuralDrawingAnchor(
        {
          ...twoCell,
          from: { ...marker, row: 1_048_575 },
          to: { ...marker, row: 1_048_575 },
        },
        { count: 1, index: 1_048_575, kind: 'insert-rows' },
      ),
    ).toMatchObject({ to: { row: 1_048_576 } });
    expect(
      transformXlsxStructuralDrawingAnchor(
        {
          ...twoCell,
          from: { ...marker, column: 16_383 },
          to: { ...marker, column: 16_383 },
        },
        { count: 1, index: 16_383, kind: 'insert-columns' },
      ),
    ).toMatchObject({ to: { column: 16_384 } });
    expect(
      transformXlsxStructuralDrawingAnchor(
        {
          ...twoCell,
          editAs: 'one-cell',
          from: { ...marker, row: 1_048_575 },
          to: { ...marker, row: 1_048_576 },
        },
        { count: 1, index: 1_048_575, kind: 'insert-rows' },
      ),
    ).toBeNull();
    expect(
      transformXlsxStructuralDrawingAnchor(
        {
          ...twoCell,
          editAs: 'one-cell',
          from: { ...marker, column: 16_383, row: 1_048_576 },
          to: { ...marker, column: 16_383, row: 1_048_576 },
        },
        { count: 1, index: 16_383, kind: 'insert-columns' },
      ),
    ).toMatchObject({ to: { column: 16_384, row: 1_048_576 } });
  });
});
