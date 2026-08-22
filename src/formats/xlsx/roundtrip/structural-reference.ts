import {
  parseXlsxCellReference,
  xlsxColumnName,
} from '../internal/cell-reference';
import { XLSX_MAX_COLUMNS, XLSX_MAX_ROWS } from '../internal/resource-limits';
import type {
  XlsxPageBreak,
  XlsxRange,
  XlsxDrawingMarker,
  XlsxWorksheetViewSelection,
} from '../types';

export interface XlsxStructuralReferenceOperation {
  count: number;
  index: number;
  kind: 'delete-columns' | 'delete-rows' | 'insert-columns' | 'insert-rows';
}

export interface XlsxStructuralDrawingAnchor {
  editAs?: 'absolute' | 'one-cell' | 'two-cell';
  from?: XlsxDrawingMarker;
  kind: 'absolute' | 'one-cell' | 'two-cell';
  to?: XlsxDrawingMarker;
}

function transformCoordinate(
  value: number,
  operation: XlsxStructuralReferenceOperation,
): number | null {
  const insertion = operation.kind.startsWith('insert-');
  if (insertion) {
    return value < operation.index ? value : value + operation.count;
  }
  const end = operation.index + operation.count - 1;
  if (value < operation.index) return value;
  if (value <= end) return null;
  return value - operation.count;
}

function transformInterval(
  start: number,
  end: number,
  operation: XlsxStructuralReferenceOperation,
): readonly [number, number] | null {
  if (operation.kind.startsWith('insert-')) {
    return [
      transformCoordinate(start, operation)!,
      transformCoordinate(end, operation)!,
    ];
  }
  const deletedEnd = operation.index + operation.count - 1;
  if (end < operation.index) return [start, end];
  if (start > deletedEnd) {
    return [start - operation.count, end - operation.count];
  }
  const outputStart = Math.min(start, operation.index);
  const deletedInside = Math.min(operation.count, end - operation.index + 1);
  const outputEnd = end - deletedInside;
  return outputStart > outputEnd ? null : [outputStart, outputEnd];
}

export function transformXlsxStructuralCell(
  row: number,
  column: number,
  operation: XlsxStructuralReferenceOperation,
): { address: string; column: number; row: number } | null {
  const rowOperation =
    operation.kind === 'delete-rows' || operation.kind === 'insert-rows';
  const columnOperation =
    operation.kind === 'delete-columns' || operation.kind === 'insert-columns';
  const transformedRow = rowOperation
    ? transformCoordinate(row, operation)
    : row;
  const transformedColumn = columnOperation
    ? transformCoordinate(column, operation)
    : column;
  if (transformedRow === null || transformedColumn === null) return null;
  return {
    address: `${xlsxColumnName(transformedColumn)!}${transformedRow}`,
    column: transformedColumn,
    row: transformedRow,
  };
}

export function transformXlsxStructuralRange(
  range: XlsxRange,
  operation: XlsxStructuralReferenceOperation,
): XlsxRange | null {
  const rowOperation =
    operation.kind === 'delete-rows' || operation.kind === 'insert-rows';
  const columnOperation =
    operation.kind === 'delete-columns' || operation.kind === 'insert-columns';
  const rows = rowOperation
    ? transformInterval(range.start.row, range.end.row, operation)
    : ([range.start.row, range.end.row] as const);
  const columns = columnOperation
    ? transformInterval(range.start.column, range.end.column, operation)
    : ([range.start.column, range.end.column] as const);
  if (rows === null || columns === null) return null;
  const startAddress = `${xlsxColumnName(columns[0])!}${rows[0]}`;
  const endAddress = `${xlsxColumnName(columns[1])!}${rows[1]}`;
  return {
    end: { column: columns[1], row: rows[1] },
    reference:
      startAddress === endAddress
        ? startAddress
        : `${startAddress}:${endAddress}`,
    start: { column: columns[0], row: rows[0] },
  };
}

export function transformXlsxStructuralPageBreak(
  pageBreak: XlsxPageBreak,
  axis: 'column' | 'row',
  operation: XlsxStructuralReferenceOperation,
): XlsxPageBreak | null {
  const positionOperation =
    axis === 'row'
      ? operation.kind.endsWith('-rows')
      : operation.kind.endsWith('-columns');
  const positionLimit = axis === 'row' ? XLSX_MAX_ROWS : XLSX_MAX_COLUMNS;
  let position = pageBreak.position;
  if (positionOperation) {
    const transformed = transformCoordinate(position, operation);
    if (transformed === null) return null;
    position = Math.min(transformed, positionLimit);
  }

  const extentOperation =
    axis === 'row'
      ? operation.kind.endsWith('-columns')
      : operation.kind.endsWith('-rows');
  const extentLimit = axis === 'row' ? XLSX_MAX_COLUMNS : XLSX_MAX_ROWS;
  let start = pageBreak.start;
  let end = pageBreak.end;
  const fullExtent = start === 0 && end === extentLimit - 1;
  if (extentOperation && !fullExtent) {
    const transformed = transformInterval(start, end, {
      ...operation,
      index: operation.index - 1,
    });
    if (transformed === null || transformed[0] >= extentLimit) return null;
    start = transformed[0];
    end = Math.min(transformed[1], extentLimit - 1);
  }
  return { ...pageBreak, end, position, start };
}

function rangeStartAddress(range: XlsxRange): string {
  return `${xlsxColumnName(range.start.column)!}${range.start.row}`;
}

export function transformXlsxStructuralVisualCell(
  address: string,
  operation: XlsxStructuralReferenceOperation,
): string {
  const parsed = parseXlsxCellReference(address)!;
  const transformed = transformXlsxStructuralCell(
    parsed.row,
    parsed.column,
    operation,
  );
  if (transformed !== null) return transformed.address;
  const row = operation.kind.endsWith('-rows') ? operation.index : parsed.row;
  const column = operation.kind.endsWith('-columns')
    ? operation.index
    : parsed.column;
  return `${xlsxColumnName(column)!}${row}`;
}

export function transformXlsxStructuralViewSelection(
  selection: XlsxWorksheetViewSelection,
  operation: XlsxStructuralReferenceOperation,
): XlsxWorksheetViewSelection | null {
  const transformedRanges = selection.ranges.flatMap((range, sourceIndex) => {
    const transformed = transformXlsxStructuralRange(range, operation);
    return transformed === null ? [] : [{ range: transformed, sourceIndex }];
  });
  if (transformedRanges.length === 0) return null;
  const sourceActiveIndex = selection.activeCellId ?? 0;
  const activeIndex = Math.max(
    0,
    transformedRanges.findIndex(
      (entry) => entry.sourceIndex === sourceActiveIndex,
    ),
  );
  let activeCell: string | undefined;
  if (selection.activeCell !== undefined) {
    const parsedActiveCell = parseXlsxCellReference(selection.activeCell)!;
    activeCell = transformXlsxStructuralCell(
      parsedActiveCell.row,
      parsedActiveCell.column,
      operation,
    )?.address;
  }
  if (selection.activeCell !== undefined && activeCell === undefined) {
    activeCell = rangeStartAddress(transformedRanges[activeIndex]!.range);
  }
  return {
    ...(activeCell === undefined ? {} : { activeCell }),
    ...(selection.activeCellId === undefined
      ? {}
      : { activeCellId: activeIndex }),
    pane: selection.pane,
    ranges: transformedRanges.map((entry) => entry.range),
  };
}

function visualMarker(
  marker: XlsxDrawingMarker,
  operation: XlsxStructuralReferenceOperation,
): XlsxDrawingMarker | null {
  const transformed = transformXlsxStructuralCell(
    marker.row,
    marker.column,
    operation,
  );
  const row =
    transformed?.row ??
    (operation.kind.endsWith('-rows') ? operation.index : marker.row);
  const column =
    transformed?.column ??
    (operation.kind.endsWith('-columns') ? operation.index : marker.column);
  if (row > XLSX_MAX_ROWS || column > XLSX_MAX_COLUMNS) {
    return null;
  }
  return { ...marker, column, row };
}

export function transformXlsxStructuralDrawingAnchor(
  anchor: XlsxStructuralDrawingAnchor,
  operation: XlsxStructuralReferenceOperation,
): XlsxStructuralDrawingAnchor | null {
  if (anchor.kind === 'absolute' || anchor.editAs === 'absolute') return anchor;
  const from = anchor.from!;
  if (anchor.kind === 'one-cell') {
    const transformedFrom = visualMarker(from, operation);
    return transformedFrom === null
      ? null
      : { ...anchor, from: transformedFrom };
  }
  const to = anchor.to!;
  if (anchor.editAs === 'one-cell') {
    const transformedFrom = visualMarker(from, operation);
    if (transformedFrom === null) return null;
    const rowDelta = transformedFrom.row - from.row;
    const columnDelta = transformedFrom.column - from.column;
    const transformedTo = {
      ...to,
      column: to.column + columnDelta,
      row: to.row + rowDelta,
    };
    if (
      transformedTo.column > XLSX_MAX_COLUMNS ||
      transformedTo.row > XLSX_MAX_ROWS
    ) {
      return null;
    }
    return { ...anchor, from: transformedFrom, to: transformedTo };
  }
  const rowOperation = operation.kind.endsWith('-rows');
  const columnOperation = operation.kind.endsWith('-columns');
  const rows = rowOperation
    ? transformInterval(from.row, to.row, operation)
    : ([from.row, to.row] as const);
  const columns = columnOperation
    ? transformInterval(from.column, to.column, operation)
    : ([from.column, to.column] as const);
  if (
    rows === null ||
    columns === null ||
    rows[1] > XLSX_MAX_ROWS ||
    columns[1] > XLSX_MAX_COLUMNS
  ) {
    return null;
  }
  return {
    ...anchor,
    from: {
      ...from,
      column: columns[0],
      row: rows[0],
    },
    to: { ...to, column: columns[1], row: rows[1] },
  };
}
