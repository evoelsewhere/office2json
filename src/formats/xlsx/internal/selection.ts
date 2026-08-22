import { XlsxParseError } from '../errors';
import type { XlsxRange, XlsxSelection, XlsxSheet } from '../types';
import {
  parseXlsxCellReference,
  parseXlsxRangeReference,
  xlsxColumnName,
} from './cell-reference';
import {
  type ResolvedXlsxResourceLimits,
  XLSX_MAX_COLUMNS,
  XLSX_MAX_ROWS,
  XlsxResourceLimitError,
} from './resource-limits';

export type XlsxResolvedSheetSelection =
  | { kind: 'full-sheet' }
  | { kind: 'not-selected' }
  | {
      endRowPrefix: readonly number[];
      kind: 'selected-ranges';
      ranges: readonly XlsxRange[];
    };

function selectionFailure(message: string): never {
  throw new XlsxParseError({
    code: 'invalid-selection',
    message,
    severity: 'error',
  });
}

function foldedSheetName(value: unknown): string | undefined {
  return typeof value === 'string' && value.length > 0 && value.length <= 31
    ? value.toUpperCase()
    : undefined;
}

function columnRange(value: string): XlsxRange | undefined {
  const match = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})$/u.exec(value);
  if (!match) return undefined;
  const start = parseXlsxCellReference(`${match[1]}1`);
  const end = parseXlsxCellReference(`${match[2]}1`);
  if (!start || !end || start.column > end.column) return undefined;
  const reference = `${xlsxColumnName(start.column)!}:${xlsxColumnName(end.column)!}`;
  return {
    end: { column: end.column, row: XLSX_MAX_ROWS },
    reference,
    start: { column: start.column, row: 1 },
  };
}

function rowRange(value: string): XlsxRange | undefined {
  const match = /^(\$?[1-9]\d{0,6}):(\$?[1-9]\d{0,6})$/u.exec(value);
  if (!match) return undefined;
  const start = Number(match[1]!.replace('$', ''));
  const end = Number(match[2]!.replace('$', ''));
  if (start > end || end > XLSX_MAX_ROWS) {
    return undefined;
  }
  return {
    end: { column: XLSX_MAX_COLUMNS, row: end },
    reference: `${start}:${end}`,
    start: { column: 1, row: start },
  };
}

export function parseXlsxSelectionRange(value: unknown): XlsxRange | undefined {
  if (typeof value !== 'string') return undefined;
  return (
    parseXlsxRangeReference(value) ?? columnRange(value) ?? rowRange(value)
  );
}

function sortedRanges(ranges: readonly XlsxRange[]): XlsxRange[] {
  return [...ranges].sort(
    (left, right) =>
      left.start.row - right.start.row ||
      left.start.column - right.start.column ||
      left.end.row - right.end.row ||
      left.end.column - right.end.column,
  );
}

function rangeSelection(
  ranges: readonly XlsxRange[],
): XlsxResolvedSheetSelection {
  const sorted = sortedRanges(ranges);
  let maximumEndRow = 0;
  const endRowPrefix = sorted.map((range) => {
    maximumEndRow = Math.max(maximumEndRow, range.end.row);
    return maximumEndRow;
  });
  return {
    endRowPrefix,
    kind: 'selected-ranges',
    ranges: sorted,
  };
}

function binarySearchIterations(length: number): unknown[] {
  return Array.from({ length: Math.ceil(Math.log2(length + 1)) });
}

function lastRangeStartingAtOrBefore(
  ranges: readonly XlsxRange[],
  row: number,
): number {
  let lower = 0;
  let upper = ranges.length;
  binarySearchIterations(ranges.length).forEach(() => {
    if (lower >= upper) return;
    const middle = lower + Math.floor((upper - lower) / 2);
    if (ranges[middle]!.start.row <= row) lower = middle + 1;
    else upper = middle;
  });
  return lower - 1;
}

function endRowLowerBound(
  endRowPrefix: readonly number[],
  row: number,
  length: number,
): number {
  let lower = 0;
  let upper = length;
  binarySearchIterations(length).forEach(() => {
    const middle = lower + Math.floor((upper - lower) / 2);
    if (endRowPrefix[middle]! < row) lower = middle + 1;
    else upper = middle;
  });
  return lower;
}

export function xlsxSelectionIncludesCell(
  selection: XlsxResolvedSheetSelection,
  row: number,
  column: number,
): boolean {
  if (selection.kind !== 'selected-ranges') {
    return selection.kind === 'full-sheet';
  }
  const window = xlsxSelectionCandidateWindow(selection, row);
  if (!window) return false;
  for (let index = window.last; index >= window.first; index -= 1) {
    const range = selection.ranges[index]!;
    if (
      range.end.row >= row &&
      range.start.column <= column &&
      range.end.column >= column
    ) {
      return true;
    }
  }
  return false;
}

export interface XlsxSelectionCandidateWindow {
  first: number;
  last: number;
}

export function xlsxSelectionCandidateWindow(
  selection: XlsxResolvedSheetSelection,
  row: number,
): XlsxSelectionCandidateWindow | null {
  if (selection.kind !== 'selected-ranges') return null;
  const last = lastRangeStartingAtOrBefore(selection.ranges, row);
  if (last < 0 || selection.endRowPrefix[last]! < row) return null;
  const first = endRowLowerBound(selection.endRowPrefix, row, last + 1);
  return { first, last };
}

export function xlsxSelectionIncludesRow(
  selection: XlsxResolvedSheetSelection,
  row: number,
): boolean {
  if (selection.kind !== 'selected-ranges') {
    return selection.kind === 'full-sheet';
  }
  return xlsxSelectionCandidateWindow(selection, row) !== null;
}

function rangeEntries(selection: XlsxSelection): [string, readonly string[]][] {
  if (selection.ranges === undefined) return [];
  if (Object.prototype.toString.call(selection.ranges) !== '[object Object]') {
    selectionFailure('XLSX selection ranges must be a record of arrays');
  }
  return Object.entries(selection.ranges);
}

export function resolveXlsxSelection(
  selection: XlsxSelection | undefined,
  sheets: readonly XlsxSheet[],
  limits: ResolvedXlsxResourceLimits,
): XlsxResolvedSheetSelection[] {
  if (selection === undefined) {
    return sheets.map(() => ({ kind: 'full-sheet' }));
  }
  if (Object.prototype.toString.call(selection) !== '[object Object]') {
    selectionFailure('XLSX selection must be an object');
  }
  if (
    selection.sheetNames !== undefined &&
    !Array.isArray(selection.sheetNames)
  ) {
    selectionFailure('XLSX selection sheetNames must be an array');
  }

  const sheetIndex = new Map<string, number>();
  for (const [index, sheet] of sheets.entries()) {
    sheetIndex.set(sheet.name.toUpperCase(), index);
  }
  const fullSheets = new Set<number>();
  const selectedNames = new Set<string>();
  for (const value of selection.sheetNames ?? []) {
    const folded = foldedSheetName(value);
    if (folded === undefined) {
      selectionFailure('XLSX selection contains an invalid sheet name');
    }
    if (selectedNames.has(folded)) {
      selectionFailure('XLSX selection contains duplicate sheet names');
    }
    selectedNames.add(folded);
    const index = sheetIndex.get(folded);
    if (index === undefined) {
      selectionFailure('XLSX selection references an unknown sheet');
    }
    fullSheets.add(index);
  }

  const rangesBySheet = new Map<number, XlsxRange[]>();
  const rangeNames = new Set<string>();
  let rangeAreas = 0;
  for (const [name, values] of rangeEntries(selection)) {
    const folded = foldedSheetName(name);
    if (folded === undefined) {
      selectionFailure('XLSX selection contains an invalid range sheet name');
    }
    if (rangeNames.has(folded)) {
      selectionFailure('XLSX selection contains duplicate range sheet names');
    }
    rangeNames.add(folded);
    const index = sheetIndex.get(folded);
    if (index === undefined) {
      selectionFailure('XLSX selection ranges reference an unknown sheet');
    }
    if (sheets[index]!.kind !== 'worksheet') {
      selectionFailure('XLSX selection ranges require a worksheet');
    }
    if (!Array.isArray(values) || values.length === 0) {
      selectionFailure('XLSX selection ranges must be non-empty arrays');
    }
    const normalized = new Map<string, XlsxRange>();
    for (const value of values) {
      rangeAreas += 1;
      if (rangeAreas > limits.maxRangeAreas) {
        throw new XlsxResourceLimitError(
          'maxRangeAreas',
          rangeAreas,
          limits.maxRangeAreas,
        );
      }
      const range = parseXlsxSelectionRange(value);
      if (!range) selectionFailure('XLSX selection contains an invalid range');
      normalized.set(range.reference, range);
    }
    rangesBySheet.set(index, [...normalized.values()]);
  }

  return sheets.map((_, index) => {
    if (fullSheets.has(index)) return { kind: 'full-sheet' };
    const ranges = rangesBySheet.get(index);
    return ranges ? rangeSelection(ranges) : { kind: 'not-selected' };
  });
}
