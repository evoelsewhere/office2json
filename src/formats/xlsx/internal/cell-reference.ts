import type { XlsxRange } from '../types';
import { XLSX_MAX_COLUMNS, XLSX_MAX_ROWS } from './resource-limits';

export interface XlsxParsedCellReference {
  absoluteColumn: boolean;
  absoluteRow: boolean;
  address: string;
  column: number;
  row: number;
}

function parseColumnName(value: string): number | undefined {
  let column = 0;
  for (const character of value) {
    const code = character.toUpperCase().codePointAt(0)!;
    column = column * 26 + (code - 0x40);
    if (!Number.isSafeInteger(column) || column > XLSX_MAX_COLUMNS) {
      return undefined;
    }
  }
  return column;
}

export function xlsxColumnName(column: number): string | undefined {
  if (
    !Number.isSafeInteger(column) ||
    column < 1 ||
    column > XLSX_MAX_COLUMNS
  ) {
    return undefined;
  }
  let remaining = column;
  let output = '';
  Array.from({ length: 3 }).forEach(() => {
    if (remaining === 0) return;
    const offset = (remaining - 1) % 26;
    output = String.fromCodePoint(0x41 + offset) + output;
    remaining = Math.floor((remaining - 1) / 26);
  });
  return output;
}

export function parseXlsxCellReference(
  value: unknown,
): XlsxParsedCellReference | undefined {
  if (typeof value !== 'string') return undefined;
  const match = /^(\$?)([A-Za-z]+)(\$?)([1-9]\d*)$/u.exec(value);
  if (!match) return undefined;
  const columnName = match[2]!;
  const column = parseColumnName(columnName);
  const row = Number(match[4]);
  if (
    column === undefined ||
    !Number.isSafeInteger(row) ||
    row > XLSX_MAX_ROWS
  ) {
    return undefined;
  }
  const canonicalColumn = xlsxColumnName(column)!;
  return {
    absoluteColumn: match[1] === '$',
    absoluteRow: match[3] === '$',
    address: `${canonicalColumn}${row}`,
    column,
    row,
  };
}

export function parseXlsxRangeReference(value: unknown): XlsxRange | undefined {
  if (typeof value !== 'string') return undefined;
  const references = value.split(':');
  if (references.length > 2) return undefined;
  const startText = references[0]!;
  const endText = references.at(-1)!;
  const start = parseXlsxCellReference(startText);
  const end = parseXlsxCellReference(endText);
  if (!start || !end || start.row > end.row || start.column > end.column) {
    return undefined;
  }
  const reference =
    start.address === end.address
      ? start.address
      : `${start.address}:${end.address}`;
  return {
    end: { column: end.column, row: end.row },
    reference,
    start: { column: start.column, row: start.row },
  };
}
