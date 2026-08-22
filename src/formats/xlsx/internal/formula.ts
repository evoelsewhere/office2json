import { parseXlsxCellReference, xlsxColumnName } from './cell-reference';
import { XLSX_MAX_COLUMNS, XLSX_MAX_ROWS } from './resource-limits';

export interface XlsxFormulaCoordinate {
  column: number;
  row: number;
}

function validCoordinate(value: XlsxFormulaCoordinate): boolean {
  return (
    Number.isSafeInteger(value.column) &&
    value.column >= 1 &&
    value.column <= XLSX_MAX_COLUMNS &&
    Number.isSafeInteger(value.row) &&
    value.row >= 1 &&
    value.row <= XLSX_MAX_ROWS
  );
}

function identifierCharacter(value: string | undefined): boolean {
  return value !== undefined && /[A-Za-z0-9_.]/u.test(value);
}

function qualifierCharacter(value: string): boolean {
  return /[A-Za-z0-9_$.:[\]]/u.test(value);
}

function isSheetQualifier(expression: string, end: number): boolean {
  for (const character of expression.slice(end)) {
    if (character === '!' || character === '[') return true;
    if (!qualifierCharacter(character)) return false;
  }
  return false;
}

function consumeQuoted(
  expression: string,
  start: number,
  quote: '"' | "'",
): number | undefined {
  const end = expression.indexOf(quote, start + 1);
  return end === -1 ? undefined : end + 1;
}

function shiftedReference(
  source: string,
  rowDelta: number,
  columnDelta: number,
): string | undefined {
  const parsed = parseXlsxCellReference(source);
  if (!parsed) return source;
  const column = parsed.absoluteColumn
    ? parsed.column
    : parsed.column + columnDelta;
  const row = parsed.absoluteRow ? parsed.row : parsed.row + rowDelta;
  if (
    column < 1 ||
    column > XLSX_MAX_COLUMNS ||
    row < 1 ||
    row > XLSX_MAX_ROWS
  ) {
    return undefined;
  }
  return `${parsed.absoluteColumn ? '$' : ''}${xlsxColumnName(column)!}${
    parsed.absoluteRow ? '$' : ''
  }${row}`;
}

/**
 * Translates the relative A1 tokens in a shared formula without evaluating it.
 * The scanner deliberately treats strings, sheet qualifiers, external-book
 * qualifiers, and structured-reference brackets as opaque formula tokens.
 */
export function translateXlsxSharedFormula(
  expression: string,
  source: XlsxFormulaCoordinate,
  target: XlsxFormulaCoordinate,
): string | undefined {
  if (!validCoordinate(source) || !validCoordinate(target)) return undefined;
  const rowDelta = target.row - source.row;
  const columnDelta = target.column - source.column;
  if (rowDelta === 0 && columnDelta === 0) return expression;

  let bracketDepth = 0;
  let skipUntil = 0;
  let output = '';
  for (const [cursor, character] of expression.split('').entries()) {
    if (cursor < skipUntil) continue;
    if (character === '"' || character === "'") {
      const end = consumeQuoted(expression, cursor, character);
      if (end === undefined) return undefined;
      output += expression.slice(cursor, end);
      skipUntil = end;
      continue;
    }
    if (character === '[') {
      bracketDepth += 1;
      output += character;
      continue;
    }
    if (character === ']') {
      bracketDepth -= 1;
      output += character;
      continue;
    }
    if (bracketDepth > 0 || identifierCharacter(expression[cursor - 1])) {
      output += character;
      continue;
    }

    const candidate = /^(?:\$?)[A-Za-z]+(?:\$?)[1-9]\d*/u.exec(
      expression.slice(cursor),
    )?.[0];
    if (candidate === undefined) {
      output += character;
      continue;
    }
    const end = cursor + candidate.length;
    if (
      identifierCharacter(expression[end]) ||
      expression[end] === '(' ||
      isSheetQualifier(expression, end)
    ) {
      output += candidate;
      skipUntil = end;
      continue;
    }
    const shifted = shiftedReference(candidate, rowDelta, columnDelta);
    if (shifted === undefined) return undefined;
    output += shifted;
    skipUntil = end;
  }
  return bracketDepth === 0 ? output : undefined;
}
