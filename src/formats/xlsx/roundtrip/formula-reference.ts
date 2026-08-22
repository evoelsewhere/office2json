import {
  parseXlsxCellReference,
  xlsxColumnName,
} from '../internal/cell-reference';
import { XLSX_MAX_COLUMNS, XLSX_MAX_ROWS } from '../internal/resource-limits';
import type { XlsxRange } from '../types';
import {
  transformXlsxStructuralRange,
  type XlsxStructuralReferenceOperation,
} from './structural-reference';

export type XlsxStructuralSourceFormulaResult =
  | { expression: string; kind: 'preserved' | 'transformed' }
  | { kind: 'deleted' }
  | { kind: 'unsupported' };

interface SourceTokens {
  end: string;
  prefix: string;
  qualifier?: string;
  start: string;
}

function sameSheet(left: string, right: string): boolean {
  return left.toUpperCase() === right.toUpperCase();
}

function decodedQuotedQualifier(quoted: string): string {
  return quoted.slice(1, -1).replaceAll("''", "'");
}

function validBareQualifier(value: string): boolean {
  if (value.length === 0) return false;
  for (const character of value) {
    if (
      character.trim().length === 0 ||
      "'![]:+*/^&=<>(),;{}".includes(character)
    ) {
      return false;
    }
  }
  return true;
}

function sourceTokens(expression: string): SourceTokens | undefined {
  let reference = expression;
  let prefix = '';
  let qualifier: string | undefined;
  if (expression.startsWith("'")) {
    const qualified = /('(?:[^']|'')+')!/u.exec(expression);
    if (qualified === null) return undefined;
    const quoted = qualified[1]!;
    qualifier = decodedQuotedQualifier(quoted);
    prefix = qualified[0];
    reference = expression.slice(prefix.length);
  } else {
    const separator = expression.indexOf('!');
    if (separator !== -1) {
      const bare = expression.slice(0, separator);
      if (!validBareQualifier(bare)) return undefined;
      qualifier = bare;
      prefix = expression.slice(0, separator + 1);
      reference = expression.slice(separator + 1);
    }
  }
  const areas = reference.split(':');
  if (areas.length > 2) return undefined;
  return {
    end: areas.at(-1)!,
    prefix,
    ...(qualifier === undefined ? {} : { qualifier }),
    start: areas[0]!,
  };
}

function authoredReference(
  source: ReturnType<typeof parseXlsxCellReference> & object,
  column: number,
  row: number,
): string {
  return `${source.absoluteColumn ? '$' : ''}${xlsxColumnName(column)!}${
    source.absoluteRow ? '$' : ''
  }${row}`;
}

export function xlsxStructuralSourceFormulaArea(
  expression: string,
): number | undefined {
  const source = sourceTokens(expression);
  if (source === undefined) return undefined;
  const start = parseXlsxCellReference(source.start);
  const end = parseXlsxCellReference(source.end);
  if (!start || !end || start.row > end.row || start.column > end.column) {
    return undefined;
  }
  return (end.row - start.row + 1) * (end.column - start.column + 1);
}

/**
 * Rewrites the single local A1 cell/range grammar used by sparkline source
 * formulas. Names, external books, 3-D ranges, unions, and arbitrary formula
 * expressions are deliberately rejected so a structural edit cannot guess at
 * dependencies.
 */
export function transformXlsxStructuralSourceFormula(
  expression: string,
  ownerSheetName: string,
  targetSheetName: string,
  operation: XlsxStructuralReferenceOperation,
): XlsxStructuralSourceFormulaResult {
  const source = sourceTokens(expression);
  if (source === undefined) return { kind: 'unsupported' };
  const applies = sameSheet(
    source.qualifier ?? ownerSheetName,
    targetSheetName,
  );
  if (!applies) return { expression, kind: 'preserved' };
  const startSource = source.start;
  const endSource = source.end;
  const start = parseXlsxCellReference(startSource);
  const end = parseXlsxCellReference(endSource);
  if (!start || !end || start.row > end.row || start.column > end.column) {
    return { kind: 'unsupported' };
  }
  const range: XlsxRange = {
    end: { column: end.column, row: end.row },
    reference: start.address,
    start: { column: start.column, row: start.row },
  };
  const transformed = transformXlsxStructuralRange(range, operation);
  if (transformed === null) return { kind: 'deleted' };
  if (
    transformed.end.column > XLSX_MAX_COLUMNS ||
    transformed.end.row > XLSX_MAX_ROWS
  ) {
    return { kind: 'unsupported' };
  }
  const transformedStart = authoredReference(
    start,
    transformed.start.column,
    transformed.start.row,
  );
  const transformedEnd = authoredReference(
    end,
    transformed.end.column,
    transformed.end.row,
  );
  const reference =
    endSource === startSource
      ? transformedStart
      : `${transformedStart}:${transformedEnd}`;
  const output = `${source.prefix}${reference}`;
  return {
    expression: output,
    kind: output === expression ? 'preserved' : 'transformed',
  };
}
