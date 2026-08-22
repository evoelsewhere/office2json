import { XlsxParseError } from '../errors';
import type { XlsxCellValue } from '../types';
import type { XlsxSharedStringTable } from './shared-strings';

const ERROR_CODES = new Set([
  '#BLOCKED!',
  '#BUSY!',
  '#CALC!',
  '#CONNECT!',
  '#DIV/0!',
  '#FIELD!',
  '#GETTING_DATA',
  '#N/A',
  '#NAME?',
  '#NULL!',
  '#NUM!',
  '#REF!',
  '#SPILL!',
  '#UNKNOWN!',
  '#VALUE!',
]);

const NUMBER_PATTERN = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/u;
const DATE_PATTERN = /^(\d{4,})-(\d{2})-(\d{2})$/u;
const TIME_PATTERN =
  /^(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?(?:Z|[+-](?:(?:0\d|1[0-3]):[0-5]\d|14:00))?$/u;
const DURATION_PATTERN =
  /^-?P(?=\d|T\d)(?:\d+Y)?(?:\d+M)?(?:\d+D)?(?:T(?=\d)(?:\d+H)?(?:\d+M)?(?:\d+(?:\.\d+)?S)?)?$/u;

export type XlsxScalarCellType = 'b' | 'd' | 'e' | 'n' | 's' | 'str';

function valueFailure(part: string, cell: string, message: string): never {
  throw new XlsxParseError({
    cell,
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function decimal(value: string, part: string, cell: string): number {
  if (!NUMBER_PATTERN.test(value)) {
    valueFailure(part, cell, 'Cell number is invalid');
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    valueFailure(part, cell, 'Cell number is outside the finite range');
  }
  return parsed === 0 ? 0 : parsed;
}

function daysInMonth(year: string, month: number): number {
  if (month === 2) {
    const tail = Number(year.slice(-4));
    const leap = tail % 400 === 0 || (tail % 4 === 0 && tail % 100 !== 0);
    return leap ? 29 : 28;
  }
  return month === 4 || month === 6 || month === 9 || month === 11 ? 30 : 31;
}

function validDate(value: string): boolean {
  const match = DATE_PATTERN.exec(value);
  if (!match || /^0+$/u.test(match[1]!)) return false;
  const month = Number(match[2]);
  const day = Number(match[3]);
  return (
    month >= 1 &&
    month <= 12 &&
    day >= 1 &&
    day <= daysInMonth(match[1]!, month)
  );
}

function validTime(value: string): boolean {
  const match = TIME_PATTERN.exec(value);
  if (!match) return false;
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  const second = Number(match[3]);
  if (hour > 23 || minute > 59 || second > 59) return false;
  return true;
}

function directDate(value: string, part: string, cell: string): XlsxCellValue {
  if (validDate(value)) {
    return {
      kind: 'date',
      normalized: value,
      precision: 'date',
      source: { kind: 'iso', value },
    };
  }
  const separator = value.indexOf('T');
  const date = value.slice(0, separator);
  const time = value.slice(separator + 1);
  if (validDate(date) && validTime(time)) {
    return {
      kind: 'date',
      normalized: value,
      precision: 'date-time',
      source: { kind: 'iso', value },
    };
  }
  if (validTime(value)) {
    return {
      kind: 'date',
      normalized: value,
      precision: 'time',
      source: { kind: 'iso', value },
    };
  }
  if (DURATION_PATTERN.test(value)) {
    return {
      kind: 'date',
      normalized: value,
      precision: 'duration',
      source: { kind: 'iso', value },
    };
  }
  valueFailure(part, cell, 'Cell ISO date is invalid');
}

function sharedString(
  value: string,
  table: XlsxSharedStringTable,
  part: string,
  cell: string,
): XlsxCellValue {
  if (!/^(?:0|[1-9]\d*)$/u.test(value)) {
    valueFailure(part, cell, 'Cell shared-string index is invalid');
  }
  const index = Number(value);
  if (!Number.isSafeInteger(index) || index >= table.values.length) {
    valueFailure(part, cell, 'Cell shared-string index is out of range');
  }
  const item = table.values[index]!;
  return {
    kind: 'text',
    ...(item.runs === undefined
      ? {}
      : { runs: item.runs.map((run) => ({ text: run.text })) }),
    text: item.text,
  };
}

export function parseXlsxScalarCellValue(
  type: XlsxScalarCellType,
  value: string,
  sharedStrings: XlsxSharedStringTable,
  part: string,
  cell: string,
): XlsxCellValue {
  if (type === 'n') {
    return { kind: 'number', value: decimal(value, part, cell) };
  }
  if (type === 'b') {
    if (value === '0') return { kind: 'boolean', value: false };
    if (value === '1') return { kind: 'boolean', value: true };
    valueFailure(part, cell, 'Cell boolean is invalid');
  }
  if (type === 'e') {
    if (!ERROR_CODES.has(value)) {
      valueFailure(part, cell, 'Cell error code is invalid');
    }
    return { code: value, kind: 'error' };
  }
  if (type === 's') {
    return sharedString(value, sharedStrings, part, cell);
  }
  if (type === 'd') return directDate(value, part, cell);
  return { kind: 'text', text: value };
}
