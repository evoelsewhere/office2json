import type { XlsxCellValue } from '../types';

type XlsxDatePrecision = Extract<XlsxCellValue, { kind: 'date' }>['precision'];

interface DayFraction {
  nanoseconds: number;
  wholeDays: number;
}

interface CivilDate {
  day: number;
  month: number;
  year: number;
}

function nanosecondsPerDay(): number {
  return 86_400_000_000_000;
}

function splitDay(value: number): DayFraction {
  let wholeDays = Math.floor(value);
  let nanoseconds = Math.round((value - wholeDays) * nanosecondsPerDay());
  if (nanoseconds === nanosecondsPerDay()) {
    wholeDays += 1;
    nanoseconds = 0;
  }
  return { nanoseconds, wholeDays };
}

function civilFromDays(days: number): CivilDate | undefined {
  const shifted = days + 719_468;
  const era = Math.floor(shifted / 146_097);
  const dayOfEra = shifted - era * 146_097;
  const yearOfEra = Math.floor(
    (dayOfEra -
      Math.floor(dayOfEra / 1460) +
      Math.floor(dayOfEra / 36_524) -
      Math.floor(dayOfEra / 146_096)) /
      365,
  );
  let year = yearOfEra + era * 400;
  const dayOfYear =
    dayOfEra -
    (365 * yearOfEra + Math.floor(yearOfEra / 4) - Math.floor(yearOfEra / 100));
  const monthPosition = Math.floor((5 * dayOfYear + 2) / 153);
  const day = dayOfYear - Math.floor((153 * monthPosition + 2) / 5) + 1;
  const month = monthPosition + (monthPosition < 10 ? 3 : -9);
  year += month <= 2 ? 1 : 0;
  if (year > 9999) return undefined;
  return { day, month, year };
}

function pad(value: number, length: number): string {
  return value.toString().padStart(length, '0');
}

function formatDate(value: CivilDate): string {
  return `${pad(value.year, 4)}-${pad(value.month, 2)}-${pad(value.day, 2)}`;
}

function formatSeconds(seconds: number, nanoseconds: number): string {
  const base = pad(seconds, 2);
  if (nanoseconds === 0) return base;
  return `${base}.${pad(nanoseconds, 9).replace(/0+$/u, '')}`;
}

function formatDurationSeconds(seconds: number, nanoseconds: number): string {
  if (nanoseconds === 0) return seconds.toString();
  return `${seconds}.${pad(nanoseconds, 9).replace(/0+$/u, '')}`;
}

function formatTime(nanoseconds: number): string {
  const nanosecondsPerSecond = 1_000_000_000;
  const totalSeconds = Math.floor(nanoseconds / nanosecondsPerSecond);
  const hour = Math.floor(totalSeconds / 3600);
  const minute = Math.floor((totalSeconds % 3600) / 60);
  const second = totalSeconds % 60;
  const remainder = nanoseconds % nanosecondsPerSecond;
  return `${pad(hour, 2)}:${pad(minute, 2)}:${formatSeconds(
    second,
    remainder,
  )}`;
}

function serialDate(
  split: DayFraction,
  dateSystem: '1900' | '1904',
): CivilDate | undefined {
  if (dateSystem === '1900') {
    if (split.wholeDays < 60) {
      return civilFromDays(-25_568 + split.wholeDays);
    }
    if (split.wholeDays === 60) return undefined;
    return civilFromDays(-25_568 + split.wholeDays - 1);
  }
  return civilFromDays(-24_107 + split.wholeDays);
}

function formatDuration(value: number): string {
  const split = splitDay(Math.abs(value));
  if (split.wholeDays === 0 && split.nanoseconds === 0) return 'PT0S';
  const days = split.wholeDays;
  let remainder = split.nanoseconds;
  const hourNanoseconds = 3_600_000_000_000;
  const minuteNanoseconds = 60_000_000_000;
  const secondNanoseconds = 1_000_000_000;
  const hours = Math.floor(remainder / hourNanoseconds);
  remainder %= hourNanoseconds;
  const minutes = Math.floor(remainder / minuteNanoseconds);
  remainder %= minuteNanoseconds;
  const seconds = Math.floor(remainder / secondNanoseconds);
  const nanoseconds = remainder % secondNanoseconds;

  let output = Math.sign(value) === -1 ? '-P' : 'P';
  if (days !== 0) output += `${days}D`;
  if (split.nanoseconds !== 0) {
    output += 'T';
    if (hours !== 0) output += `${hours}H`;
    if (minutes !== 0) output += `${minutes}M`;
    if (seconds !== 0 || nanoseconds !== 0) {
      output += `${formatDurationSeconds(seconds, nanoseconds)}S`;
    }
  }
  return output;
}

export function normalizeXlsxSerialDate(
  value: number,
  dateSystem: '1900' | '1904',
  precision: XlsxDatePrecision,
): string | null {
  if (!Number.isSafeInteger(Math.floor(Math.abs(value)))) return null;
  if (precision === 'duration') return formatDuration(value);
  const split = splitDay(value);
  if (split.wholeDays < 0) return null;
  if (precision === 'time') return formatTime(split.nanoseconds);
  const date = serialDate(split, dateSystem);
  if (!date) return null;
  const normalizedDate = formatDate(date);
  return precision === 'date'
    ? normalizedDate
    : `${normalizedDate}T${formatTime(split.nanoseconds)}`;
}
