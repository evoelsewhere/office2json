import { describe, expect, it } from 'vitest';

import { normalizeXlsxSerialDate } from '../../src/formats/xlsx/internal/date-system';

describe('XLSX serial date normalization', () => {
  it.each([
    [0, '1899-12-31'],
    [1, '1900-01-01'],
    [31, '1900-01-31'],
    [32, '1900-02-01'],
    [59, '1900-02-28'],
    [60, null],
    [60.5, null],
    [61, '1900-03-01'],
    [36585, '2000-02-29'],
  ] as const)('normalizes 1900-system date serial %s', (serial, expected) => {
    expect(normalizeXlsxSerialDate(serial, '1900', 'date')).toBe(expected);
  });

  it.each([
    [0, '1904-01-01'],
    [59, '1904-02-29'],
    [60, '1904-03-01'],
    [35123, '2000-02-29'],
  ] as const)('normalizes 1904-system date serial %s', (serial, expected) => {
    expect(normalizeXlsxSerialDate(serial, '1904', 'date')).toBe(expected);
  });

  it('normalizes time-only values independently of the compatibility date', () => {
    expect(normalizeXlsxSerialDate(0, '1900', 'time')).toBe('00:00:00');
    expect(normalizeXlsxSerialDate(0.5, '1900', 'time')).toBe('12:00:00');
    expect(normalizeXlsxSerialDate(60.25, '1900', 'time')).toBe('06:00:00');
    expect(normalizeXlsxSerialDate(1 / 86_400, '1900', 'time')).toBe(
      '00:00:01',
    );
    expect(normalizeXlsxSerialDate(0.5 / 86_400, '1900', 'time')).toBe(
      '00:00:00.5',
    );
    expect(normalizeXlsxSerialDate(0.101 / 86_400, '1900', 'time')).toBe(
      '00:00:00.101',
    );
  });

  it('carries a rounded fractional day into the next calendar date', () => {
    const almostOne = 0.999_999_999_999_999_9;
    expect(normalizeXlsxSerialDate(almostOne, '1900', 'date')).toBe(
      '1900-01-01',
    );
    expect(normalizeXlsxSerialDate(almostOne, '1900', 'time')).toBe('00:00:00');
  });

  it('normalizes date-time values without host date or time-zone APIs', () => {
    expect(normalizeXlsxSerialDate(61.5, '1900', 'date-time')).toBe(
      '1900-03-01T12:00:00',
    );
    expect(normalizeXlsxSerialDate(0.25, '1904', 'date-time')).toBe(
      '1904-01-01T06:00:00',
    );
    expect(normalizeXlsxSerialDate(60.5, '1900', 'date-time')).toBeNull();
  });

  it.each([
    [0, 'PT0S'],
    [-0, 'PT0S'],
    [0.5, 'PT12H'],
    [1.5, 'P1DT12H'],
    [-1.5, '-P1DT12H'],
    [1 / 86_400, 'PT1S'],
    [1 / 1440, 'PT1M'],
    [0.5 / 86_400, 'PT0.5S'],
    [1 / 86_400_000_000_000, 'PT0.000000001S'],
    [1 + 1 / 24 + 1 / 1440 + 1 / 86_400, 'P1DT1H1M1S'],
  ] as const)('normalizes duration serial %s', (serial, expected) => {
    expect(normalizeXlsxSerialDate(serial, '1900', 'duration')).toBe(expected);
  });

  it('distinguishes zero, day-only, and time-bearing durations', () => {
    expect(normalizeXlsxSerialDate(0, '1900', 'duration')).toBe('PT0S');
    expect(normalizeXlsxSerialDate(1, '1900', 'duration')).toBe('P1D');
    expect(normalizeXlsxSerialDate(0.5, '1900', 'duration')).toBe('PT12H');
  });

  it.each([Number.NaN, Infinity, -Infinity, Number.MAX_VALUE])(
    'returns null for non-normalizable serial %s',
    (serial) => {
      expect(normalizeXlsxSerialDate(serial, '1900', 'date')).toBeNull();
      expect(normalizeXlsxSerialDate(serial, '1900', 'duration')).toBeNull();
    },
  );

  it.each(['date', 'time', 'date-time'] as const)(
    'does not invent a calendar/time value for a negative %s serial',
    (precision) => {
      expect(normalizeXlsxSerialDate(-0.25, '1900', precision)).toBeNull();
    },
  );

  it('returns null outside the bounded four-digit calendar domain', () => {
    expect(normalizeXlsxSerialDate(2_958_465, '1900', 'date')).toBe(
      '9999-12-31',
    );
    expect(normalizeXlsxSerialDate(2_958_466, '1900', 'date')).toBeNull();
    expect(
      normalizeXlsxSerialDate(Number.MAX_SAFE_INTEGER, '1900', 'date'),
    ).toBeNull();
  });
});
