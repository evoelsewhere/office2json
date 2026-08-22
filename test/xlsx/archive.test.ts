import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import {
  assertXlsxArchiveWithinLimits,
  assertXlsxInputWithinLimits,
  copyXlsxInputBytes,
} from '../../src/formats/xlsx/internal/archive';
import {
  resolveXlsxResourceLimits,
  XlsxResourceLimitError,
} from '../../src/formats/xlsx/internal/resource-limits';

async function loadedZip(
  parts: Readonly<Record<string, string>>,
): Promise<JSZip> {
  const zip = new JSZip();
  for (const [name, value] of Object.entries(parts)) zip.file(name, value);
  return JSZip.loadAsync(
    await zip.generateAsync({ compression: 'DEFLATE', type: 'uint8array' }),
  );
}

function captureParseError(action: () => unknown): XlsxParseError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected XLSX archive validation to fail');
}

describe('XLSX archive input ownership', () => {
  it.each([
    new Uint8Array([1, 2, 3, 4]),
    new Uint8Array([0, 1, 2, 3, 4, 5]).subarray(1, 5),
    new Uint8Array([1, 2, 3, 4]).buffer,
    new Blob([new Uint8Array([1, 2, 3, 4])]),
  ])('copies caller-owned input without sharing bytes', async (input) => {
    const before =
      input instanceof Blob
        ? new Uint8Array(await input.arrayBuffer())
        : new Uint8Array(
            input instanceof ArrayBuffer
              ? input.slice(0)
              : input.buffer.slice(
                  input.byteOffset,
                  input.byteOffset + input.byteLength,
                ),
          );
    const output = await copyXlsxInputBytes(input);

    expect(output).toEqual(before);
    output[0] = 99;
    const after =
      input instanceof Blob
        ? new Uint8Array(await input.arrayBuffer())
        : new Uint8Array(
            input instanceof ArrayBuffer ? input : input.buffer,
            input instanceof Uint8Array ? input.byteOffset : 0,
            input instanceof Uint8Array ? input.byteLength : input.byteLength,
          );
    expect(after).toEqual(before);
  });

  it('accepts input exactly at maxInputBytes and rejects one over', () => {
    const limits = resolveXlsxResourceLimits({ maxInputBytes: 4 });

    expect(() =>
      assertXlsxInputWithinLimits(new Uint8Array(4), limits),
    ).not.toThrow();
    expect(() =>
      assertXlsxInputWithinLimits(new Uint8Array(5), limits),
    ).toThrow(XlsxResourceLimitError);
    try {
      assertXlsxInputWithinLimits(new Uint8Array(5), limits);
    } catch (error) {
      expect(error).toMatchObject({
        actual: 5,
        limit: 4,
        limitName: 'maxInputBytes',
      });
    }
  });
});

describe('XLSX archive preflight', () => {
  it('accepts entry, part, and total byte counts exactly at limits', async () => {
    const zip = await loadedZip({ 'a.xml': '1234', 'b.xml': '5678' });
    const limits = resolveXlsxResourceLimits({
      maxEntries: 2,
      maxPartBytes: 4,
      maxMediaBytes: 4,
      maxTotalUncompressedBytes: 8,
      maxXmlBytes: 4,
    });

    expect(() => assertXlsxArchiveWithinLimits(zip, limits)).not.toThrow();
  });

  it('validates but does not count directory entries and accepts empty files', async () => {
    const zip = await loadedZip({ 'xl/empty.xml': '' });
    const limits = resolveXlsxResourceLimits({
      maxEntries: 1,
      maxMediaBytes: 1,
      maxPartBytes: 1,
      maxTotalUncompressedBytes: 1,
      maxXmlBytes: 1,
    });

    expect(() => assertXlsxArchiveWithinLimits(zip, limits)).not.toThrow();
  });

  it.each([
    [{ 'a.xml': '1', 'b.xml': '2' }, { maxEntries: 1 }, 'maxEntries', 2, 1],
    [
      { 'a.xml': '12345' },
      { maxMediaBytes: 4, maxPartBytes: 4, maxXmlBytes: 4 },
      'maxPartBytes',
      5,
      4,
    ],
    [
      { 'a.xml': '1234', 'b.xml': '56789' },
      { maxTotalUncompressedBytes: 8 },
      'maxTotalUncompressedBytes',
      9,
      8,
    ],
  ] as const)(
    'rejects one-over archive limit %#',
    async (parts, overrides, limitName, actual, limit) => {
      const zip = await loadedZip(parts);
      const limits = resolveXlsxResourceLimits(overrides);

      expect(() => assertXlsxArchiveWithinLimits(zip, limits)).toThrow(
        XlsxResourceLimitError,
      );
      try {
        assertXlsxArchiveWithinLimits(zip, limits);
      } catch (error) {
        expect(error).toMatchObject({ actual, limit, limitName });
      }
    },
  );

  it.each([
    [
      { 'xl/sharedStrings.xml': 'a', 'xl/shared%53trings.xml': 'b' },
      'Archive contains duplicate canonical part names',
    ],
    [{ 'xl/workbook.xml?query': 'a' }, 'Archive contains an invalid part name'],
    [{ '../outside.xml': 'a' }, 'Archive contains an invalid part name'],
  ] as const)(
    'rejects unsafe archive identities %#',
    async (parts, message) => {
      const zip = await loadedZip(parts);
      const error = captureParseError(() =>
        assertXlsxArchiveWithinLimits(zip, resolveXlsxResourceLimits()),
      );

      expect(error.diagnostic).toEqual({
        code: 'invalid-package',
        message,
        severity: 'error',
      });
      expect(error.cause).toBeInstanceOf(TypeError);
      if (message === 'Archive contains duplicate canonical part names') {
        expect((error.cause as TypeError).message).toBe(
          'Duplicate canonical ZIP part identity',
        );
      }
    },
  );

  it.each([-1, 1.5, Number.MAX_SAFE_INTEGER + 1, '1', undefined])(
    'rejects invalid declared entry size metadata %s',
    async (size) => {
      const zip = await loadedZip({ 'a.xml': '1' });
      const entry = zip.file('a.xml') as JSZip.JSZipObject & {
        _data?: unknown;
      };
      entry._data = { uncompressedSize: size };

      const error = captureParseError(() =>
        assertXlsxArchiveWithinLimits(zip, resolveXlsxResourceLimits()),
      );
      expect(error.diagnostic).toEqual({
        code: 'invalid-package',
        message: 'Archive entry has invalid uncompressed-size metadata',
        severity: 'error',
      });
      expect(error.cause).toEqual(
        new TypeError('Invalid ZIP entry size metadata'),
      );
    },
  );
});
