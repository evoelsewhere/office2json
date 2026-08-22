import { describe, expect, it } from 'vitest';

import { canonicalXlsxJson } from '../../src/formats/xlsx/roundtrip/canonical-json';
import { assertXlsxRoundTripDataTree } from '../../src/formats/xlsx/roundtrip/data-tree';
import {
  canonicalXlsxSha256,
  sha256XlsxBytes,
  sha256XlsxText,
} from '../../src/formats/xlsx/roundtrip/digest';
import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import { createXlsxCapabilityManifest } from '../../src/formats/xlsx/roundtrip/capability';
import { createXlsxRoundTripDocument } from '../../src/formats/xlsx/roundtrip/keys';
import { normalizeXlsxRoundTripSource } from '../../src/formats/xlsx/roundtrip/source';
import {
  defaultXlsxWriteLimits,
  resolveXlsxWriteLimits,
  writeLimitFailure,
} from '../../src/formats/xlsx/roundtrip/write-limits';
import { resolveXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import { parseXlsx } from '../../src/formats/xlsx/parser';
import { createIndependentXlsx } from '../black-box/xlsx-package';

function capture(action: () => unknown): XlsxWriteError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected XLSX write kernel to fail');
}

describe('XLSX round-trip canonical kernel', () => {
  it('canonicalizes plain JSON with recursively sorted object keys', () => {
    expect(
      canonicalXlsxJson({ z: 1, a: { y: true, b: ['é', null, -0] } }),
    ).toBe('{"a":{"b":["é",null,0],"y":true},"z":1}');
    expect(canonicalXlsxJson({ b: 1, a: 2 })).toBe(
      canonicalXlsxJson({ a: 2, b: 1 }),
    );
  });

  it.each([NaN, Infinity, -Infinity, undefined, 1n, Symbol('x'), () => 1])(
    'rejects non-canonical value %#',
    (value) => {
      expect(() => canonicalXlsxJson(value)).toThrow(TypeError);
    },
  );

  it('rejects cycles, non-plain objects, arrays, and excessive depth', () => {
    const cycle: Record<string, unknown> = {};
    cycle.self = cycle;
    expect(() => canonicalXlsxJson(cycle)).toThrow('cycles');
    expect(() => canonicalXlsxJson(new Date(0))).toThrow('plain objects');
    const subclass: unknown[] = [];
    Object.setPrototypeOf(subclass, { custom: true });
    expect(() => canonicalXlsxJson(subclass)).toThrow('plain arrays');
    const nested = (count: number): unknown => {
      let value: unknown = null;
      Array.from({ length: count }).forEach(() => {
        value = { child: value };
      });
      return value;
    };
    expect(() => canonicalXlsxJson(nested(128))).not.toThrow();
    expect(() => canonicalXlsxJson(nested(129))).toThrow('maximum depth');
    const nestedArray = (count: number): unknown => {
      let value: unknown = null;
      Array.from({ length: count }).forEach(() => {
        value = [value];
      });
      return value;
    };
    expect(() => canonicalXlsxJson(nestedArray(128))).not.toThrow();
    expect(() => canonicalXlsxJson(nestedArray(129))).toThrow('maximum depth');
  });

  it('returns stable canonical error messages', () => {
    expect(() => canonicalXlsxJson(NaN)).toThrow(
      'Canonical XLSX JSON requires finite numbers',
    );
    expect(() => canonicalXlsxJson(undefined)).toThrow(
      'Canonical XLSX JSON does not support undefined',
    );
    const cycle: Record<string, unknown> = {};
    cycle.self = cycle;
    expect(() => canonicalXlsxJson(cycle)).toThrow(
      'Canonical XLSX JSON does not support cycles',
    );
  });

  it('generates source-bound stable and distinct workbook and sheet keys', async () => {
    const bytes = await createIndependentXlsx();
    const document = await parseXlsx(bytes);
    const first = await createXlsxRoundTripDocument(
      document,
      'a'.repeat(64),
      'profile',
    );
    const repeated = await createXlsxRoundTripDocument(
      document,
      'a'.repeat(64),
      'profile',
    );
    const sourceChanged = await createXlsxRoundTripDocument(
      document,
      'b'.repeat(64),
      'profile',
    );
    const profileChanged = await createXlsxRoundTripDocument(
      document,
      'a'.repeat(64),
      'other-profile',
    );
    expect(first).toEqual(repeated);
    expect(first.key).not.toBe(sourceChanged.key);
    expect(first.key).not.toBe(profileChanged.key);
    expect(first.sheets[0]?.key).not.toBe(sourceChanged.sheets[0]?.key);
  });

  it('copies every source input form and enforces the direct source limit', async () => {
    const bytes = await createIndependentXlsx();
    const readerLimits = resolveXlsxResourceLimits();
    const exactLimits = {
      ...defaultXlsxWriteLimits(),
      maxSourcePackageBytes: bytes.byteLength,
    };
    const normalized = await normalizeXlsxRoundTripSource(
      bytes,
      readerLimits,
      exactLimits,
    );
    normalized.bytes.fill(0);
    expect(bytes.some((byte) => byte !== 0)).toBe(true);
    const arrayBuffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer;
    const normalizedBuffer = await normalizeXlsxRoundTripSource(
      arrayBuffer,
      readerLimits,
      exactLimits,
    );
    normalizedBuffer.bytes.fill(0);
    expect(new Uint8Array(arrayBuffer).some((byte) => byte !== 0)).toBe(true);
    const normalizedBlob = await normalizeXlsxRoundTripSource(
      new Blob([arrayBuffer]),
      readerLimits,
      exactLimits,
    );
    expect(normalizedBlob.byteLength).toBe(bytes.byteLength);
    const error = await (async () => {
      try {
        await normalizeXlsxRoundTripSource(bytes, readerLimits, {
          ...exactLimits,
          maxSourcePackageBytes: bytes.byteLength - 1,
        });
      } catch (reason) {
        expect(reason).toBeInstanceOf(XlsxWriteError);
        return reason as XlsxWriteError;
      }
      throw new Error('Expected source limit failure');
    })();
    expect(error.diagnostic).toMatchObject({
      actual: bytes.byteLength,
      limit: bytes.byteLength - 1,
      limitName: 'maxSourcePackageBytes',
    });
  });

  it('hashes bytes, text, and canonical values with browser-neutral SHA-256', async () => {
    const expected =
      'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad';
    expect(await sha256XlsxBytes(new TextEncoder().encode('abc'))).toBe(
      expected,
    );
    expect(await sha256XlsxText('abc')).toBe(expected);
    expect(await canonicalXlsxSha256({ b: 1, a: 2 })).toBe(
      await canonicalXlsxSha256({ a: 2, b: 1 }),
    );
  });

  it('validates a plain non-repeating JSON data tree without invoking getters', () => {
    const limits = defaultXlsxWriteLimits();
    expect(() =>
      assertXlsxRoundTripDataTree({ a: [1, true, null, 'text'] }, limits),
    ).not.toThrow();
    let invoked = false;
    const accessor = Object.defineProperty({}, 'danger', {
      enumerable: true,
      get() {
        invoked = true;
        return 'secret';
      },
    });
    expect(
      capture(() => assertXlsxRoundTripDataTree(accessor, limits)).diagnostic
        .message,
    ).toBe('XLSX round-trip JSON contains an accessor or hidden property');
    expect(invoked).toBe(false);
  });

  it('rejects repeated references, symbols, sparse arrays, prototypes, and values', () => {
    const limits = defaultXlsxWriteLimits();
    const shared = {};
    const invalidValues: Array<readonly [unknown, string]> = [
      [
        { left: shared, right: shared },
        'XLSX round-trip JSON contains a repeated object reference',
      ],
      [{ [Symbol('key')]: true }, 'XLSX round-trip JSON contains a symbol key'],
      [
        Object.assign(Object.create(null), { value: 1 }),
        'XLSX round-trip JSON requires plain objects',
      ],
      [[new Uint8Array([1])], 'XLSX round-trip JSON requires plain objects'],
      [{ value: undefined }, 'XLSX round-trip JSON contains a non-JSON value'],
      [{ value: 1n }, 'XLSX round-trip JSON contains a non-JSON value'],
      [{ value: NaN }, 'XLSX round-trip JSON contains a non-finite number'],
    ];
    const sparse: unknown[] = [];
    sparse.length = 2;
    sparse[0] = 1;
    invalidValues.push([
      sparse,
      'XLSX round-trip JSON contains a sparse or extended array',
    ]);
    const subclass: unknown[] = [];
    Object.setPrototypeOf(subclass, { custom: true });
    invalidValues.push([
      subclass,
      'XLSX round-trip JSON requires plain arrays',
    ]);
    for (const [value, message] of invalidValues) {
      expect(
        capture(() => assertXlsxRoundTripDataTree(value, limits)).diagnostic
          .message,
      ).toBe(message);
    }
  });

  it('enforces JSON byte, depth, and object budgets exactly', () => {
    expect(() =>
      assertXlsxRoundTripDataTree('é', {
        ...defaultXlsxWriteLimits(),
        maxSnapshotJsonBytes: 2,
      }),
    ).not.toThrow();
    expect(
      capture(() =>
        assertXlsxRoundTripDataTree('é', {
          ...defaultXlsxWriteLimits(),
          maxSnapshotJsonBytes: 1,
        }),
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      code: 'resource-limit-exceeded',
      limit: 1,
      limitName: 'maxSnapshotJsonBytes',
    });
    expect(() =>
      assertXlsxRoundTripDataTree(
        { child: {} },
        {
          ...defaultXlsxWriteLimits(),
          maxSnapshotDepth: 1,
          maxSnapshotObjects: 2,
        },
      ),
    ).not.toThrow();
    expect(
      capture(() =>
        assertXlsxRoundTripDataTree(
          { child: {} },
          {
            ...defaultXlsxWriteLimits(),
            maxSnapshotDepth: 1,
            maxSnapshotObjects: 1,
          },
        ),
      ).diagnostic.limitName,
    ).toBe('maxSnapshotObjects');
    expect(
      capture(() =>
        assertXlsxRoundTripDataTree(
          { child: { leaf: 1 } },
          {
            ...defaultXlsxWriteLimits(),
            maxSnapshotDepth: 1,
          },
        ),
      ).diagnostic.limitName,
    ).toBe('maxSnapshotDepth');
    expect(() =>
      assertXlsxRoundTripDataTree([[1]], {
        ...defaultXlsxWriteLimits(),
        maxSnapshotDepth: 2,
      }),
    ).not.toThrow();
    expect(
      capture(() =>
        assertXlsxRoundTripDataTree([[1]], {
          ...defaultXlsxWriteLimits(),
          maxSnapshotDepth: 1,
        }),
      ).diagnostic.limitName,
    ).toBe('maxSnapshotDepth');
  });

  it('resolves every write limit and rejects unknown, unsafe, or conflicting values', () => {
    const defaults = defaultXlsxWriteLimits();
    const exact = resolveXlsxWriteLimits({
      maxOperationBytes: 2,
      maxOutputBytes: 4,
      maxSourcePackageBytes: 3,
      maxTotalOperationBytes: 2,
    });
    expect(exact).toMatchObject({
      maxOperationBytes: 2,
      maxOutputBytes: 4,
      maxSourcePackageBytes: 3,
      maxTotalOperationBytes: 2,
    });
    expect(Object.keys(defaults).sort()).toEqual(Object.keys(exact).sort());
    for (const value of [0, -1, 1.5, NaN, Infinity, 2 ** 53]) {
      expect(() => resolveXlsxWriteLimits({ maxOperations: value })).toThrow(
        TypeError,
      );
    }
    expect(() => resolveXlsxWriteLimits({ unknown: 1 } as never)).toThrow(
      'Unknown XLSX write limit',
    );
    expect(() => resolveXlsxWriteLimits(null as never)).toThrow(TypeError);
    expect(() => resolveXlsxWriteLimits([] as never)).toThrow(
      'XLSX write limits must be a plain object',
    );
    expect(() => resolveXlsxWriteLimits(new Date() as never)).toThrow(
      'XLSX write limits must be a plain object',
    );
    expect(() => resolveXlsxWriteLimits(Object.create(null) as never)).toThrow(
      'XLSX write limits must be a plain object',
    );
    const spoofed = Object.defineProperty({}, Symbol.toStringTag, {
      value: 'Spoofed',
    });
    expect(() => resolveXlsxWriteLimits(spoofed)).toThrow(
      'XLSX write limits must be a plain object',
    );
    expect(() => resolveXlsxWriteLimits({ maxOperations: 0 })).toThrow(
      'maxOperations must be a positive safe integer',
    );
    expect(() =>
      resolveXlsxWriteLimits({
        maxOperationBytes: 3,
        maxTotalOperationBytes: 2,
      }),
    ).toThrow('maxOperationBytes');
    expect(() =>
      resolveXlsxWriteLimits({
        maxOutputBytes: 2,
        maxSourcePackageBytes: 3,
      }),
    ).toThrow('maxSourcePackageBytes');
  });

  it('emits bounded structured resource errors', () => {
    const error = capture(() =>
      writeLimitFailure('maxOperations', 2, 1, 'xl/workbook.xml'),
    );
    expect(error.name).toBe('XlsxWriteError');
    expect(error.diagnostic).toEqual({
      actual: 2,
      code: 'resource-limit-exceeded',
      limit: 1,
      limitName: 'maxOperations',
      message: 'XLSX write resource limit maxOperations exceeded',
      part: 'xl/workbook.xml',
      severity: 'error',
    });
  });

  it('publishes stable domain and operation capability inventories', () => {
    const manifest = createXlsxCapabilityManifest();
    expect(manifest.domains.map((entry) => entry.domain)).toStrictEqual([
      'active-content',
      'calculation',
      'cells',
      'charts',
      'comments',
      'conditional-formatting',
      'connections',
      'defined-names',
      'document-properties',
      'drawings-images',
      'external-links',
      'filters-sorts',
      'formulas',
      'hyperlinks',
      'known-extensions',
      'merges',
      'modern-cell-metadata',
      'pivots',
      'print-layout',
      'protection',
      'rows-columns',
      'shared-strings',
      'sheet-metadata',
      'sparklines',
      'styles',
      'tables',
      'unknown-extensions',
      'validation',
      'views',
      'workbook-sheets',
    ]);
    expect(manifest.operations.map((entry) => entry.operation)).toStrictEqual([
      'add-worksheet',
      'clear-cell',
      'delete-columns',
      'delete-rows',
      'delete-worksheet',
      'insert-columns',
      'insert-rows',
      'rename-worksheet',
      'set-cell',
      'set-cell-style',
      'set-column',
      'set-hyperlink',
      'set-row',
    ]);
  });
});
