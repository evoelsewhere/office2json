import { beforeAll, describe, expect, it } from 'vitest';

import { readXlsxRoundTrip } from '../../src/formats/xlsx/roundtrip';
import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import type { XlsxRoundTripSnapshot } from '../../src/formats/xlsx/roundtrip/types';
import { validateXlsxSnapshotShape } from '../../src/formats/xlsx/roundtrip/validate-snapshot';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import { createIndependentXlsx } from '../black-box/xlsx-package';

function clone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

function capture(value: unknown): XlsxWriteError {
  try {
    validateXlsxSnapshotShape(value, defaultXlsxWriteLimits());
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected snapshot shape validation to fail');
}

describe('XLSX round-trip snapshot schema', () => {
  let source: XlsxRoundTripSnapshot;

  beforeAll(async () => {
    source = await readXlsxRoundTrip(await createIndependentXlsx());
  });

  it('accepts the exact portable snapshot shape', () => {
    expect(
      validateXlsxSnapshotShape(clone(source), defaultXlsxWriteLimits()),
    ).toEqual(source);
  });

  it.each([
    [null, 'XLSX round-trip snapshot root shape is invalid'],
    [[], 'XLSX round-trip snapshot root shape is invalid'],
    ['snapshot', 'XLSX round-trip snapshot root shape is invalid'],
  ])('rejects invalid root %#', (value, message) => {
    expect(capture(value).diagnostic).toMatchObject({
      code: 'invalid-roundtrip-json',
      message,
    });
  });

  it('rejects missing and unknown root fields', () => {
    const missing = clone(source) as unknown as Record<string, unknown>;
    delete missing.stateHash;
    expect(capture(missing).diagnostic.message).toBe(
      'XLSX round-trip snapshot root shape is invalid',
    );
    const extra = { ...clone(source), extra: true };
    expect(capture(extra).diagnostic.message).toBe(
      'XLSX round-trip snapshot root shape is invalid',
    );
    const sameLength = clone(source) as unknown as Record<string, unknown>;
    delete sameLength.stateHash;
    sameLength.otherHash = '0'.repeat(64);
    expect(capture(sameLength).diagnostic.message).toBe(
      'XLSX round-trip snapshot root shape is invalid',
    );
  });

  it.each([
    [
      'format',
      'xlsx',
      'invalid-roundtrip-json',
      'XLSX round-trip snapshot format is invalid',
    ],
    [
      'schemaVersion',
      '2',
      'unsupported-snapshot-version',
      'XLSX round-trip snapshot schema version is unsupported',
    ],
    [
      'keyAlgorithmVersion',
      'other',
      'unsupported-snapshot-version',
      'XLSX round-trip key algorithm version is unsupported',
    ],
  ] as const)(
    'rejects invalid root literal %s',
    (key, value, code, message) => {
      const snapshot = clone(source) as unknown as Record<string, unknown>;
      snapshot[key] = value;
      expect(capture(snapshot).diagnostic).toMatchObject({ code, message });
    },
  );

  it.each([
    [
      'baseDocumentHash',
      `${'0'.repeat(63)}g`,
      'XLSX base document hash is invalid',
    ],
    ['stateHash', '0'.repeat(63), 'XLSX state hash is invalid'],
    [
      'sourceManifestHash',
      '0'.repeat(65),
      'XLSX source manifest hash is invalid',
    ],
    [
      'sourceManifestHash',
      'A'.repeat(64),
      'XLSX source manifest hash is invalid',
    ],
  ] as const)('rejects invalid hash %s', (key, value, message) => {
    const snapshot = clone(source) as unknown as Record<string, unknown>;
    snapshot[key] = value;
    expect(capture(snapshot).diagnostic.message).toBe(message);
  });

  it('validates operations type and exact count boundary', () => {
    const wrong = clone(source) as unknown as Record<string, unknown>;
    wrong.operations = {};
    expect(capture(wrong).diagnostic.message).toBe(
      'XLSX round-trip operations must be an array',
    );
    const operation = {
      cell: 'A1',
      kind: 'clear-cell',
      operationId: 'one',
      sheetKey: source.document.sheets[0]!.key,
    };
    const one = clone(source);
    one.operations = [operation] as never;
    expect(() =>
      validateXlsxSnapshotShape(one, {
        ...defaultXlsxWriteLimits(),
        maxOperations: 1,
      }),
    ).not.toThrow();
    const two = clone(source);
    two.operations = [operation, { ...operation, operationId: 'two' }] as never;
    expect(
      captureWithLimits(two, { maxOperations: 1 }).diagnostic,
    ).toMatchObject({
      actual: 2,
      code: 'resource-limit-exceeded',
      limit: 1,
      limitName: 'maxOperations',
      message: 'XLSX round-trip operation count exceeds its limit',
    });
  });

  it('rejects invalid source shape, identity, conformance, and Base64', () => {
    const extra = clone(source) as unknown as {
      source: Record<string, unknown>;
    };
    extra.source.extra = true;
    expect(capture(extra).diagnostic.message).toBe(
      'XLSX round-trip source shape is invalid',
    );
    for (const byteLength of [0, -1, 1.5, '1']) {
      const snapshot = clone(source) as unknown as {
        source: Record<string, unknown>;
      };
      snapshot.source.byteLength = byteLength;
      expect(capture(snapshot).diagnostic.message).toBe(
        'XLSX round-trip source byteLength is invalid',
      );
    }
    const conformance = clone(source);
    conformance.source.conformance = 'unknown' as 'strict';
    expect(capture(conformance).diagnostic.message).toBe(
      'XLSX round-trip source conformance is invalid',
    );
    const type = clone(source) as unknown as {
      source: Record<string, unknown>;
    };
    type.source.packageBase64 = 1;
    expect(capture(type).diagnostic.message).toBe(
      'XLSX round-trip source Base64 is invalid',
    );
    const lexical = clone(source);
    lexical.source.packageBase64 = '!bad';
    const lexicalError = capture(lexical);
    expect(lexicalError.diagnostic.message).toBe(
      'XLSX round-trip source Base64 is invalid',
    );
    expect(lexicalError.cause).toBeInstanceOf(RangeError);
    const length = clone(source);
    length.source.byteLength += 1;
    expect(capture(length).diagnostic.message).toBe(
      'XLSX round-trip source byteLength does not match Base64',
    );
    const sha = clone(source);
    sha.source.sha256 = 'g'.repeat(64);
    expect(capture(sha).diagnostic.message).toBe(
      'XLSX round-trip source SHA-256 is invalid',
    );
    expect(() =>
      validateXlsxSnapshotShape(clone(source), {
        ...defaultXlsxWriteLimits(),
        maxSourcePackageBytes: source.source.byteLength,
      }),
    ).not.toThrow();
    expect(
      captureWithLimits(source, {
        maxSourcePackageBytes: source.source.byteLength - 1,
      }).diagnostic,
    ).toMatchObject({
      actual: source.source.byteLength,
      code: 'resource-limit-exceeded',
      limit: source.source.byteLength - 1,
      limitName: 'maxSourcePackageBytes',
      message: 'XLSX source package exceeds its write limit',
    });
  });

  it('rejects preservation shape, flags, security mode, and profile changes', () => {
    const extra = clone(source) as unknown as {
      preservation: Record<string, unknown>;
    };
    extra.preservation.extra = true;
    expect(capture(extra).diagnostic.message).toBe(
      'XLSX round-trip preservation shape is invalid',
    );
    for (const key of [
      'containsActiveContent',
      'containsDigitalSignatures',
      'containsExternalRelationships',
      'containsOpaqueContent',
    ] as const) {
      const snapshot = clone(source) as unknown as {
        preservation: Record<string, unknown>;
      };
      snapshot.preservation[key] = 'false';
      expect(capture(snapshot).diagnostic.message).toBe(
        'XLSX round-trip preservation flag is invalid',
      );
    }
    const security = clone(source);
    security.preservation.securityMode = 'unsafe' as 'reject-active';
    expect(capture(security).diagnostic.message).toBe(
      'XLSX round-trip security mode is invalid',
    );
    const profile = clone(source);
    profile.supportProfile.id = 'other' as 'xlsx-agent-ready';
    expect(capture(profile).diagnostic).toMatchObject({
      code: 'unsupported-snapshot-version',
      message: 'XLSX round-trip capability profile is unsupported',
    });
  });

  it.each([null, [], 'document'])(
    'rejects invalid document preview %#',
    (document) => {
      const snapshot = clone(source) as unknown as Record<string, unknown>;
      snapshot.document = document;
      expect(capture(snapshot).diagnostic.message).toBe(
        'XLSX round-trip document preview is invalid',
      );
    },
  );

  function captureWithLimits(
    value: unknown,
    overrides: Partial<ReturnType<typeof defaultXlsxWriteLimits>>,
  ): XlsxWriteError {
    try {
      validateXlsxSnapshotShape(value, {
        ...defaultXlsxWriteLimits(),
        ...overrides,
      });
    } catch (error) {
      expect(error).toBeInstanceOf(XlsxWriteError);
      return error as XlsxWriteError;
    }
    throw new Error('Expected limited snapshot validation to fail');
  }
});
