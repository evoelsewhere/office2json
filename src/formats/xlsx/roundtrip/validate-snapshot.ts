import { decodedBase64ByteLength } from '../../../common/binary/base64';

import { canonicalXlsxJson } from './canonical-json';
import { createXlsxCapabilityManifest } from './capability';
import { assertXlsxRoundTripDataTree } from './data-tree';
import { XlsxWriteError } from './errors';
import type { ResolvedXlsxWriteLimits, XlsxRoundTripSnapshot } from './types';

const ROOT_KEYS = [
  'baseDocumentHash',
  'document',
  'format',
  'keyAlgorithmVersion',
  'operations',
  'preservation',
  'schemaVersion',
  'source',
  'sourceManifestHash',
  'stateHash',
  'supportProfile',
] as const;
const SOURCE_KEYS = [
  'byteLength',
  'conformance',
  'packageBase64',
  'sha256',
] as const;
const PRESERVATION_KEYS = [
  'containsActiveContent',
  'containsDigitalSignatures',
  'containsExternalRelationships',
  'containsOpaqueContent',
  'securityMode',
] as const;
const SHA256_PATTERN = /^[0-9a-f]{64}$/u;

function invalid(message: string): never {
  throw new XlsxWriteError('invalid-roundtrip-json', message);
}

function exactRecord(
  value: unknown,
  keys: readonly string[],
  message: string,
): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    invalid(message);
  }
  const names = Object.getOwnPropertyNames(value);
  if (
    names.length !== keys.length ||
    !keys.every((key) => Object.hasOwn(value, key))
  ) {
    invalid(message);
  }
  return value as Record<string, unknown>;
}

function assertSha256(
  value: unknown,
  message: string,
): asserts value is string {
  if (typeof value !== 'string' || !SHA256_PATTERN.test(value))
    invalid(message);
}

export function validateXlsxSnapshotShape(
  value: unknown,
  limits: ResolvedXlsxWriteLimits,
): XlsxRoundTripSnapshot {
  assertXlsxRoundTripDataTree(value, limits);
  const root = exactRecord(
    value,
    ROOT_KEYS,
    'XLSX round-trip snapshot root shape is invalid',
  );
  if (root.format !== 'xlsx-roundtrip') {
    invalid('XLSX round-trip snapshot format is invalid');
  }
  if (root.schemaVersion !== '1') {
    throw new XlsxWriteError(
      'unsupported-snapshot-version',
      'XLSX round-trip snapshot schema version is unsupported',
    );
  }
  if (root.keyAlgorithmVersion !== 'xlsx-snapshot-key-v1') {
    throw new XlsxWriteError(
      'unsupported-snapshot-version',
      'XLSX round-trip key algorithm version is unsupported',
    );
  }
  assertSha256(root.baseDocumentHash, 'XLSX base document hash is invalid');
  assertSha256(root.stateHash, 'XLSX state hash is invalid');
  assertSha256(root.sourceManifestHash, 'XLSX source manifest hash is invalid');
  if (!Array.isArray(root.operations)) {
    invalid('XLSX round-trip operations must be an array');
  }
  if (root.operations.length > limits.maxOperations) {
    throw new XlsxWriteError(
      'resource-limit-exceeded',
      'XLSX round-trip operation count exceeds its limit',
      {
        actual: root.operations.length,
        limit: limits.maxOperations,
        limitName: 'maxOperations',
      },
    );
  }
  const source = exactRecord(
    root.source,
    SOURCE_KEYS,
    'XLSX round-trip source shape is invalid',
  );
  if (
    !Number.isSafeInteger(source.byteLength) ||
    Number(source.byteLength) <= 0
  ) {
    invalid('XLSX round-trip source byteLength is invalid');
  }
  if (
    source.conformance !== 'strict' &&
    source.conformance !== 'transitional'
  ) {
    invalid('XLSX round-trip source conformance is invalid');
  }
  if (typeof source.packageBase64 !== 'string') {
    invalid('XLSX round-trip source Base64 is invalid');
  }
  let decodedLength: number;
  try {
    decodedLength = decodedBase64ByteLength(source.packageBase64);
  } catch (cause) {
    throw new XlsxWriteError(
      'invalid-roundtrip-json',
      'XLSX round-trip source Base64 is invalid',
      {},
      { cause },
    );
  }
  if (decodedLength !== source.byteLength) {
    invalid('XLSX round-trip source byteLength does not match Base64');
  }
  if (decodedLength > limits.maxSourcePackageBytes) {
    throw new XlsxWriteError(
      'resource-limit-exceeded',
      'XLSX source package exceeds its write limit',
      {
        actual: decodedLength,
        limit: limits.maxSourcePackageBytes,
        limitName: 'maxSourcePackageBytes',
      },
    );
  }
  assertSha256(source.sha256, 'XLSX round-trip source SHA-256 is invalid');

  const preservation = exactRecord(
    root.preservation,
    PRESERVATION_KEYS,
    'XLSX round-trip preservation shape is invalid',
  );
  for (const key of PRESERVATION_KEYS.slice(0, 4)) {
    if (typeof preservation[key] !== 'boolean') {
      invalid('XLSX round-trip preservation flag is invalid');
    }
  }
  if (
    preservation.securityMode !== 'reject-active' &&
    preservation.securityMode !== 'preserve-opaque'
  ) {
    invalid('XLSX round-trip security mode is invalid');
  }
  if (
    canonicalXlsxJson(root.supportProfile) !==
    canonicalXlsxJson(createXlsxCapabilityManifest())
  ) {
    throw new XlsxWriteError(
      'unsupported-snapshot-version',
      'XLSX round-trip capability profile is unsupported',
    );
  }
  if (
    root.document === null ||
    typeof root.document !== 'object' ||
    Array.isArray(root.document)
  ) {
    invalid('XLSX round-trip document preview is invalid');
  }
  return root as unknown as XlsxRoundTripSnapshot;
}
