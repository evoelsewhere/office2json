import { encodeBase64 } from '../../../common/binary/base64';

import { parseXlsxPreservingActiveContent } from '../parser';
import { resolveXlsxResourceLimits } from '../internal/resource-limits';
import { canonicalXlsxSha256 } from './digest';
import { createXlsxCapabilityManifest } from './capability';
import { XlsxWriteError } from './errors';
import {
  inspectXlsxPackageGraph,
  type XlsxPackageGraph,
} from './internal/package-graph';
import {
  createXlsxRoundTripDocument,
  XLSX_KEY_ALGORITHM_VERSION,
} from './keys';
import { normalizeXlsxRoundTripSource } from './source';
import type {
  XlsxRoundTripInput,
  XlsxRoundTripReadOptions,
  XlsxRoundTripSnapshot,
} from './types';
import { resolveXlsxWriteLimits } from './write-limits';

function assertReadOptions(options: XlsxRoundTripReadOptions): void {
  if (
    Object.prototype.toString.call(options) !== '[object Object]' ||
    Object.getPrototypeOf(options) !== Object.prototype
  ) {
    throw new TypeError('XLSX round-trip read options must be a plain object');
  }
  for (const key of Object.keys(options)) {
    if (key !== 'limits' && key !== 'securityMode') {
      throw new TypeError(`Unknown XLSX round-trip read option ${key}`);
    }
  }
  if (
    options.securityMode !== undefined &&
    options.securityMode !== 'reject-active' &&
    options.securityMode !== 'preserve-opaque'
  ) {
    throw new TypeError('XLSX round-trip security mode is invalid');
  }
}

function assertSecurity(
  graph: XlsxPackageGraph,
  securityMode: 'preserve-opaque' | 'reject-active',
): void {
  if (graph.containsActiveContent && securityMode === 'reject-active') {
    throw new XlsxWriteError(
      'preservation-conflict',
      'XLSX round-trip source contains active or embedded content',
      { featureClass: 'active-content' },
    );
  }
}

export async function readXlsxRoundTrip(
  input: XlsxRoundTripInput,
  options: XlsxRoundTripReadOptions = {},
): Promise<XlsxRoundTripSnapshot> {
  assertReadOptions(options);
  const readerLimits = resolveXlsxResourceLimits(options.limits);
  const writeLimits = resolveXlsxWriteLimits(undefined);
  const source = await normalizeXlsxRoundTripSource(
    input,
    readerLimits,
    writeLimits,
  );
  const graph = await inspectXlsxPackageGraph(source.bytes, readerLimits);
  const securityMode = options.securityMode ?? 'reject-active';
  assertSecurity(graph, securityMode);
  const parsed = await parseXlsxPreservingActiveContent(source.bytes, {
    errorMode: 'strict',
    imageMode: 'none',
    limits: options.limits ?? {},
    pivotCacheMode: 'metadata',
  });
  const supportProfile = createXlsxCapabilityManifest();
  const document = await createXlsxRoundTripDocument(
    parsed,
    source.sha256,
    supportProfile.id,
  );
  const baseDocumentHash = await canonicalXlsxSha256(document);
  const snapshot: XlsxRoundTripSnapshot = {
    baseDocumentHash,
    document,
    format: 'xlsx-roundtrip',
    keyAlgorithmVersion: XLSX_KEY_ALGORITHM_VERSION,
    operations: [],
    preservation: {
      containsActiveContent: graph.containsActiveContent,
      containsDigitalSignatures: graph.containsDigitalSignatures,
      containsExternalRelationships: graph.containsExternalRelationships,
      containsOpaqueContent: graph.containsOpaqueContent,
      securityMode,
    },
    schemaVersion: '1',
    source: {
      byteLength: source.byteLength,
      conformance: graph.conformance,
      packageBase64: encodeBase64(source.bytes),
      sha256: source.sha256,
    },
    sourceManifestHash: graph.manifestHash,
    stateHash: baseDocumentHash,
    supportProfile,
  };
  return JSON.parse(JSON.stringify(snapshot)) as XlsxRoundTripSnapshot;
}
