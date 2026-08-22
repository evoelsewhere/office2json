import { assertXlsxInputWithinLimits } from '../internal/archive';
import type { ResolvedXlsxResourceLimits } from '../internal/resource-limits';
import type { XlsxInput } from '../types';
import { sha256XlsxBytes } from './digest';
import { writeLimitFailure } from './write-limits';
import type { ResolvedXlsxWriteLimits } from './types';

export interface NormalizedXlsxRoundTripSource {
  byteLength: number;
  bytes: Uint8Array;
  sha256: string;
}

export async function normalizeXlsxRoundTripSource(
  input: XlsxInput,
  readerLimits: ResolvedXlsxResourceLimits,
  writeLimits: ResolvedXlsxWriteLimits,
): Promise<NormalizedXlsxRoundTripSource> {
  assertXlsxInputWithinLimits(input, readerLimits);
  const bytes =
    input instanceof Blob
      ? new Uint8Array(await input.arrayBuffer())
      : input instanceof ArrayBuffer
        ? new Uint8Array(input.slice(0))
        : input.slice();
  if (bytes.byteLength > writeLimits.maxSourcePackageBytes) {
    writeLimitFailure(
      'maxSourcePackageBytes',
      bytes.byteLength,
      writeLimits.maxSourcePackageBytes,
    );
  }
  return {
    byteLength: bytes.byteLength,
    bytes,
    sha256: await sha256XlsxBytes(bytes),
  };
}
