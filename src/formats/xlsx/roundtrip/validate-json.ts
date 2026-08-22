import { resolveXlsxWriteLimits } from './write-limits';
import type { XlsxRoundTripSnapshot, XlsxWriteOptions } from './types';
import { verifyXlsxRoundTripSnapshot } from './verify';
import { assertXlsxWriteOptions } from './write-options';

export async function validateXlsxRoundTripJson(
  value: unknown,
  options: XlsxWriteOptions = {},
): Promise<XlsxRoundTripSnapshot> {
  assertXlsxWriteOptions(options);
  const writeLimits = resolveXlsxWriteLimits(options.limits);
  const verified = await verifyXlsxRoundTripSnapshot(
    value,
    options,
    writeLimits,
  );
  return structuredClone(verified.snapshot);
}
