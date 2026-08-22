import { XlsxWriteError } from './errors';
import type { ResolvedXlsxWriteLimits, XlsxWriteLimits } from './types';

export function defaultXlsxWriteLimits(): ResolvedXlsxWriteLimits {
  return {
    maxDependencyEdges: 1_000_000,
    maxDirtyParts: 10_000,
    maxFormulaRewriteTokens: 1_000_000,
    maxGeneratedMediaBytes: 100_000_000,
    maxGeneratedXmlBytes: 100_000_000,
    maxOperationBytes: 1_000_000,
    maxOperations: 100_000,
    maxOutputBytes: 200_000_000,
    maxPatchBytes: 100_000_000,
    maxPatchCount: 1_000_000,
    maxPatchedParts: 10_000,
    maxReferenceUpdates: 1_000_000,
    maxSnapshotDepth: 64,
    maxSnapshotJsonBytes: 300_000_000,
    maxSnapshotObjects: 2_000_000,
    maxSourcePackageBytes: 200_000_000,
    maxTotalOperationBytes: 100_000_000,
    maxValidationPasses: 8,
  };
}

export function resolveXlsxWriteLimits(
  overrides: XlsxWriteLimits | undefined,
): ResolvedXlsxWriteLimits {
  const resolved = { ...defaultXlsxWriteLimits() };
  if (overrides !== undefined) {
    if (
      Object.prototype.toString.call(overrides) !== '[object Object]' ||
      Object.getPrototypeOf(overrides) !== Object.prototype
    ) {
      throw new TypeError('XLSX write limits must be a plain object');
    }
    for (const key of Object.keys(overrides) as Array<keyof XlsxWriteLimits>) {
      if (!(key in resolved))
        throw new TypeError(`Unknown XLSX write limit ${key}`);
      const value = overrides[key];
      if (!Number.isSafeInteger(value) || Number(value) <= 0) {
        throw new TypeError(`${key} must be a positive safe integer`);
      }
      resolved[key] = value!;
    }
  }
  if (resolved.maxOperationBytes > resolved.maxTotalOperationBytes) {
    throw new TypeError(
      'maxOperationBytes cannot exceed maxTotalOperationBytes',
    );
  }
  if (resolved.maxSourcePackageBytes > resolved.maxOutputBytes) {
    throw new TypeError('maxSourcePackageBytes cannot exceed maxOutputBytes');
  }
  return resolved;
}

export function writeLimitFailure(
  limitName: keyof XlsxWriteLimits,
  actual: number,
  limit: number,
  part?: string,
): never {
  throw new XlsxWriteError(
    'resource-limit-exceeded',
    `XLSX write resource limit ${limitName} exceeded`,
    { actual, limit, limitName, ...(part === undefined ? {} : { part }) },
  );
}
