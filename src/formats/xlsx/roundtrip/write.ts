import { parseXlsx } from '../parser';
import { resolveXlsxResourceLimits } from '../internal/resource-limits';
import { canonicalXlsxJson } from './canonical-json';
import { writeXlsxCellEditPackage } from './cell-edit-package';
import { createXlsxCapabilityManifest } from './capability';
import { sha256XlsxBytes } from './digest';
import { XlsxWriteError } from './errors';
import { resolveXlsxWriteLimits, writeLimitFailure } from './write-limits';
import { assertXlsxWriteOptions } from './write-options';
import type {
  XlsxRoundTripDocument,
  XlsxRoundTripSnapshot,
  XlsxWriteOptions,
  XlsxWriteResult,
} from './types';
import { verifyXlsxRoundTripSnapshot } from './verify';

export async function verifyXlsxEditedSemantics(
  data: Uint8Array,
  expected: XlsxRoundTripDocument,
  options: XlsxWriteOptions,
): Promise<void> {
  let parsed: Awaited<ReturnType<typeof parseXlsx>>;
  try {
    parsed = await parseXlsx(data, {
      errorMode: 'strict',
      imageMode: 'none',
      limits: options.readerLimits ?? {},
      pivotCacheMode: 'metadata',
    });
  } catch {
    throw new XlsxWriteError(
      'semantic-verification-failed',
      'Strictly reparsing the edited XLSX package failed',
      { fidelity: 'R2' },
    );
  }
  if (parsed.sheets.length !== expected.sheets.length) {
    throw new XlsxWriteError(
      'semantic-verification-failed',
      'Edited XLSX worksheet inventory differs from the operation preview',
      { fidelity: 'R2' },
    );
  }
  const actual: XlsxRoundTripDocument = {
    ...parsed,
    key: expected.key,
    sheets: parsed.sheets.map((sheet, index) => ({
      ...sheet,
      key: expected.sheets[index]!.key,
    })),
  };
  if (canonicalXlsxJson(actual) !== canonicalXlsxJson(expected)) {
    throw new XlsxWriteError(
      'semantic-verification-failed',
      'Edited XLSX semantics differ from the operation preview',
      { fidelity: 'R2' },
    );
  }
}

export async function writeXlsxRoundTrip(
  value: XlsxRoundTripSnapshot,
  options: XlsxWriteOptions = {},
): Promise<XlsxWriteResult> {
  assertXlsxWriteOptions(options);
  const writeLimits = resolveXlsxWriteLimits(options.limits);
  const verified = await verifyXlsxRoundTripSnapshot(
    value,
    options,
    writeLimits,
  );
  if (verified.plan.operations.length !== 0) {
    if (options.minimumEditedFidelity === 'R3') {
      throw new XlsxWriteError(
        'producer-verification-failed',
        'The XLSX cell-edit profile has no producer R3 evidence',
        { fidelity: 'R3' },
      );
    }
    const validationPasses = 4;
    if (validationPasses > writeLimits.maxValidationPasses) {
      writeLimitFailure(
        'maxValidationPasses',
        validationPasses,
        writeLimits.maxValidationPasses,
      );
    }
    const readerLimits = resolveXlsxResourceLimits(options.readerLimits);
    const edited = await writeXlsxCellEditPackage(
      verified.bytes,
      verified.graph,
      verified.baseDocument,
      verified.plan,
      options,
      writeLimits,
      readerLimits,
    );
    await verifyXlsxEditedSemantics(
      edited.data,
      verified.plan.document,
      options,
    );
    const formula = verified.plan.operations.find(
      (operation) =>
        operation.kind === 'set-cell' && operation.content.kind === 'formula',
    );
    return {
      data: edited.data,
      report: {
        diagnostics:
          formula === undefined
            ? []
            : [
                {
                  code: 'recalculation-required',
                  message:
                    'The edited XLSX formula has no cached result and requires producer recalculation',
                  operationId: formula.operationId,
                  severity: 'warning',
                },
              ],
        level: 'R2',
        outputSha256: await sha256XlsxBytes(edited.data),
        parts: edited.parts,
        sourceSha256: verified.snapshot.source.sha256,
        supportProfile: createXlsxCapabilityManifest(),
      },
    };
  }
  return {
    data: verified.bytes,
    report: {
      diagnostics: [],
      level: 'R0',
      outputSha256: verified.snapshot.source.sha256,
      parts: verified.graph.parts.map((part) => ({
        byteLength: part.byteLength,
        disposition: 'copy',
        name: part.name,
        sha256: part.sha256,
        sourceByteLength: part.byteLength,
        sourceSha256: part.sha256,
      })),
      sourceSha256: verified.snapshot.source.sha256,
      supportProfile: createXlsxCapabilityManifest(),
    },
  };
}
