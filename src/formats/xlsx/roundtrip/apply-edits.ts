import { resolveXlsxResourceLimits } from '../internal/resource-limits';
import { assertXlsxRoundTripDataTree } from './data-tree';
import { replayXlsxCellOperations } from './operation-planner';
import type {
  XlsxEditOperation,
  XlsxRoundTripSnapshot,
  XlsxWriteOptions,
} from './types';
import { verifyXlsxRoundTripSnapshot } from './verify';
import { resolveXlsxWriteLimits } from './write-limits';
import { assertXlsxWriteOptions } from './write-options';

export async function applyXlsxEdits(
  snapshot: XlsxRoundTripSnapshot,
  operations: readonly XlsxEditOperation[],
  options: XlsxWriteOptions = {},
): Promise<XlsxRoundTripSnapshot> {
  assertXlsxWriteOptions(options);
  if (!Array.isArray(operations)) {
    throw new TypeError('XLSX edit operations must be an array');
  }
  const writeLimits = resolveXlsxWriteLimits(options.limits);
  assertXlsxRoundTripDataTree(operations, writeLimits);
  const verified = await verifyXlsxRoundTripSnapshot(
    snapshot,
    options,
    writeLimits,
  );
  const combined: unknown[] = [];
  for (const operation of verified.plan.operations) combined.push(operation);
  for (const operation of operations) combined.push(operation);
  const plan = await replayXlsxCellOperations(
    verified.baseDocument,
    combined,
    writeLimits,
    resolveXlsxResourceLimits(options.readerLimits),
  );
  return structuredClone({
    ...verified.snapshot,
    document: plan.document,
    operations: plan.operations,
    stateHash: plan.stateHash,
  });
}
