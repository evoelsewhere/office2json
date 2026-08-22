import { XlsxWriteError } from './errors';
import type { ResolvedXlsxWriteLimits } from './types';
import { writeLimitFailure } from './write-limits';

interface PendingValue {
  depth: number;
  value: unknown;
}

function invalid(message: string): never {
  throw new XlsxWriteError('invalid-roundtrip-json', message);
}

function dataProperty(value: object, name: string): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, name);
  if (
    descriptor === undefined ||
    !('value' in descriptor) ||
    !descriptor.enumerable
  ) {
    invalid('XLSX round-trip JSON contains an accessor or hidden property');
  }
  return descriptor.value;
}

function consumeUtf8(
  value: string,
  state: { bytes: number },
  limits: ResolvedXlsxWriteLimits,
): void {
  state.bytes += new TextEncoder().encode(value).byteLength;
  if (
    !Number.isSafeInteger(state.bytes) ||
    state.bytes > limits.maxSnapshotJsonBytes
  ) {
    writeLimitFailure(
      'maxSnapshotJsonBytes',
      state.bytes,
      limits.maxSnapshotJsonBytes,
    );
  }
}

export function assertXlsxRoundTripDataTree(
  value: unknown,
  limits: ResolvedXlsxWriteLimits,
): void {
  const pending: PendingValue[] = [{ depth: 0, value }];
  const seen = new WeakSet<object>();
  const stringState = { bytes: 0 };
  let objects = 0;
  for (const current of pending) {
    if (current.depth > limits.maxSnapshotDepth) {
      writeLimitFailure(
        'maxSnapshotDepth',
        current.depth,
        limits.maxSnapshotDepth,
      );
    }
    if (typeof current.value === 'string') {
      consumeUtf8(current.value, stringState, limits);
      continue;
    }
    if (current.value === null || typeof current.value === 'boolean') continue;
    if (typeof current.value === 'number') {
      if (Number.isFinite(current.value)) continue;
      invalid('XLSX round-trip JSON contains a non-finite number');
    }
    if (typeof current.value !== 'object') {
      invalid('XLSX round-trip JSON contains a non-JSON value');
    }
    if (seen.has(current.value)) {
      invalid('XLSX round-trip JSON contains a repeated object reference');
    }
    seen.add(current.value);
    objects += 1;
    if (objects > limits.maxSnapshotObjects) {
      writeLimitFailure(
        'maxSnapshotObjects',
        objects,
        limits.maxSnapshotObjects,
      );
    }
    if (Object.getOwnPropertySymbols(current.value).length !== 0) {
      invalid('XLSX round-trip JSON contains a symbol key');
    }
    if (Array.isArray(current.value)) {
      if (Object.getPrototypeOf(current.value) !== Array.prototype) {
        invalid('XLSX round-trip JSON requires plain arrays');
      }
      if (
        Object.getOwnPropertyNames(current.value).length !==
        current.value.length + 1
      ) {
        invalid('XLSX round-trip JSON contains a sparse or extended array');
      }
      pending.push(
        ...current.value.map((_child, index) => ({
          depth: current.depth + 1,
          value: dataProperty(current.value as object, String(index)),
        })),
      );
      continue;
    }
    if (Object.getPrototypeOf(current.value) !== Object.prototype) {
      invalid('XLSX round-trip JSON requires plain objects');
    }
    for (const name of Object.getOwnPropertyNames(current.value)) {
      consumeUtf8(name, stringState, limits);
      pending.push({
        depth: current.depth + 1,
        value: dataProperty(current.value, name),
      });
    }
  }
}
