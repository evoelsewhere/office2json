import type JSZip from 'jszip';

import { XlsxParseError } from '../errors';
import type { XlsxInput } from '../types';
import { canonicalizeXlsxPartName } from './package-identity';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';

interface ZipEntryMetadata {
  uncompressedSize?: unknown;
}

interface InspectedZipObject extends JSZip.JSZipObject {
  _data?: Promise<unknown> | ZipEntryMetadata;
}

function archiveFailure(message: string, cause: unknown): never {
  throw new XlsxParseError(
    { code: 'invalid-package', message, severity: 'error' },
    { cause },
  );
}

function inputByteLength(input: XlsxInput): number {
  if (input instanceof Blob) return input.size;
  return input.byteLength;
}

export function assertXlsxInputWithinLimits(
  input: XlsxInput,
  limits: ResolvedXlsxResourceLimits,
): void {
  const actual = inputByteLength(input);
  if (actual > limits.maxInputBytes) {
    throw new XlsxResourceLimitError(
      'maxInputBytes',
      actual,
      limits.maxInputBytes,
    );
  }
}

export async function copyXlsxInputBytes(
  input: XlsxInput,
): Promise<Uint8Array> {
  if (input instanceof Blob) {
    return new Uint8Array(await input.arrayBuffer());
  }
  const source =
    input instanceof ArrayBuffer
      ? new Uint8Array(input)
      : new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  const owned = new Uint8Array(source.byteLength);
  owned.set(source);
  return owned;
}

function originalEntryName(file: InspectedZipObject): string {
  return typeof file.unsafeOriginalName === 'string'
    ? file.unsafeOriginalName
    : file.name;
}

function canonicalEntryName(file: InspectedZipObject): string {
  const originalName = originalEntryName(file);
  const partName = file.dir
    ? `${originalName}__directory_entry__`
    : originalName;
  try {
    return canonicalizeXlsxPartName(partName);
  } catch (cause) {
    archiveFailure('Archive contains an invalid part name', cause);
  }
}

function declaredUncompressedSize(file: InspectedZipObject): number {
  const data = file._data;
  // JSZip represents an empty entry as a fulfilled Promise after loadAsync.
  const size = data instanceof Promise ? 0 : data?.uncompressedSize;
  if (typeof size !== 'number' || !Number.isSafeInteger(size) || size < 0) {
    archiveFailure(
      'Archive entry has invalid uncompressed-size metadata',
      new TypeError('Invalid ZIP entry size metadata'),
    );
  }
  return size;
}

export function assertXlsxArchiveWithinLimits(
  zip: JSZip,
  limits: ResolvedXlsxResourceLimits,
): void {
  const entries = Object.values(zip.files) as InspectedZipObject[];
  const canonicalParts = new Set<string>();
  const files: Array<{ entry: InspectedZipObject; part: string }> = [];

  for (const entry of entries) {
    const part = canonicalEntryName(entry);
    if (entry.dir) continue;
    if (canonicalParts.has(part)) {
      archiveFailure(
        'Archive contains duplicate canonical part names',
        new TypeError('Duplicate canonical ZIP part identity'),
      );
    }
    canonicalParts.add(part);
    files.push({ entry, part });
  }

  if (files.length > limits.maxEntries) {
    throw new XlsxResourceLimitError(
      'maxEntries',
      files.length,
      limits.maxEntries,
    );
  }

  let totalUncompressedBytes = 0;
  for (const { entry, part } of files) {
    const size = declaredUncompressedSize(entry);
    if (size > limits.maxPartBytes) {
      throw new XlsxResourceLimitError(
        'maxPartBytes',
        size,
        limits.maxPartBytes,
        part,
      );
    }
    totalUncompressedBytes += size;
    if (
      !Number.isSafeInteger(totalUncompressedBytes) ||
      totalUncompressedBytes > limits.maxTotalUncompressedBytes
    ) {
      throw new XlsxResourceLimitError(
        'maxTotalUncompressedBytes',
        totalUncompressedBytes,
        limits.maxTotalUncompressedBytes,
      );
    }
  }
}
