import type { XlsxWriteOptions } from './types';

export function assertXlsxWriteOptions(options: XlsxWriteOptions): void {
  if (
    Object.prototype.toString.call(options) !== '[object Object]' ||
    Object.getPrototypeOf(options) !== Object.prototype
  ) {
    throw new TypeError('XLSX write options must be a plain object');
  }
  for (const key of Object.keys(options)) {
    if (
      key !== 'acknowledgeOpaqueContent' &&
      key !== 'limits' &&
      key !== 'minimumEditedFidelity' &&
      key !== 'readerLimits'
    ) {
      throw new TypeError(`Unknown XLSX write option ${key}`);
    }
  }
  if (
    options.acknowledgeOpaqueContent !== undefined &&
    typeof options.acknowledgeOpaqueContent !== 'boolean'
  ) {
    throw new TypeError('XLSX acknowledgeOpaqueContent must be boolean');
  }
  if (
    options.minimumEditedFidelity !== undefined &&
    options.minimumEditedFidelity !== 'R1' &&
    options.minimumEditedFidelity !== 'R2' &&
    options.minimumEditedFidelity !== 'R3'
  ) {
    throw new TypeError('XLSX minimumEditedFidelity is invalid');
  }
}
