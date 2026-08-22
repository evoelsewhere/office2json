export { applyXlsxEdits } from './apply-edits';
export { createXlsxCapabilityManifest } from './capability';
export { XlsxWriteError } from './errors';
export { readXlsxRoundTrip } from './read-snapshot';
export type {
  XlsxCapabilityManifest,
  XlsxEditOperation,
  XlsxRoundTripDocument,
  XlsxRoundTripReadOptions,
  XlsxRoundTripSnapshot,
  XlsxWriteDiagnostic,
  XlsxWriteDiagnosticCode,
  XlsxWriteLimits,
  XlsxWriteOptions,
  XlsxWriteReport,
  XlsxWriteResult,
} from './types';
export { validateXlsxRoundTripJson } from './validate-json';
export { writeXlsxRoundTrip } from './write';
