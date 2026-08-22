import type { XlsxDiagnostic } from './types';

export class XlsxParseError extends Error {
  readonly diagnostic: XlsxDiagnostic;

  constructor(diagnostic: XlsxDiagnostic, options?: ErrorOptions) {
    super(diagnostic.message, options);
    this.name = 'XlsxParseError';
    this.diagnostic = diagnostic;
  }
}
