import type { XlsxWriteDiagnostic, XlsxWriteDiagnosticCode } from './types';

export class XlsxWriteError extends Error {
  readonly diagnostic: XlsxWriteDiagnostic;

  constructor(
    code: XlsxWriteDiagnosticCode,
    message: string,
    fields: Omit<XlsxWriteDiagnostic, 'code' | 'message' | 'severity'> = {},
    options: ErrorOptions = {},
  ) {
    super(message, options);
    this.name = 'XlsxWriteError';
    this.diagnostic = {
      code,
      message,
      severity: 'error',
      ...fields,
    };
  }
}
