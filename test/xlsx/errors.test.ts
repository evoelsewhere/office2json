import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import type { XlsxDiagnostic } from '../../src/formats/xlsx/types';

describe('XlsxParseError', () => {
  it('carries the typed diagnostic and cause', () => {
    const diagnostic: XlsxDiagnostic = {
      code: 'invalid-document-structure',
      message: 'Workbook root is missing',
      part: 'xl/workbook.xml',
      severity: 'error',
    };
    const cause = new Error('invalid root');
    const error = new XlsxParseError(diagnostic, { cause });

    expect(error).toMatchObject({
      cause,
      diagnostic,
      message: diagnostic.message,
      name: 'XlsxParseError',
    });
  });
});
