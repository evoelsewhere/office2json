import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { validateXlsxChartSheetPart } from '../../src/formats/xlsx/internal/chart-sheet';
import { XlsxPartReader } from '../../src/formats/xlsx/internal/part-reader';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import { XLSX_SPREADSHEET_NAMESPACES } from '../../src/formats/xlsx/internal/workbook-discovery';

const PART = 'xl/chartsheets/sheet1.xml';

async function validate(
  xml: string | null,
  dialect: 'strict' | 'transitional' = 'transitional',
): Promise<void> {
  const zip = new JSZip();
  if (xml !== null) zip.file(PART, xml);
  const reader = new XlsxPartReader(zip, [], defaultXlsxResourceLimits());
  await validateXlsxChartSheetPart(PART, dialect, reader);
}

describe('XLSX chart-sheet streaming', () => {
  it.each(['transitional', 'strict'] as const)(
    'validates a %s chart-sheet root without retaining its tree',
    async (dialect) => {
      const namespace = XLSX_SPREADSHEET_NAMESPACES[dialect];
      await expect(
        validate(
          `<s:chartsheet xmlns:s="${namespace}"><s:sheetViews/></s:chartsheet>`,
          dialect,
        ),
      ).resolves.toBeUndefined();
    },
  );

  it.each([
    [
      `<worksheet xmlns="${XLSX_SPREADSHEET_NAMESPACES.transitional}"/>`,
      'chartsheet root is missing or has the wrong namespace',
    ],
    [
      '<chartsheet xmlns="urn:wrong"/>',
      'chartsheet root is missing or has the wrong namespace',
    ],
  ] as const)(
    'rejects an invalid chart-sheet root %#',
    async (xml, message) => {
      await expect(validate(xml)).rejects.toMatchObject({
        diagnostic: {
          code: 'invalid-document-structure',
          message,
          part: PART,
          severity: 'error',
        },
        name: 'XlsxParseError',
      });
    },
  );

  it('requires a selected chart-sheet part', async () => {
    await expect(validate(null)).rejects.toMatchObject({
      diagnostic: {
        code: 'missing-required-part',
        message: `Required XLSX part is missing: ${PART}`,
        part: PART,
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });
});
