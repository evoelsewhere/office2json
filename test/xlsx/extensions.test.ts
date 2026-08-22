import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  parseXlsxWithDiagnostics,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
} from '../../src/formats/xlsx';
import {
  createIndependentXlsx,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

function worksheet(body: string): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:payload">${body}</worksheet>`;
}

function extension(uri = 'urn:future:feature'): string {
  return `<ext uri="${uri}"><x:payload>never exposed</x:payload></ext>`;
}

async function source(body: string): Promise<Uint8Array> {
  return createIndependentXlsx({
    'xl/worksheets/sheet1.xml': worksheet(body),
  });
}

const WARNING = {
  code: 'unsupported-feature',
  message: 'Worksheet extension content was omitted',
  part: 'xl/worksheets/sheet1.xml',
  severity: 'warning',
  sheet: 'Sheet1',
} as const;

describe('XLSX worksheet extension diagnostics', () => {
  it.each([
    `<sheetData/><extLst>${extension()}</extLst>`,
    `<sheetViews><sheetView workbookViewId="0"><extLst>${extension()}</extLst></sheetView></sheetViews><sheetData/>`,
    `<sheetData><row><extLst>${extension()}</extLst></row></sheetData>`,
    `<sheetData><row><c><extLst>${extension()}</extLst></c></row></sheetData>`,
  ])(
    'safely omits and diagnoses an unsupported extension owner %#',
    async (body) => {
      const input = await source(body);
      const result = await parseXlsxWithDiagnostics(input);
      expect(result.diagnostics).toStrictEqual([WARNING]);
      expect(JSON.stringify(result)).not.toContain('urn:future:feature');
      expect(JSON.stringify(result)).not.toContain('never exposed');
      await expect(
        parseXlsx(input, { errorMode: 'strict' }),
      ).rejects.toMatchObject({
        diagnostic: { ...WARNING, severity: 'error' },
      });
    },
  );

  it('deduplicates multiple unknown entries to one bounded worksheet warning', async () => {
    const result = await parseXlsxWithDiagnostics(
      await source(
        `<sheetData/><extLst>${extension('urn:a')}${extension('urn:b')}</extLst>`,
      ),
    );
    expect(result.diagnostics).toStrictEqual([WARNING]);
  });

  it('does not misclassify a non-extension worksheet-view payload', async () => {
    const result = await parseXlsxWithDiagnostics(
      await source(
        '<sheetViews><sheetView workbookViewId="0"><pivotSelection/></sheetView></sheetViews><sheetData/>',
      ),
      { errorMode: 'strict' },
    );
    expect(result.diagnostics).toStrictEqual([]);
  });

  it('preserves unknown extension bytes through standalone JSON and exact R0', async () => {
    const input = await source(`<sheetData/><extLst>${extension()}</extLst>`);
    const snapshot = await readXlsxRoundTrip(input);
    const output = await writeXlsxRoundTrip(
      await validateXlsxRoundTripJson(
        JSON.parse(JSON.stringify(snapshot)) as unknown,
      ),
    );
    expect(output.data).toStrictEqual(input);
    expect(output.report.level).toBe('R0');
  });

  it.each([
    [
      '<sheetData/><extLst><ext/></extLst>',
      'Worksheet extension URI is invalid',
    ],
    [
      '<sheetData/><extLst><ext uri=" "/></extLst>',
      'Worksheet extension URI is invalid',
    ],
    [
      '<sheetData/><extLst><x:ext uri="urn:x"/></extLst>',
      'Worksheet extension entry is invalid',
    ],
    [
      '<sheetData/><extLst><wrong uri="urn:x"/></extLst>',
      'Worksheet extension entry is invalid',
    ],
  ] as const)(
    'rejects malformed worksheet extension structure %# in both modes',
    async (body, message) => {
      const input = await source(body);
      await expect(parseXlsx(input)).rejects.toMatchObject({
        diagnostic: { message },
      });
      await expect(
        parseXlsx(input, { errorMode: 'strict' }),
      ).rejects.toMatchObject({ diagnostic: { message } });
    },
  );
});
