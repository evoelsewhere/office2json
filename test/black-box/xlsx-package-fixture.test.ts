import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  createIndependentXlsx,
  independentWorkbook,
  independentWorksheet,
  XLSX_OFFICE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from './xlsx-package';

describe('independent XLSX package fixture', () => {
  it('contains a literal minimal workbook graph', async () => {
    const zip = await JSZip.loadAsync(await createIndependentXlsx());

    expect(
      Object.keys(zip.files).filter((name) => !zip.files[name]?.dir),
    ).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels',
      'xl/worksheets/sheet1.xml',
      'xl/sharedStrings.xml',
      'xl/styles.xml',
    ]);
    await expect(
      zip.file('xl/workbook.xml')?.async('string'),
    ).resolves.toContain('<sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/>');
    await expect(
      zip.file('xl/worksheets/sheet1.xml')?.async('string'),
    ).resolves.toContain('<c r="A1" t="s"><v>0</v></c>');
  });

  it('replaces and removes exact parts without production helpers', async () => {
    const replacement = independentWorksheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Replaced</t></is></c></row>',
    );
    const zip = await JSZip.loadAsync(
      await createIndependentXlsx({
        'xl/sharedStrings.xml': null,
        'xl/worksheets/sheet1.xml': replacement,
      }),
    );

    expect(zip.file('xl/sharedStrings.xml')).toBeNull();
    await expect(
      zip.file('xl/worksheets/sheet1.xml')?.async('string'),
    ).resolves.toBe(replacement);
  });

  it('exposes raw namespace-specific workbook and worksheet builders', () => {
    expect(
      independentWorkbook('<sheet name="Data" sheetId="7" r:id="rIdData"/>'),
    ).toContain(`xmlns:r="${XLSX_OFFICE_REL_NS}"`);
    expect(independentWorksheet('<row/>')).toContain(
      `xmlns="${XLSX_SPREADSHEET_NS}"`,
    );
  });

  it.each(['DEFLATE', 'STORE'] as const)(
    'supports %s without changing uncompressed XML',
    async (compression) => {
      const zip = await JSZip.loadAsync(
        await createIndependentXlsx({}, { compression }),
      );

      await expect(
        zip.file('xl/sharedStrings.xml')?.async('string'),
      ).resolves.toContain('<si><t>Black box</t></si>');
    },
  );
});
