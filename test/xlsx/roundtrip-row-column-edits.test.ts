import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  applyXlsxEdits,
  parseXlsx,
  readXlsxRoundTrip,
  writeXlsxRoundTrip,
  XlsxWriteError,
} from '../../src/formats/xlsx';
import { patchXlsxWorksheetPartWithReport } from '../../src/formats/xlsx/roundtrip/worksheet-patch';
import { patchXlsxWorksheetProperties } from '../../src/formats/xlsx/roundtrip/worksheet-properties-patch';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import {
  createIndependentXlsx,
  XLSX_CONTENT_TYPES_NS,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

function portable<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

async function capture(
  action: () => Promise<unknown>,
): Promise<XlsxWriteError> {
  try {
    await action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected row or column edit to fail');
}

describe('XLSX verified row and column property edits', () => {
  it('sets existing sizes and visibility with deterministic R1/R2 evidence', async () => {
    const source = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><cols><col min="1" max="2" width="10" customWidth="1" style="0"/><col min="3" max="3" width="8" customWidth="1"/></cols><sheetData><row r="1" ht="12" customHeight="1" outlineLevel="1"><c r="A1"><f>1+1</f><v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData></worksheet>`,
    });
    const snapshot = await readXlsxRoundTrip(source);
    const sheet = snapshot.document.sheets[0]!;
    const edited = await applyXlsxEdits(snapshot, [
      {
        height: 20,
        kind: 'set-row',
        operationId: 'row-height',
        row: 1,
        sheetKey: sheet.key,
      },
      {
        hidden: true,
        kind: 'set-row',
        operationId: 'row-hidden',
        row: 1,
        sheetKey: sheet.key,
      },
      {
        hidden: false,
        kind: 'set-row',
        operationId: 'row-two',
        row: 2,
        sheetKey: sheet.key,
      },
      {
        end: 2,
        kind: 'set-column',
        operationId: 'column-width',
        sheetKey: sheet.key,
        start: 1,
        width: 25,
      },
      {
        end: 2,
        hidden: false,
        kind: 'set-column',
        operationId: 'column-visible',
        sheetKey: sheet.key,
        start: 1,
      },
      {
        end: 3,
        hidden: true,
        kind: 'set-column',
        operationId: 'column-three',
        sheetKey: sheet.key,
        start: 3,
      },
    ]);
    const first = await writeXlsxRoundTrip(portable(edited));
    const second = await writeXlsxRoundTrip(portable(edited));
    expect(second).toEqual(first);
    expect(first.report.level).toBe('R2');
    expect(
      first.report.parts
        .filter((part) => part.disposition === 'patch')
        .map((part) => part.name),
    ).toEqual(['xl/worksheets/sheet1.xml']);
    const parsed = await parseXlsx(first.data, { errorMode: 'strict' });
    const parsedSheet = parsed.sheets[0]!;
    expect(parsedSheet.kind).toBe('worksheet');
    if (parsedSheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(parsedSheet.rows[0]).toMatchObject({
      height: 20,
      hidden: true,
      index: 1,
      outlineLevel: 1,
    });
    expect(parsedSheet.rows[0]!.cells[0]!.content).toEqual({
      cached: { kind: 'number', value: 2 },
      formula: { expression: '1+1', kind: 'normal' },
      kind: 'formula',
    });
    expect(parsedSheet.rows[1]).toMatchObject({ hidden: false, index: 2 });
    expect(parsedSheet.columns).toEqual([
      {
        end: 2,
        hidden: false,
        start: 1,
        style: 0,
        width: 25,
      },
      { end: 3, hidden: true, start: 3, width: 8 },
    ]);

    const mixed = await applyXlsxEdits(snapshot, [
      {
        cell: 'A1',
        content: { kind: 'value', value: { kind: 'number', value: 7 } },
        kind: 'set-cell',
        operationId: 'mixed-cell',
        sheetKey: sheet.key,
      },
      {
        height: 20,
        kind: 'set-row',
        operationId: 'mixed-row',
        row: 1,
        sheetKey: sheet.key,
      },
    ]);
    const mixedSheet = mixed.document.sheets[0]!;
    if (mixedSheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    const sourceSheetBytes = await (
      await JSZip.loadAsync(source)
    )
      .file('xl/worksheets/sheet1.xml')!
      .async('uint8array');
    const cellPatch = patchXlsxWorksheetPartWithReport(
      sourceSheetBytes,
      [
        {
          cell: mixedSheet.rows[0]!.cells[0]!,
          operationId: 'mixed-cell',
        },
      ],
      defaultXlsxWriteLimits(),
      'xl/worksheets/sheet1.xml',
    );
    const propertyPatch = patchXlsxWorksheetProperties(
      cellPatch.data,
      [{ height: 20, kind: 'set-row', operationId: 'mixed-row', row: 1 }],
      defaultXlsxWriteLimits(),
      'xl/worksheets/sheet1.xml',
    );
    const aggregatePatchBytes = cellPatch.patchBytes + propertyPatch.patchBytes;
    const individualPatchBytes = Math.max(
      cellPatch.patchBytes,
      propertyPatch.patchBytes,
    );
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(mixed, {
            limits: { maxPatchBytes: individualPatchBytes },
          }),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: aggregatePatchBytes,
      limit: individualPatchBytes,
      limitName: 'maxPatchBytes',
      part: 'xl/worksheets/sheet1.xml',
    });
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(mixed, { limits: { maxPatchCount: 1 } }),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxPatchCount',
      part: 'xl/worksheets/sheet1.xml',
    });

    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(edited, { limits: { maxPatchCount: 1 } }),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 4,
      limit: 1,
      limitName: 'maxPatchCount',
    });
  });

  it('edits prefixed Strict row and column properties', async () => {
    const strictSheetNs = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictRelNs =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const source = await createIndependentXlsx({
      '[Content_Types].xml': `<Types xmlns="${XLSX_CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelNs}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${strictRelNs}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheetNs}" xmlns:r="${strictRelNs}"><s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets></s:workbook>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheetNs}"><s:cols><s:col min="1" max="1" width="9" customWidth="1"/></s:cols><s:sheetData><s:row r="1"><s:c r="A1"><s:v>1</s:v></s:c></s:row></s:sheetData></s:worksheet>`,
    });
    const snapshot = await readXlsxRoundTrip(source);
    const edited = await applyXlsxEdits(snapshot, [
      {
        height: 18,
        kind: 'set-row',
        operationId: 'strict-row',
        row: 1,
        sheetKey: snapshot.document.sheets[0]!.key,
      },
      {
        end: 1,
        hidden: true,
        kind: 'set-column',
        operationId: 'strict-column',
        sheetKey: snapshot.document.sheets[0]!.key,
        start: 1,
      },
    ]);
    const result = await writeXlsxRoundTrip(edited);
    expect(result.report.level).toBe('R2');
    expect((await readXlsxRoundTrip(result.data)).source.conformance).toBe(
      'strict',
    );
    const parsed = await parseXlsx(result.data, { errorMode: 'strict' });
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows[0]).toMatchObject({ height: 18, index: 1 });
    expect(sheet.columns[0]).toMatchObject({ hidden: true, start: 1 });
  });
});
