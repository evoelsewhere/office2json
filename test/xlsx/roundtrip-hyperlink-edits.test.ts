import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  applyXlsxEdits,
  parseXlsx,
  readXlsxRoundTrip,
  writeXlsxRoundTrip,
  XlsxWriteError,
} from '../../src/formats/xlsx';
import { patchXlsxInternalHyperlinks } from '../../src/formats/xlsx/roundtrip/hyperlink-patch';
import { patchXlsxWorksheetPartWithReport } from '../../src/formats/xlsx/roundtrip/worksheet-patch';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import {
  createIndependentXlsx,
  independentWorksheet,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_NS,
  XLSX_OFFICE_REL_TYPE,
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
  throw new Error('Expected hyperlink edit to fail');
}

async function zipPart(data: Uint8Array, name: string): Promise<Uint8Array> {
  return (await JSZip.loadAsync(data)).file(name)!.async('uint8array');
}

describe('XLSX verified hyperlink edits', () => {
  it('updates, appends, and removes internal targets with R1/R2 evidence', async () => {
    const source = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="B2"><v>42</v></c></row></sheetData><hyperlinks><hyperlink ref="A1" location="Old!A1" display="Display" tooltip="Tooltip"/></hyperlinks></worksheet>`,
    });
    const snapshot = await readXlsxRoundTrip(source);
    const edited = await applyXlsxEdits(snapshot, [
      {
        cell: 'A1',
        kind: 'set-hyperlink',
        operationId: 'update-a1-link',
        sheetKey: snapshot.document.sheets[0]!.key,
        target: { kind: 'internal', location: "'New & Sheet'!B2" },
      },
      {
        cell: 'B2',
        kind: 'set-hyperlink',
        operationId: 'add-b2-link',
        sheetKey: snapshot.document.sheets[0]!.key,
        target: { kind: 'internal', location: 'New!C3' },
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
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.hyperlinks).toEqual([
      expect.objectContaining({
        display: 'Display',
        target: { kind: 'internal', location: "'New & Sheet'!B2" },
        tooltip: 'Tooltip',
      }),
      expect.objectContaining({
        target: { kind: 'internal', location: 'New!C3' },
      }),
    ]);
    expect(sheet.hyperlinks[1]!.range.reference).toBe('B2');

    const mixed = await applyXlsxEdits(snapshot, [
      {
        cell: 'A1',
        content: { kind: 'value', value: { kind: 'number', value: 7 } },
        kind: 'set-cell',
        operationId: 'set-a1',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
      {
        cell: 'B2',
        kind: 'set-hyperlink',
        operationId: 'link-b2',
        sheetKey: snapshot.document.sheets[0]!.key,
        target: { kind: 'internal', location: 'Sheet2!A1' },
      },
    ]);
    const mixedSheet = mixed.document.sheets[0]!;
    if (mixedSheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    const sourceWorksheet = await zipPart(source, 'xl/worksheets/sheet1.xml');
    const cellPatch = patchXlsxWorksheetPartWithReport(
      sourceWorksheet,
      [{ cell: mixedSheet.rows[0]!.cells[0]!, operationId: 'set-a1' }],
      defaultXlsxWriteLimits(),
      'xl/worksheets/sheet1.xml',
    );
    const linkPatch = patchXlsxInternalHyperlinks(
      cellPatch.data,
      [
        {
          cell: 'B2',
          operationId: 'link-b2',
          target: { kind: 'internal', location: 'Sheet2!A1' },
        },
      ],
      defaultXlsxWriteLimits(),
      'xl/worksheets/sheet1.xml',
    );
    const aggregatePatchBytes = cellPatch.patchBytes + linkPatch.patchBytes;
    const individualPatchLimit = Math.max(
      cellPatch.patchBytes,
      linkPatch.patchBytes,
    );
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(mixed, {
            limits: { maxPatchBytes: individualPatchLimit },
          }),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: aggregatePatchBytes,
      limit: individualPatchLimit,
      limitName: 'maxPatchBytes',
    });
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(mixed, { limits: { maxPatchCount: 1 } }),
        )
      ).diagnostic,
    ).toMatchObject({ actual: 2, limit: 1, limitName: 'maxPatchCount' });

    const reparsed = await readXlsxRoundTrip(first.data);
    const removed = await applyXlsxEdits(reparsed, [
      {
        cell: 'A1',
        kind: 'set-hyperlink',
        operationId: 'remove-a1-link',
        sheetKey: reparsed.document.sheets[0]!.key,
        target: null,
      },
      {
        cell: 'B2',
        kind: 'set-hyperlink',
        operationId: 'remove-b2-link',
        sheetKey: reparsed.document.sheets[0]!.key,
        target: null,
      },
    ]);
    const removedResult = await writeXlsxRoundTrip(removed);
    const removedSheet = (await parseXlsx(removedResult.data)).sheets[0]!;
    expect(
      removedSheet.kind === 'worksheet' ? removedSheet.hyperlinks : null,
    ).toEqual([]);
  });

  it('adds an internal target without formula or date-style closure coupling', async () => {
    const formula = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': independentWorksheet(
          '<row r="1"><c r="A1"><f t="array" ref="A1">1+1</f><v>2</v></c></row>',
        ),
      }),
    );
    const edited = await applyXlsxEdits(formula, [
      {
        cell: 'A1',
        kind: 'set-hyperlink',
        operationId: 'link-formula',
        sheetKey: formula.document.sheets[0]!.key,
        target: { kind: 'internal', location: 'Sheet1!B2' },
      },
    ]);
    await expect(writeXlsxRoundTrip(edited)).resolves.toMatchObject({
      report: { level: 'R2' },
    });
    const dateStyles = `<styleSheet xmlns="${XLSX_SPREADSHEET_NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="1"><font/></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>`;
    const date = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/styles.xml': dateStyles,
        'xl/worksheets/sheet1.xml': independentWorksheet(
          '<row r="1"><c r="A1" s="1"><v>2</v></c></row>',
        ),
      }),
    );
    const linkedDate = await applyXlsxEdits(date, [
      {
        cell: 'A1',
        kind: 'set-hyperlink',
        operationId: 'link-date',
        sheetKey: date.document.sheets[0]!.key,
        target: { kind: 'internal', location: 'Sheet1!B2' },
      },
    ]);
    await expect(writeXlsxRoundTrip(linkedDate)).resolves.toMatchObject({
      report: { level: 'R2' },
    });
  });

  it('creates, updates, and removes external hyperlink relationships deterministically', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const external = await applyXlsxEdits(snapshot, [
      {
        cell: 'A1',
        kind: 'set-hyperlink',
        operationId: 'external-link',
        sheetKey: snapshot.document.sheets[0]!.key,
        target: {
          kind: 'external',
          location: 'Section',
          url: 'https://example.invalid/',
        },
      },
    ]);
    const externalResult = await writeXlsxRoundTrip(external);
    expect(externalResult.report.level).toBe('R2');
    await expect(
      writeXlsxRoundTrip(external, {
        limits: { maxDependencyEdges: 2, maxDirtyParts: 2 },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(external, {
            limits: { maxDependencyEdges: 1 },
          }),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxDependencyEdges',
    });
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(external, {
            limits: { maxDirtyParts: 1 },
          }),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxDirtyParts',
    });
    expect(
      externalResult.report.parts.find(
        (part) => part.name === 'xl/worksheets/_rels/sheet1.xml.rels',
      )?.disposition,
    ).toBe('add');
    const createdSheet = (await parseXlsx(externalResult.data)).sheets[0]!;
    expect(
      createdSheet.kind === 'worksheet'
        ? createdSheet.hyperlinks[0]!.target
        : null,
    ).toEqual({
      kind: 'external',
      location: 'Section',
      url: 'https://example.invalid/',
    });

    const source = await createIndependentXlsx({
      '[Content_Types].xml': `<Types xmlns="${XLSX_CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="sheet2" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="styles" Type="${XLSX_OFFICE_REL_TYPE}styles" Target="styles.xml"/><Relationship Id="strings" Type="${XLSX_OFFICE_REL_TYPE}sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
      'xl/workbook.xml': `<workbook xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheets><sheet name="First" sheetId="1" r:id="sheet1"/><sheet name="Second" sheetId="2" r:id="sheet2"/></sheets></workbook>`,
      'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="link" Type="${XLSX_OFFICE_REL_TYPE}hyperlink" Target="https://old-first.invalid/" TargetMode="External"/></Relationships>`,
      'xl/worksheets/_rels/sheet2.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="link" Type="${XLSX_OFFICE_REL_TYPE}hyperlink" Target="https://example.invalid/" TargetMode="External"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><hyperlinks><hyperlink ref="A1" r:id="link"/></hyperlinks></worksheet>`,
      'xl/worksheets/sheet2.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="B2"><v>2</v></c></row></sheetData><hyperlinks><hyperlink ref="B2" location="Internal!A1"/><hyperlink ref="A1" r:id="link"/></hyperlinks></worksheet>`,
    });
    const existing = await readXlsxRoundTrip(source);
    const ownerScoped = await applyXlsxEdits(existing, [
      {
        cell: 'A1',
        kind: 'set-hyperlink',
        operationId: 'owner-scoped-external',
        sheetKey: existing.document.sheets[0]!.key,
        target: { kind: 'external', url: 'https://example.invalid/' },
      },
    ]);
    const ownerScopedResult = await writeXlsxRoundTrip(ownerScoped);
    expect(
      new TextDecoder().decode(
        await zipPart(
          ownerScopedResult.data,
          'xl/worksheets/_rels/sheet1.xml.rels',
        ),
      ),
    ).toContain('Target="https://example.invalid/"');
    const updated = await applyXlsxEdits(existing, [
      {
        cell: 'A1',
        kind: 'set-hyperlink',
        operationId: 'update-external',
        sheetKey: existing.document.sheets[1]!.key,
        target: {
          kind: 'external',
          location: 'Updated',
          url: 'mailto:updated@example.invalid',
        },
      },
    ]);
    const updatedResult = await writeXlsxRoundTrip(updated);
    const updatedRelationships = new TextDecoder().decode(
      await zipPart(updatedResult.data, 'xl/worksheets/_rels/sheet2.xml.rels'),
    );
    expect(updatedRelationships).toContain('Id="link"');
    expect(updatedRelationships).toContain(
      'Target="mailto:updated@example.invalid"',
    );
    const updatedSnapshot = await readXlsxRoundTrip(updatedResult.data);
    const removed = await applyXlsxEdits(updatedSnapshot, [
      {
        cell: 'A1',
        kind: 'set-hyperlink',
        operationId: 'remove-external',
        sheetKey: updatedSnapshot.document.sheets[1]!.key,
        target: null,
      },
    ]);
    const removedResult = await writeXlsxRoundTrip(removed);
    const removedRelationships = new TextDecoder().decode(
      await zipPart(removedResult.data, 'xl/worksheets/_rels/sheet2.xml.rels'),
    );
    expect(removedRelationships).not.toContain('Id="link"');
    expect(removedRelationships).toContain('<Relationships');
    const removedSheet = (await parseXlsx(removedResult.data)).sheets[1]!;
    expect(
      removedSheet.kind === 'worksheet' ? removedSheet.hyperlinks : null,
    ).toEqual([
      expect.objectContaining({
        target: { kind: 'internal', location: 'Internal!A1' },
      }),
    ]);
  });

  it('creates a Strict external hyperlink relationship with R2 evidence', async () => {
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
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheetNs}"><s:sheetData><s:row r="1"><s:c r="A1"><s:v>1</s:v></s:c></s:row></s:sheetData></s:worksheet>`,
    });
    const snapshot = await readXlsxRoundTrip(source);
    const edited = await applyXlsxEdits(snapshot, [
      {
        cell: 'A1',
        kind: 'set-hyperlink',
        operationId: 'strict-external-link',
        sheetKey: snapshot.document.sheets[0]!.key,
        target: { kind: 'external', url: 'https://strict.invalid/' },
      },
    ]);
    const result = await writeXlsxRoundTrip(edited);
    expect(result.report.level).toBe('R2');
    expect((await readXlsxRoundTrip(result.data)).source.conformance).toBe(
      'strict',
    );
    expect(
      new TextDecoder().decode(
        await zipPart(result.data, 'xl/worksheets/_rels/sheet1.xml.rels'),
      ),
    ).toContain(`Type="${strictRelNs}/hyperlink"`);
    const parsed = await parseXlsx(result.data, { errorMode: 'strict' });
    const sheet = parsed.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.hyperlinks[0]!.target : null,
    ).toEqual({ kind: 'external', url: 'https://strict.invalid/' });
  });
});
