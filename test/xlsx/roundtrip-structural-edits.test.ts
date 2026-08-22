import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';

import {
  applyXlsxEdits,
  parseXlsx,
  readXlsxRoundTrip,
  writeXlsxRoundTrip,
} from '../../src/formats/xlsx';
import { patchXlsxTableStructure } from '../../src/formats/xlsx/roundtrip/table-structure-patch';
import {
  patchXlsxCommentAnchors,
  patchXlsxCommentVmlAnchors,
} from '../../src/formats/xlsx/roundtrip/comment-structure-patch';
import { patchXlsxDrawingStructure } from '../../src/formats/xlsx/roundtrip/drawing-structure-patch';
import { patchXlsxChartStructure } from '../../src/formats/xlsx/roundtrip/chart-structure-patch';
import { patchXlsxSparklineStructure } from '../../src/formats/xlsx/roundtrip/sparkline-structure-patch';
import { patchXlsxWorksheetStructure } from '../../src/formats/xlsx/roundtrip/worksheet-structure-patch';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import {
  createIndependentXlsx,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

function portable<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

describe('XLSX verified structural row and column edits', () => {
  it('inserts and deletes authored cells in ordered atomic batches', async () => {
    const source = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="3" hidden="1"><c r="C3" s="0"><v>3</v></c></row></sheetData></worksheet>`,
    });
    const snapshot = await readXlsxRoundTrip(source);
    const sheetKey = snapshot.document.sheets[0]!.key;
    const edited = await applyXlsxEdits(snapshot, [
      {
        count: 2,
        index: 2,
        kind: 'insert-rows',
        operationId: 'insert-rows',
        sheetKey,
      },
      {
        count: 1,
        index: 1,
        kind: 'delete-rows',
        operationId: 'delete-rows',
        sheetKey,
      },
      {
        count: 1,
        index: 2,
        kind: 'insert-columns',
        operationId: 'insert-columns',
        sheetKey,
      },
      {
        count: 1,
        index: 1,
        kind: 'delete-columns',
        operationId: 'delete-columns',
        sheetKey,
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
    expect(sheet.rows).toEqual([
      expect.objectContaining({
        cells: [
          expect.objectContaining({
            address: 'C4',
            column: 3,
            content: { kind: 'value', value: { kind: 'number', value: 3 } },
            style: 0,
          }),
        ],
        hidden: true,
        index: 4,
      }),
    ]);
  });

  it('shifts a prefixed Strict worksheet without producer software', async () => {
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
        count: 1,
        index: 1,
        kind: 'insert-rows',
        operationId: 'strict-insert-row',
        sheetKey: snapshot.document.sheets[0]!.key,
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
    expect(sheet.rows[0]).toMatchObject({
      cells: [expect.objectContaining({ address: 'A2' })],
      index: 2,
    });
  });

  it('keeps declared dimensions and merged ranges aligned', async () => {
    const source = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><dimension ref="A1:B2"/><sheetViews><sheetView workbookViewId="0" topLeftCell="A2"><selection activeCell="B2" activeCellId="1" sqref="A1:A2 B2"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row></sheetData><protectedRanges><protectedRange name="Input" sqref="A1:B2"/></protectedRanges><autoFilter ref="A1:B2"><sortState ref="A1:B2"><sortCondition ref="A1:A2"/></sortState></autoFilter><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><conditionalFormatting sqref="A1:B2"><cfRule type="top10" priority="1" rank="1"/></conditionalFormatting><dataValidations count="1" disablePrompts="1"><dataValidation sqref="A1:B2"/></dataValidations><hyperlinks><hyperlink ref="A1:B1" location="Sheet1!A1"/></hyperlinks><rowBreaks count="1" manualBreakCount="1"><brk id="2" min="0" max="1" man="1"/></rowBreaks><colBreaks count="1" manualBreakCount="0"><brk id="2" min="0" max="1" pt="1"/></colBreaks></worksheet>`,
    });
    const snapshot = await readXlsxRoundTrip(source);
    const edited = await applyXlsxEdits(snapshot, [
      {
        count: 1,
        index: 1,
        kind: 'insert-rows',
        operationId: 'insert-layout-row',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    const result = await writeXlsxRoundTrip(edited);
    const parsed = await parseXlsx(result.data, { errorMode: 'strict' });
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.declaredDimension?.reference).toBe('A2:B3');
    expect(sheet.mergedRanges.map((range) => range.reference)).toEqual([
      'A2:B2',
    ]);
    expect(sheet.hyperlinks.map((link) => link.range.reference)).toEqual([
      'A2:B2',
    ]);
    expect(sheet.autoFilter?.range.reference).toBe('A2:B3');
    expect(sheet.autoFilter?.sort?.range.reference).toBe('A2:B3');
    expect(sheet.autoFilter?.sort?.conditions[0]?.range.reference).toBe(
      'A2:A3',
    );
    expect(sheet.dataValidations[0]?.ranges[0]?.reference).toBe('A2:B3');
    expect(sheet.dataValidationSettings).toEqual({ disablePrompts: true });
    expect(sheet.conditionalFormattings[0]?.ranges[0]?.reference).toBe('A2:B3');
    expect(sheet.protectedRanges[0]?.ranges[0]?.reference).toBe('A2:B3');
    expect(sheet.print?.rowBreaks?.[0]).toMatchObject({
      end: 1,
      position: 3,
      start: 0,
    });
    expect(sheet.print?.columnBreaks?.[0]).toMatchObject({
      end: 2,
      position: 2,
      start: 1,
    });
    expect(sheet.views[0]?.topLeftCell).toBe('A3');
    expect(sheet.views[0]?.selections[0]).toMatchObject({
      activeCell: 'B3',
      activeCellId: 1,
      ranges: [{ reference: 'A2:A3' }, { reference: 'B3' }],
    });
    expect(result.report.level).toBe('R2');
  });

  it('keeps worksheet tables and their owned parts aligned', async () => {
    const tablePart = 'xl/tables/table1.xml';
    const contentTypes = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/${tablePart}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`;
    const generatedSource = await createIndependentXlsx({
      '[Content_Types].xml': contentTypes,
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${XLSX_OFFICE_REL_TYPE}officeDocument" Target="xl/workbook.xml"/><Relationship Id="unowned-table" Type="${XLSX_OFFICE_REL_TYPE}table" Target="xl/tables/table2.xml"/></Relationships>`,
      'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="table" Type="${XLSX_OFFICE_REL_TYPE}table" Target="../tables/table1.xml"/><Relationship Id="external-table" Type="${XLSX_OFFICE_REL_TYPE}table" Target="https://example.invalid/table.xml" TargetMode="External"/><Relationship Id="internal-link" Type="${XLSX_OFFICE_REL_TYPE}hyperlink" Target="../tables/table2.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheetData><row r="1"><c r="A1" t="str"><v>A</v></c><c r="B1" t="str"><v>B</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>4</v></c></row></sheetData><tableParts count="1"><tablePart r:id="table"/></tableParts></worksheet>`,
      [tablePart]: `<table xmlns="${XLSX_SPREADSHEET_NS}" id="1" name="Table1" displayName="Table1" ref="A1:B3"><autoFilter ref="A1:B3"><sortState ref="A1:B3"><sortCondition ref="A2:A3"/></sortState></autoFilter><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>`,
      'xl/tables/table2.xml': `<table xmlns="${XLSX_SPREADSHEET_NS}" id="2" name="Unused" displayName="Unused" ref="D10:E12"><tableColumns count="2"><tableColumn id="1" name="D"/><tableColumn id="2" name="E"/></tableColumns></table>`,
    });
    const datedSource = await JSZip.loadAsync(generatedSource);
    const tableBytes = await datedSource.file(tablePart)!.async('uint8array');
    datedSource.file(tablePart, tableBytes, {
      date: new Date('2001-02-03T04:05:06.000Z'),
    });
    const source = await datedSource.generateAsync({ type: 'uint8array' });
    const snapshot = await readXlsxRoundTrip(source);
    const structuralOperation = {
      count: 1,
      index: 2,
      kind: 'insert-rows' as const,
      operationId: 'insert-table-row',
      sheetKey: snapshot.document.sheets[0]!.key,
    };
    const edited = await applyXlsxEdits(snapshot, [structuralOperation]);
    const first = await writeXlsxRoundTrip(edited);
    const second = await writeXlsxRoundTrip(portable(edited));
    expect(second.data).toEqual(first.data);
    expect(
      first.report.parts
        .filter((part) => part.disposition === 'patch')
        .map((part) => part.name),
    ).toEqual(['xl/tables/table1.xml', 'xl/worksheets/sheet1.xml']);
    const parsed = await parseXlsx(first.data, { errorMode: 'strict' });
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.tables[0]?.range.reference).toBe('A1:B4');
    expect(sheet.tables[0]?.autoFilter?.range.reference).toBe('A1:B4');
    expect(
      sheet.tables[0]?.autoFilter?.sort?.conditions[0]?.range.reference,
    ).toBe('A3:A4');
    expect(first.report.level).toBe('R2');

    const sourceZip = await JSZip.loadAsync(source);
    const outputZip = await JSZip.loadAsync(first.data);
    expect(outputZip.file(tablePart)!.date).toEqual(
      sourceZip.file(tablePart)!.date,
    );
    const request = {
      count: structuralOperation.count,
      index: structuralOperation.index,
      kind: structuralOperation.kind,
      operationId: structuralOperation.operationId,
    };
    const worksheetPatch = patchXlsxWorksheetStructure(
      await sourceZip.file('xl/worksheets/sheet1.xml')!.async('uint8array'),
      [request],
      defaultXlsxWriteLimits(),
      'xl/worksheets/sheet1.xml',
    );
    const tablePatch = patchXlsxTableStructure(
      await sourceZip.file(tablePart)!.async('uint8array'),
      [request],
      defaultXlsxWriteLimits(),
      tablePart,
    );
    const patchBytes = worksheetPatch.patchBytes + tablePatch.patchBytes;
    const patchCount = worksheetPatch.patchCount + tablePatch.patchCount;
    const generatedXmlBytes = first.report.parts
      .filter((part) => part.disposition === 'patch')
      .reduce((total, part) => total + part.byteLength, 0);
    await expect(
      writeXlsxRoundTrip(edited, {
        limits: {
          maxDependencyEdges: 2,
          maxDirtyParts: 2,
          maxGeneratedXmlBytes: generatedXmlBytes,
          maxPatchBytes: patchBytes,
          maxPatchCount: patchCount,
          maxPatchedParts: 2,
        },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    for (const [limitName, limit] of [
      ['maxDependencyEdges', 1],
      ['maxDirtyParts', 1],
      ['maxGeneratedXmlBytes', generatedXmlBytes - 1],
      ['maxPatchBytes', patchBytes - 1],
      ['maxPatchCount', patchCount - 1],
      ['maxPatchedParts', 1],
    ] as const) {
      await expect(
        writeXlsxRoundTrip(edited, { limits: { [limitName]: limit } }),
      ).rejects.toMatchObject({
        diagnostic: { code: 'resource-limit-exceeded', limitName },
      });
    }

    const rowEdit = await applyXlsxEdits(snapshot, [
      {
        hidden: true,
        kind: 'set-row',
        operationId: 'table-row-property',
        row: 1,
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    await expect(writeXlsxRoundTrip(rowEdit)).rejects.toMatchObject({
      diagnostic: { featureClass: 'unsupported-part', part: tablePart },
    });
    const outside = await applyXlsxEdits(snapshot, [
      {
        count: 1,
        index: 5,
        kind: 'insert-rows',
        operationId: 'outside-table',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    const outsideResult = await writeXlsxRoundTrip(outside);
    expect(
      outsideResult.report.parts.find((part) => part.name === tablePart)
        ?.disposition,
    ).toBe('copy');
    for (const operation of [
      {
        count: 1,
        index: 2,
        kind: 'delete-rows' as const,
        operationId: 'delete-table-data-row',
      },
      {
        count: 1,
        index: 3,
        kind: 'insert-columns' as const,
        operationId: 'insert-outside-table-column',
      },
      {
        count: 1,
        index: 3,
        kind: 'delete-columns' as const,
        operationId: 'delete-outside-table-column',
      },
    ]) {
      const candidate = await applyXlsxEdits(snapshot, [
        {
          ...operation,
          sheetKey: snapshot.document.sheets[0]!.key,
        },
      ]);
      await expect(writeXlsxRoundTrip(candidate)).resolves.toMatchObject({
        report: { level: 'R2' },
      });
    }
  });

  it('keeps legacy, threaded, and VML comment anchors aligned', async () => {
    const threadedRelationship =
      'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
    const personRelationship =
      'http://schemas.microsoft.com/office/2017/10/relationships/person';
    const threadedNamespace =
      'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
    const contentTypes = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/><Override PartName="/xl/comments-unused.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/><Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/><Override PartName="/xl/threadedComments/threadedUnused.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/><Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/></Types>`;
    const generatedCommentSource = await createIndependentXlsx({
      '[Content_Types].xml': contentTypes,
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${XLSX_OFFICE_REL_TYPE}officeDocument" Target="xl/workbook.xml"/><Relationship Id="unused-comments" Type="${XLSX_OFFICE_REL_TYPE}comments" Target="xl/comments-unused.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="rIdSheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="styles" Type="${XLSX_OFFICE_REL_TYPE}styles" Target="styles.xml"/><Relationship Id="strings" Type="${XLSX_OFFICE_REL_TYPE}sharedStrings" Target="sharedStrings.xml"/><Relationship Id="persons" Type="${personRelationship}" Target="persons/person.xml"/></Relationships>`,
      'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><t>One</t></text></comment><comment ref="B2" authorId="0"><text><t>Two</t></text></comment></commentList></comments>`,
      'xl/comments-unused.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>U</author></authors><commentList><comment ref="Z9" authorId="0"><text><t>Unused</t></text></comment></commentList></comments>`,
      'xl/drawings/vmlDrawing1.vml':
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape><v:shape><x:ClientData ObjectType="Note"><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape></xml>',
      'xl/persons/person.xml': `<personList xmlns="${threadedNamespace}"><person displayName="P" id="person"/></personList>`,
      'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${threadedNamespace}"><threadedComment ref="C3" dT="2024-01-01T00:00:00Z" personId="person" id="root"><text>Root</text></threadedComment><threadedComment ref="C3" dT="2024-01-01T00:01:00Z" personId="person" id="reply" parentId="root"><text>Reply</text></threadedComment></ThreadedComments>`,
      'xl/threadedComments/threadedUnused.xml': `<ThreadedComments xmlns="${threadedNamespace}"><threadedComment ref="Z9" dT="2024-01-01T00:00:00Z" personId="person" id="unused"><text>Unused</text></threadedComment></ThreadedComments>`,
      'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="comments" Type="${XLSX_OFFICE_REL_TYPE}comments" Target="../comments1.xml"/><Relationship Id="vml" Type="${XLSX_OFFICE_REL_TYPE}vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="threaded" Type="${threadedRelationship}" Target="../threadedComments/threadedComment1.xml"/><Relationship Id="internal-link" Type="${XLSX_OFFICE_REL_TYPE}hyperlink" Target="../threadedComments/threadedUnused.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="B2"><v>2</v></c></row><row r="3"><c r="C3"><v>3</v></c></row></sheetData><legacyDrawing r:id="vml"/></worksheet>`,
    });
    const commentSourceZip = await JSZip.loadAsync(generatedCommentSource);
    for (const [index, part] of [
      'xl/comments1.xml',
      'xl/drawings/vmlDrawing1.vml',
      'xl/threadedComments/threadedComment1.xml',
    ].entries()) {
      const data = await commentSourceZip.file(part)!.async('uint8array');
      commentSourceZip.file(part, data, {
        date: new Date(
          `2002-03-04T05:06:${String(index * 2).padStart(2, '0')}.000Z`,
        ),
      });
    }
    const source = await commentSourceZip.generateAsync({ type: 'uint8array' });
    const snapshot = await readXlsxRoundTrip(source);
    const edited = await applyXlsxEdits(snapshot, [
      {
        count: 1,
        index: 2,
        kind: 'insert-rows',
        operationId: 'comment-rows',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
      {
        count: 1,
        index: 2,
        kind: 'insert-columns',
        operationId: 'comment-columns',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    const result = await writeXlsxRoundTrip(edited);
    expect(result.report.level).toBe('R2');
    expect(
      result.report.parts
        .filter((part) => part.disposition === 'patch')
        .map((part) => part.name),
    ).toEqual([
      'xl/comments1.xml',
      'xl/drawings/vmlDrawing1.vml',
      'xl/threadedComments/threadedComment1.xml',
      'xl/worksheets/sheet1.xml',
    ]);
    const parsed = await parseXlsx(result.data, { errorMode: 'strict' });
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.comments.map((comment) => comment.reference)).toEqual([
      'A1',
      'C3',
      'D4',
      'D4',
    ]);
    expect(
      result.report.parts.find((part) => part.name === 'xl/persons/person.xml')
        ?.disposition,
    ).toBe('copy');
    const sourceZip = await JSZip.loadAsync(source);
    const outputZip = await JSZip.loadAsync(result.data);
    const commentParts = [
      'xl/comments1.xml',
      'xl/drawings/vmlDrawing1.vml',
      'xl/threadedComments/threadedComment1.xml',
    ];
    for (const part of commentParts) {
      expect(outputZip.file(part)!.date).toEqual(sourceZip.file(part)!.date);
    }
    const requests = [
      {
        count: 1,
        index: 2,
        kind: 'insert-rows' as const,
        operationId: 'comment-rows',
      },
      {
        count: 1,
        index: 2,
        kind: 'insert-columns' as const,
        operationId: 'comment-columns',
      },
    ];
    const worksheetPatch = patchXlsxWorksheetStructure(
      await sourceZip.file('xl/worksheets/sheet1.xml')!.async('uint8array'),
      requests,
      defaultXlsxWriteLimits(),
      'xl/worksheets/sheet1.xml',
    );
    const legacyPatch = patchXlsxCommentAnchors(
      await sourceZip.file('xl/comments1.xml')!.async('uint8array'),
      requests,
      defaultXlsxWriteLimits(),
      'xl/comments1.xml',
    );
    const threadedPatch = patchXlsxCommentAnchors(
      await sourceZip
        .file('xl/threadedComments/threadedComment1.xml')!
        .async('uint8array'),
      requests,
      defaultXlsxWriteLimits(),
      'xl/threadedComments/threadedComment1.xml',
    );
    const vmlPatch = patchXlsxCommentVmlAnchors(
      await sourceZip.file('xl/drawings/vmlDrawing1.vml')!.async('uint8array'),
      requests,
      defaultXlsxWriteLimits(),
      'xl/drawings/vmlDrawing1.vml',
    );
    const patchBytes =
      worksheetPatch.patchBytes +
      legacyPatch.patchBytes +
      threadedPatch.patchBytes +
      vmlPatch.patchBytes;
    const patchCount =
      worksheetPatch.patchCount +
      legacyPatch.patchCount +
      threadedPatch.patchCount +
      vmlPatch.patchCount;
    const generatedXmlBytes = result.report.parts
      .filter((part) => part.disposition === 'patch')
      .reduce((total, part) => total + part.byteLength, 0);
    await expect(
      writeXlsxRoundTrip(edited, {
        limits: {
          maxDependencyEdges: 5,
          maxDirtyParts: 4,
          maxGeneratedXmlBytes: generatedXmlBytes,
          maxPatchBytes: patchBytes,
          maxPatchCount: patchCount,
          maxPatchedParts: 4,
        },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    for (const [limitName, limit] of [
      ['maxDependencyEdges', 4],
      ['maxDirtyParts', 3],
      ['maxGeneratedXmlBytes', generatedXmlBytes - 1],
      ['maxPatchBytes', patchBytes - 1],
      ['maxPatchCount', patchCount - 1],
      ['maxPatchedParts', 3],
    ] as const) {
      await expect(
        writeXlsxRoundTrip(edited, { limits: { [limitName]: limit } }),
      ).rejects.toMatchObject({
        diagnostic: { code: 'resource-limit-exceeded', limitName },
      });
    }
    const outside = await applyXlsxEdits(snapshot, [
      {
        count: 1,
        index: 10,
        kind: 'insert-rows',
        operationId: 'outside-comments',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    const outsideResult = await writeXlsxRoundTrip(outside);
    for (const part of commentParts) {
      expect(
        outsideResult.report.parts.find((candidate) => candidate.name === part)
          ?.disposition,
      ).toBe('copy');
    }
    const rowEdit = await applyXlsxEdits(snapshot, [
      {
        hidden: true,
        kind: 'set-row',
        operationId: 'comment-row-property',
        row: 1,
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    await expect(writeXlsxRoundTrip(rowEdit)).rejects.toMatchObject({
      diagnostic: { featureClass: 'opaque-content' },
    });
  });

  it('keeps drawing anchors and shape placements aligned', async () => {
    const drawingNamespace =
      'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
    const drawingMainNamespace =
      'http://schemas.openxmlformats.org/drawingml/2006/main';
    const contentTypes = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`;
    const marker = (name: 'from' | 'to', column: number, row: number) =>
      `<xdr:${name}><xdr:col>${column}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${name}>`;
    const shape = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="Shape 1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12700" cy="12700"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>`;
    const generatedSource = await createIndependentXlsx({
      '[Content_Types].xml': contentTypes,
      'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="${drawingNamespace}" xmlns:a="${drawingMainNamespace}"><xdr:oneCellAnchor>${marker('from', 1, 1)}<xdr:ext cx="12700" cy="12700"/>${shape}<xdr:clientData/></xdr:oneCellAnchor><xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 2, 2)}${shape.replace('id="1"', 'id="2"').replace('Shape 1', 'Shape 2')}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
      'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="drawing" Type="${XLSX_OFFICE_REL_TYPE}drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="B2"><v>2</v></c></row></sheetData><drawing r:id="drawing"/></worksheet>`,
    });
    const sourceZip = await JSZip.loadAsync(generatedSource);
    const drawingPart = 'xl/drawings/drawing1.xml';
    const drawingBytes = await sourceZip.file(drawingPart)!.async('uint8array');
    sourceZip.file(drawingPart, drawingBytes, {
      date: new Date('2003-04-05T06:07:08.000Z'),
    });
    const source = await sourceZip.generateAsync({ type: 'uint8array' });
    const snapshot = await readXlsxRoundTrip(source);
    const operations = [
      {
        count: 1,
        index: 2,
        kind: 'insert-rows' as const,
        operationId: 'drawing-rows',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
      {
        count: 1,
        index: 2,
        kind: 'insert-columns' as const,
        operationId: 'drawing-columns',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ];
    const edited = await applyXlsxEdits(snapshot, operations);
    const result = await writeXlsxRoundTrip(edited);
    expect(result.report.level).toBe('R2');
    expect(
      result.report.parts
        .filter((part) => part.disposition === 'patch')
        .map((part) => part.name),
    ).toEqual([drawingPart, 'xl/worksheets/sheet1.xml']);
    const parsed = await parseXlsx(result.data, { errorMode: 'strict' });
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.drawings[0]?.from).toMatchObject({ column: 3, row: 3 });
    expect(sheet.drawings[1]).toMatchObject({
      from: { column: 1, row: 1 },
      to: { column: 4, row: 4 },
    });
    expect(
      (await JSZip.loadAsync(result.data)).file(drawingPart)!.date,
    ).toEqual((await JSZip.loadAsync(source)).file(drawingPart)!.date);
    const requests = operations.map(({ count, index, kind, operationId }) => ({
      count,
      index,
      kind,
      operationId,
    }));
    const originalZip = await JSZip.loadAsync(source);
    const worksheetPatch = patchXlsxWorksheetStructure(
      await originalZip.file('xl/worksheets/sheet1.xml')!.async('uint8array'),
      requests,
      defaultXlsxWriteLimits(),
      'xl/worksheets/sheet1.xml',
    );
    const drawingPatch = patchXlsxDrawingStructure(
      await originalZip.file(drawingPart)!.async('uint8array'),
      requests,
      defaultXlsxWriteLimits(),
      drawingPart,
    );
    const patchBytes = worksheetPatch.patchBytes + drawingPatch.patchBytes;
    const patchCount = worksheetPatch.patchCount + drawingPatch.patchCount;
    const generatedXmlBytes = result.report.parts
      .filter((part) => part.disposition === 'patch')
      .reduce((total, part) => total + part.byteLength, 0);
    await expect(
      writeXlsxRoundTrip(edited, {
        limits: {
          maxDependencyEdges: 3,
          maxDirtyParts: 2,
          maxGeneratedXmlBytes: generatedXmlBytes,
          maxPatchBytes: patchBytes,
          maxPatchCount: patchCount,
          maxPatchedParts: 2,
        },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    for (const [limitName, limit] of [
      ['maxDependencyEdges', 2],
      ['maxDirtyParts', 1],
      ['maxGeneratedXmlBytes', generatedXmlBytes - 1],
      ['maxPatchBytes', patchBytes - 1],
      ['maxPatchCount', patchCount - 1],
      ['maxPatchedParts', 1],
    ] as const) {
      await expect(
        writeXlsxRoundTrip(edited, { limits: { [limitName]: limit } }),
      ).rejects.toMatchObject({
        diagnostic: { code: 'resource-limit-exceeded', limitName },
      });
    }
    const outside = await applyXlsxEdits(snapshot, [
      {
        count: 1,
        index: 10,
        kind: 'insert-rows',
        operationId: 'outside-drawing',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    expect(
      (await writeXlsxRoundTrip(outside)).report.parts.find(
        (part) => part.name === drawingPart,
      )?.disposition,
    ).toBe('copy');
    const rowEdit = await applyXlsxEdits(snapshot, [
      {
        hidden: true,
        kind: 'set-row',
        operationId: 'drawing-row-property',
        row: 1,
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    await expect(writeXlsxRoundTrip(rowEdit)).rejects.toMatchObject({
      diagnostic: { featureClass: 'unsupported-part', part: drawingPart },
    });
  });

  it('keeps supported sparkline sources and locations aligned', async () => {
    const x14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
    const xm = 'http://schemas.microsoft.com/office/excel/2006/main';
    const sparklineUri = '{05c60535-1f16-4fd2-b633-f4f36f0b64e0}';
    const source = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row></sheetData><extLst><ext uri="${sparklineUri}"><x14:sparklineGroups xmlns:x14="${x14}"><x14:sparklineGroup><x14:sparklines xmlns:xm="${xm}"><x14:sparkline><xm:f>Sheet1!$A$1:$A$3</xm:f><xm:sqref>B1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst></worksheet>`,
    });
    const snapshot = await readXlsxRoundTrip(source);
    const sheetKey = snapshot.document.sheets[0]!.key;
    const operations = [
      {
        count: 1,
        index: 2,
        kind: 'insert-rows' as const,
        operationId: 'sparkline-rows',
        sheetKey,
      },
      {
        count: 1,
        index: 2,
        kind: 'insert-columns' as const,
        operationId: 'sparkline-columns',
        sheetKey,
      },
    ];
    const edited = await applyXlsxEdits(snapshot, operations);
    const previewSheet = edited.document.sheets[0]!;
    expect(previewSheet.kind).toBe('worksheet');
    if (previewSheet.kind !== 'worksheet') {
      throw new Error('Expected worksheet');
    }
    expect(previewSheet.sparklineGroups?.[0]?.sparklines[0]).toMatchObject({
      dataFormula: 'Sheet1!$A$1:$A$4',
      location: 'C1',
    });
    const result = await writeXlsxRoundTrip(portable(edited));
    expect(result.report.level).toBe('R2');
    expect(
      result.report.parts
        .filter((part) => part.disposition === 'patch')
        .map((part) => part.name),
    ).toEqual(['xl/worksheets/sheet1.xml']);
    const parsed = await parseXlsx(result.data, { errorMode: 'strict' });
    const outputSheet = parsed.sheets[0]!;
    expect(outputSheet.kind).toBe('worksheet');
    if (outputSheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(outputSheet.sparklineGroups?.[0]?.sparklines[0]).toMatchObject({
      dataFormula: 'Sheet1!$A$1:$A$4',
      location: 'C1',
    });
    const repeated = await writeXlsxRoundTrip(portable(edited));
    expect(repeated.data).toEqual(result.data);
    expect(repeated.report).toEqual(result.report);
    const requests = operations.map(({ count, index, kind, operationId }) => ({
      count,
      index,
      kind,
      operationId,
    }));
    const sourceWorksheet = await (
      await JSZip.loadAsync(source)
    )
      .file('xl/worksheets/sheet1.xml')!
      .async('uint8array');
    const worksheetPatch = patchXlsxWorksheetStructure(
      sourceWorksheet,
      requests,
      defaultXlsxWriteLimits(),
      'xl/worksheets/sheet1.xml',
    );
    const sparklinePatch = patchXlsxSparklineStructure(
      worksheetPatch.data,
      requests,
      defaultXlsxWriteLimits(),
      'xl/worksheets/sheet1.xml',
      'Sheet1',
    );
    const patchBytes = worksheetPatch.patchBytes + sparklinePatch.patchBytes;
    const patchCount = worksheetPatch.patchCount + sparklinePatch.patchCount;
    await expect(
      writeXlsxRoundTrip(portable(edited), {
        limits: { maxPatchBytes: patchBytes, maxPatchCount: patchCount },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    for (const [limitName, limit] of [
      ['maxPatchBytes', patchBytes - 1],
      ['maxPatchCount', patchCount - 1],
    ] as const) {
      await expect(
        writeXlsxRoundTrip(portable(edited), {
          limits: { [limitName]: limit },
        }),
      ).rejects.toMatchObject({
        diagnostic: { code: 'resource-limit-exceeded', limitName },
      });
    }
    await expect(
      applyXlsxEdits(snapshot, [
        {
          count: 3,
          index: 1,
          kind: 'delete-rows',
          operationId: 'delete-sparkline-source',
          sheetKey,
        },
      ]),
    ).rejects.toMatchObject({
      diagnostic: { featureClass: 'sparkline-source-deletion' },
    });
  });

  it('keeps supported chart sources, caches, and anchors aligned', async () => {
    const drawingNamespace =
      'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
    const drawingMainNamespace =
      'http://schemas.openxmlformats.org/drawingml/2006/main';
    const chartNamespace =
      'http://schemas.openxmlformats.org/drawingml/2006/chart';
    const contentTypes = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`;
    const marker = `<xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>`;
    const drawing = `<xdr:wsDr xmlns:xdr="${drawingNamespace}" xmlns:a="${drawingMainNamespace}" xmlns:c="${chartNamespace}" xmlns:r="${XLSX_OFFICE_REL_NS}"><xdr:oneCellAnchor>${marker}<xdr:ext cx="12700" cy="12700"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="12700" cy="12700"/></xdr:xfrm><a:graphic><a:graphicData uri="${chartNamespace}"><c:chart r:id="chart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;
    const stringSource = `<c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef>`;
    const numberSource = `<c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>0</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef>`;
    const chart = `<c:chartSpace xmlns:c="${chartNamespace}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat>${stringSource}</c:cat><c:val>${numberSource}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling/><c:delete val="0"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling/><c:delete val="0"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;
    const generatedSource = await createIndependentXlsx({
      '[Content_Types].xml': contentTypes,
      'xl/charts/chart1.xml': chart,
      'xl/drawings/_rels/drawing1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="chart" Type="${XLSX_OFFICE_REL_TYPE}chart" Target="../charts/chart1.xml"/></Relationships>`,
      'xl/drawings/drawing1.xml': drawing,
      'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="drawing" Type="${XLSX_OFFICE_REL_TYPE}drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheetData><row r="1"><c r="A1"><v>0</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>2</v></c></row></sheetData><drawing r:id="drawing"/></worksheet>`,
    });
    const datedSource = await JSZip.loadAsync(generatedSource);
    const chartDate = new Date('2004-05-06T07:08:10.000Z');
    datedSource.file(
      'xl/charts/chart1.xml',
      await datedSource.file('xl/charts/chart1.xml')!.async('uint8array'),
      { date: chartDate },
    );
    const source = await datedSource.generateAsync({ type: 'uint8array' });
    const snapshot = await readXlsxRoundTrip(source);
    const sheetKey = snapshot.document.sheets[0]!.key;
    const operation = {
      count: 1,
      index: 1,
      kind: 'insert-rows' as const,
      operationId: 'move-chart-ranges',
      sheetKey,
    };
    const edited = await applyXlsxEdits(snapshot, [operation]);
    const previewSheet = edited.document.sheets[0]!;
    expect(previewSheet.kind).toBe('worksheet');
    if (previewSheet.kind !== 'worksheet')
      throw new Error('Expected worksheet');
    const previewChart = previewSheet.drawings[0]!.object;
    expect(previewChart.kind).toBe('chart');
    if (previewChart.kind !== 'chart') throw new Error('Expected chart');
    expect(previewChart.plots[0]!.series[0]!.categories?.formula).toBe(
      'Sheet1!$A$3:$A$4',
    );
    expect(previewChart.plots[0]!.series[0]!.values?.formula).toBe(
      'Sheet1!$B$3:$B$4',
    );
    expect(previewSheet.drawings[0]!.from).toMatchObject({ row: 5 });
    const result = await writeXlsxRoundTrip(portable(edited));
    expect(result.report.level).toBe('R2');
    expect(
      result.report.parts
        .filter((part) => part.disposition === 'patch')
        .map((part) => part.name),
    ).toEqual([
      'xl/charts/chart1.xml',
      'xl/drawings/drawing1.xml',
      'xl/worksheets/sheet1.xml',
    ]);
    const parsed = await parseXlsx(result.data, { errorMode: 'strict' });
    const outputSheet = parsed.sheets[0]!;
    expect(outputSheet.kind).toBe('worksheet');
    if (outputSheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    const outputChart = outputSheet.drawings[0]!.object;
    expect(outputChart.kind).toBe('chart');
    if (outputChart.kind !== 'chart') throw new Error('Expected chart');
    expect(outputChart.plots[0]!.series[0]!.categories).toMatchObject({
      formula: 'Sheet1!$A$3:$A$4',
      pointCount: 2,
      points: [
        { index: 0, value: 'Q1' },
        { index: 1, value: 'Q2' },
      ],
    });
    expect(
      (await JSZip.loadAsync(result.data)).file('xl/charts/chart1.xml')!.date,
    ).toEqual(chartDate);
    const zip = await JSZip.loadAsync(source);
    const request = {
      count: operation.count,
      index: operation.index,
      kind: operation.kind,
      operationId: operation.operationId,
    };
    const worksheetPatch = patchXlsxWorksheetStructure(
      await zip.file('xl/worksheets/sheet1.xml')!.async('uint8array'),
      [request],
      defaultXlsxWriteLimits(),
      'xl/worksheets/sheet1.xml',
    );
    const drawingPatch = patchXlsxDrawingStructure(
      await zip.file('xl/drawings/drawing1.xml')!.async('uint8array'),
      [request],
      defaultXlsxWriteLimits(),
      'xl/drawings/drawing1.xml',
    );
    const chartPatch = patchXlsxChartStructure(
      await zip.file('xl/charts/chart1.xml')!.async('uint8array'),
      [request],
      defaultXlsxWriteLimits(),
      'xl/charts/chart1.xml',
      'Sheet1',
    );
    const patchBytes =
      worksheetPatch.patchBytes +
      drawingPatch.patchBytes +
      chartPatch.patchBytes;
    const patchCount =
      worksheetPatch.patchCount +
      drawingPatch.patchCount +
      chartPatch.patchCount;
    const generatedXmlBytes =
      worksheetPatch.data.byteLength +
      drawingPatch.data.byteLength +
      chartPatch.data.byteLength;
    await expect(
      writeXlsxRoundTrip(portable(edited), {
        limits: {
          maxDependencyEdges: 3,
          maxDirtyParts: 3,
          maxGeneratedXmlBytes: generatedXmlBytes,
          maxPatchBytes: patchBytes,
          maxPatchCount: patchCount,
          maxPatchedParts: 3,
        },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    for (const [limitName, limit] of [
      ['maxDependencyEdges', 2],
      ['maxDirtyParts', 2],
      ['maxGeneratedXmlBytes', generatedXmlBytes - 1],
      ['maxPatchBytes', patchBytes - 1],
      ['maxPatchCount', patchCount - 1],
      ['maxPatchedParts', 2],
    ] as const) {
      await expect(
        writeXlsxRoundTrip(portable(edited), {
          limits: { [limitName]: limit },
        }),
      ).rejects.toMatchObject({
        diagnostic: { code: 'resource-limit-exceeded', limitName },
      });
    }
    const chartOnlyContentTypes = contentTypes.replace(
      '</Types>',
      '<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/xl/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
    );
    const drawingAtA1 = drawing.replace(
      '<xdr:col>3</xdr:col>',
      '<xdr:col>0</xdr:col>',
    );
    const literalChart = `<c:chartSpace xmlns:c="${chartNamespace}"><c:chart/></c:chartSpace>`;
    const chartOnlySource = await createIndependentXlsx({
      '[Content_Types].xml': chartOnlyContentTypes,
      'xl/charts/chart1.xml': chart,
      'xl/charts/chart2.xml': chart,
      'xl/charts/chart3.xml': literalChart,
      'xl/drawings/_rels/drawing1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="chart" Type="${XLSX_OFFICE_REL_TYPE}chart" Target="../charts/chart1.xml"/><Relationship Id="chart-two" Type="${XLSX_OFFICE_REL_TYPE}chart" Target="../charts/chart2.xml"/><Relationship Id="literal" Type="${XLSX_OFFICE_REL_TYPE}chart" Target="../charts/chart3.xml"/></Relationships>`,
      'xl/drawings/_rels/drawing2.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="shared" Type="${XLSX_OFFICE_REL_TYPE}chart" Target="../charts/chart1.xml"/></Relationships>`,
      'xl/drawings/drawing1.xml': drawingAtA1,
      'xl/drawings/drawing2.xml': drawingAtA1,
      'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="drawing" Type="${XLSX_OFFICE_REL_TYPE}drawing" Target="../drawings/drawing1.xml"/><Relationship Id="unused-drawing" Type="${XLSX_OFFICE_REL_TYPE}drawing" Target="../drawings/drawing2.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheetData><row r="1"><c r="A1"><v>0</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>2</v></c></row></sheetData><drawing r:id="drawing"/></worksheet>`,
    });
    const chartOnlySnapshot = await readXlsxRoundTrip(chartOnlySource);
    const chartOnlyEdited = await applyXlsxEdits(chartOnlySnapshot, [
      {
        count: 1,
        index: 2,
        kind: 'insert-columns',
        operationId: 'chart-only-range-shift',
        sheetKey: chartOnlySnapshot.document.sheets[0]!.key,
      },
    ]);
    const chartOnlyResult = await writeXlsxRoundTrip(chartOnlyEdited, {
      limits: { maxDependencyEdges: 4 },
    });
    expect(
      chartOnlyResult.report.parts
        .filter((part) => part.disposition === 'patch')
        .map((part) => part.name),
    ).toEqual([
      'xl/charts/chart1.xml',
      'xl/charts/chart2.xml',
      'xl/worksheets/sheet1.xml',
    ]);
    for (const copied of [
      'xl/charts/chart3.xml',
      'xl/drawings/drawing1.xml',
      'xl/drawings/drawing2.xml',
    ]) {
      expect(
        chartOnlyResult.report.parts.find((part) => part.name === copied)
          ?.disposition,
      ).toBe('copy');
    }
    await expect(
      writeXlsxRoundTrip(chartOnlyEdited, {
        limits: { maxDependencyEdges: 3 },
      }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'resource-limit-exceeded',
        limitName: 'maxDependencyEdges',
      },
    });
    await expect(
      applyXlsxEdits(snapshot, [
        {
          count: 1,
          index: 3,
          kind: 'insert-rows',
          operationId: 'expand-chart-cache',
          sheetKey,
        },
      ]),
    ).rejects.toMatchObject({
      diagnostic: { featureClass: 'chart-cache-cardinality' },
    });
  });
});
