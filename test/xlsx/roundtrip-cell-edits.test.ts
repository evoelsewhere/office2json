import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  applyXlsxEdits,
  parseXlsx,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
  XlsxWriteError,
} from '../../src/formats/xlsx';
import type {
  XlsxEditOperation,
  XlsxRoundTripSnapshot,
} from '../../src/formats/xlsx/roundtrip';
import type { XlsxStyle } from '../../src/formats/xlsx/types';
import { sha256XlsxBytes } from '../../src/formats/xlsx/roundtrip/digest';
import { appendXlsxStylesPart } from '../../src/formats/xlsx/roundtrip/style-append';
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

const TWO_STYLE_XML = `<styleSheet xmlns="${XLSX_SPREADSHEET_NS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

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
  throw new Error('Expected XLSX cell edit to fail');
}

function operation(
  snapshot: XlsxRoundTripSnapshot,
  overrides: Partial<Extract<XlsxEditOperation, { kind: 'set-cell' }>> = {},
): Extract<XlsxEditOperation, { kind: 'set-cell' }> {
  return {
    cell: 'A1',
    content: { kind: 'value', value: { kind: 'text', text: 'updated' } },
    kind: 'set-cell',
    operationId: 'edit-1',
    sheetKey: snapshot.document.sheets[0]!.key,
    ...overrides,
  };
}

async function zipPart(bytes: Uint8Array, name: string): Promise<Uint8Array> {
  return (await JSZip.loadAsync(bytes)).file(name)!.async('uint8array');
}

async function createTwoSheetXlsx(): Promise<Uint8Array> {
  return createIndependentXlsx({
    '[Content_Types].xml': `<?xml version="1.0"?><Types xmlns="${XLSX_CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="rIdSheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rIdSheet2" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rIdStyles" Type="${XLSX_OFFICE_REL_TYPE}styles" Target="styles.xml"/><Relationship Id="rIdSharedStrings" Type="${XLSX_OFFICE_REL_TYPE}sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/><sheet name="Sheet2" sheetId="2" r:id="rIdSheet2"/></sheets></workbook>`,
    'xl/worksheets/sheet1.xml': independentWorksheet(
      '<row r="1"><c r="A1"><v>1</v></c></row>',
    ),
    'xl/worksheets/sheet2.xml': independentWorksheet(
      '<row r="1"><c r="A1"><v>2</v></c></row>',
    ),
  });
}

describe('XLSX verified cell edits', () => {
  it.each([
    [
      'text',
      { kind: 'value', value: { kind: 'text', text: ' <new> &\rline ' } },
    ],
    ['number', { kind: 'value', value: { kind: 'number', value: -12.5 } }],
    ['boolean', { kind: 'value', value: { kind: 'boolean', value: false } }],
    ['error', { kind: 'value', value: { code: '#N/A', kind: 'error' } }],
    ['formula', { expression: 'SUM(B2:C3)', kind: 'formula' }],
  ] as const)(
    'applies portable %s edits with R1 and R2 evidence',
    async (_name, content) => {
      const source = await createIndependentXlsx();
      const sourceBefore = source.slice();
      const snapshot = await readXlsxRoundTrip(source);
      const edited = await applyXlsxEdits(snapshot, [
        operation(snapshot, { content }),
      ]);
      expect(snapshot.operations).toEqual([]);
      expect(source).toEqual(sourceBefore);
      expect(edited.operations).toHaveLength(1);
      expect(edited.stateHash).not.toBe(edited.baseDocumentHash);
      expect(edited.document.key).toBe(snapshot.document.key);
      expect(edited.document.sheets[0]!.key).toBe(
        snapshot.document.sheets[0]!.key,
      );

      const validated = await validateXlsxRoundTripJson(portable(edited));
      const result = await writeXlsxRoundTrip(validated);
      expect(result.report.level).toBe('R2');
      expect(result.report.sourceSha256).toBe(snapshot.source.sha256);
      expect(result.report.outputSha256).toBe(
        await sha256XlsxBytes(result.data),
      );
      expect(
        result.report.parts.filter((part) => part.disposition === 'patch'),
      ).toEqual([
        expect.objectContaining({ name: 'xl/worksheets/sheet1.xml' }),
      ]);
      expect(
        result.report.parts.filter((part) => part.disposition === 'copy')
          .length,
      ).toBe(result.report.parts.length - 1);
      for (const part of result.report.parts.filter(
        (candidate) => candidate.disposition === 'copy',
      )) {
        expect(part.sha256).toBe(part.sourceSha256);
        expect(part.byteLength).toBe(part.sourceByteLength);
        expect(await zipPart(result.data, part.name)).toEqual(
          await zipPart(source, part.name),
        );
      }
      const parsed = await parseXlsx(result.data, { errorMode: 'strict' });
      const sheet = parsed.sheets[0]!;
      expect(sheet.kind).toBe('worksheet');
      expect(
        sheet.kind === 'worksheet' ? sheet.rows[0]!.cells[0]!.content : null,
      ).toEqual(
        content.kind === 'formula'
          ? {
              cached: { kind: 'missing' },
              formula: { expression: content.expression, kind: 'normal' },
              kind: 'formula',
            }
          : content,
      );
      expect(result.report.diagnostics).toEqual(
        content.kind === 'formula'
          ? [
              expect.objectContaining({
                code: 'recalculation-required',
                message:
                  'The edited XLSX formula has no cached result and requires producer recalculation',
                operationId: 'edit-1',
                severity: 'warning',
              }),
            ]
          : [],
      );
    },
  );

  it('clears cells and preserves their authored style attribute', async () => {
    const source = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': independentWorksheet(
        '<row r="1"><c r="A1" s="0" t="s"><v>0</v></c></row>',
      ),
    });
    const snapshot = await readXlsxRoundTrip(source);
    const edited = await applyXlsxEdits(snapshot, [
      {
        cell: 'A1',
        kind: 'clear-cell',
        operationId: 'clear-1',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    const result = await writeXlsxRoundTrip(edited);
    expect(
      new TextDecoder().decode(
        await zipPart(result.data, 'xl/worksheets/sheet1.xml'),
      ),
    ).toContain('<c r="A1" s="0"/>');
    const parsed = await parseXlsx(result.data);
    const sheet = parsed.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.rows[0]!.cells[0]!.content : null,
    ).toEqual({
      kind: 'blank',
    });
  });

  it('clears a normal formula as an explicit content-closure target', async () => {
    const snapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': independentWorksheet(
          '<row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row>',
        ),
      }),
    );
    const edited = await applyXlsxEdits(snapshot, [
      {
        cell: 'A1',
        kind: 'clear-cell',
        operationId: 'clear-formula',
        sheetKey: snapshot.document.sheets[0]!.key,
      },
    ]);
    await expect(writeXlsxRoundTrip(edited)).resolves.toMatchObject({
      report: { level: 'R2' },
    });
    const dependentSnapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': independentWorksheet(
          '<row r="1"><c r="A1"><f>1+1</f><v>2</v></c><c r="B1"><f>A1+1</f><v>3</v></c></row>',
        ),
      }),
    );
    const dependentEdited = await applyXlsxEdits(dependentSnapshot, [
      {
        cell: 'A1',
        kind: 'clear-cell',
        operationId: 'clear-formula',
        sheetKey: dependentSnapshot.document.sheets[0]!.key,
      },
    ]);
    expect(
      (await capture(() => writeXlsxRoundTrip(dependentEdited))).diagnostic,
    ).toMatchObject({
      cell: 'B1',
      code: 'formula-rewrite-unsupported',
      featureClass: 'formula-dependency',
    });
  });

  it('applies an existing normalized style with exact copied style bytes and R2 evidence', async () => {
    const source = await createIndependentXlsx({
      'xl/styles.xml': TWO_STYLE_XML,
      'xl/worksheets/sheet1.xml': independentWorksheet(
        '<row r="1"><c r="A1" s="0" t="s"><v>0</v></c></row>',
      ),
    });
    const snapshot = await readXlsxRoundTrip(source);
    const targetStyle = snapshot.document.styles[1]!;
    const edited = await applyXlsxEdits(snapshot, [
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'style-1',
        sheetKey: snapshot.document.sheets[0]!.key,
        style: targetStyle,
      },
    ]);
    expect(edited.document.sheets[0]!.kind).toBe('worksheet');
    if (edited.document.sheets[0]!.kind !== 'worksheet') {
      throw new Error('Expected worksheet');
    }
    expect(edited.document.sheets[0]!.rows[0]!.cells[0]!.style).toBe(1);
    const validated = await validateXlsxRoundTripJson(portable(edited));
    const result = await writeXlsxRoundTrip(validated);
    expect(result.report.level).toBe('R2');
    expect(
      result.report.parts.filter((part) => part.disposition === 'patch'),
    ).toEqual([expect.objectContaining({ name: 'xl/worksheets/sheet1.xml' })]);
    const stylePart = result.report.parts.find(
      (part) => part.name === 'xl/styles.xml',
    )!;
    expect(stylePart.disposition).toBe('copy');
    expect(stylePart.sha256).toBe(stylePart.sourceSha256);
    expect(await zipPart(result.data, 'xl/styles.xml')).toEqual(
      await zipPart(source, 'xl/styles.xml'),
    );
    const worksheetXml = new TextDecoder().decode(
      await zipPart(result.data, 'xl/worksheets/sheet1.xml'),
    );
    expect(worksheetXml).toContain('<c r="A1" t="s" s="1">');
    expect(worksheetXml).not.toContain('s="0"');
    const parsed = await parseXlsx(result.data, { errorMode: 'strict' });
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows[0]!.cells[0]!.style).toBe(1);
    expect(parsed.styles[1]).toEqual(targetStyle);
    expect(result.report.diagnostics).toEqual([]);
  });

  it('applies an existing style to an explicit self-closing blank cell', async () => {
    const source = await createIndependentXlsx({
      'xl/styles.xml': TWO_STYLE_XML,
      'xl/worksheets/sheet1.xml': independentWorksheet(
        '<row r="1"><c r="A1"/></row>',
      ),
    });
    const snapshot = await readXlsxRoundTrip(source);
    const edited = await applyXlsxEdits(snapshot, [
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'style-blank',
        sheetKey: snapshot.document.sheets[0]!.key,
        style: snapshot.document.styles[1]!,
      },
    ]);
    const result = await writeXlsxRoundTrip(edited);
    expect(
      new TextDecoder().decode(
        await zipPart(result.data, 'xl/worksheets/sheet1.xml'),
      ),
    ).toContain('<c r="A1" s="1"/>');
    expect(result.report.level).toBe('R2');
  });

  it('appends one normalized style deterministically and verifies both dirty parts', async () => {
    const source = await createIndependentXlsx({
      'xl/styles.xml': TWO_STYLE_XML,
    });
    const snapshot = await readXlsxRoundTrip(source);
    const target: XlsxStyle = {
      alignment: {
        horizontal: 'center',
        indent: 2,
        justifyLastLine: true,
        readingOrder: 'right-to-left',
        relativeIndent: -3,
        shrinkToFit: true,
        textRotation: 180,
        vertical: 'top',
        wrapText: true,
      },
      border: {
        bottom: {
          color: { argb: 'FF112233', kind: 'rgb' },
          style: 'thin',
        },
        diagonal: { style: 'dashDot' },
        diagonalDown: true,
        diagonalUp: true,
        end: { style: 'dashDotDot' },
        horizontal: { style: 'dashed' },
        left: { style: 'dotted' },
        outline: false,
        right: { style: 'double' },
        start: { style: 'hair' },
        top: { style: 'medium' },
        vertical: { style: 'mediumDashDot' },
      },
      fill: {
        backgroundColor: { kind: 'automatic' },
        foregroundColor: { index: 4, kind: 'theme', tint: 0.25 },
        kind: 'pattern',
        pattern: 'solid',
      },
      font: {
        bold: true,
        charset: 255,
        color: { index: 3, kind: 'indexed', tint: -0.2 },
        condense: true,
        extend: true,
        family: 5,
        italic: true,
        name: 'Agent & "Style"',
        outline: true,
        scheme: 'minor',
        shadow: true,
        size: 12.5,
        strike: true,
        underline: 'double-accounting',
        verticalAlignment: 'superscript',
      },
      numberFormat: '0.0000 "kg" &',
      protection: { hidden: true, locked: false },
    };
    const secondTarget: XlsxStyle = { font: { italic: true } };
    const edited = await applyXlsxEdits(snapshot, [
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'append-style-a1',
        sheetKey: snapshot.document.sheets[0]!.key,
        style: target,
      },
      {
        cell: 'B2',
        kind: 'set-cell-style',
        operationId: 'reuse-style-b2',
        sheetKey: snapshot.document.sheets[0]!.key,
        style: structuredClone(target),
      },
      {
        cell: 'C3',
        kind: 'set-cell-style',
        operationId: 'append-second-style-c3',
        sheetKey: snapshot.document.sheets[0]!.key,
        style: secondTarget,
      },
    ]);
    expect(edited.document.styles).toEqual([
      ...snapshot.document.styles,
      target,
      secondTarget,
    ]);
    const [first, second] = await Promise.all([
      writeXlsxRoundTrip(portable(edited)),
      writeXlsxRoundTrip(portable(edited)),
    ]);
    expect(second).toEqual(first);
    expect(
      first.report.parts
        .filter((part) => part.disposition === 'patch')
        .map((part) => part.name),
    ).toEqual(['xl/styles.xml', 'xl/worksheets/sheet1.xml']);
    const styleXml = new TextDecoder().decode(
      await zipPart(first.data, 'xl/styles.xml'),
    );
    expect(styleXml).toContain('<numFmts count="1">');
    expect(styleXml).toContain('<fonts count="4">');
    expect(styleXml).toContain('<fills count="4">');
    expect(styleXml).toContain('<borders count="3">');
    expect(styleXml).toContain('<cellXfs count="4">');
    expect(styleXml).toContain('formatCode="0.0000 &quot;kg&quot; &amp;"');
    const parsed = await parseXlsx(first.data, { errorMode: 'strict' });
    expect(parsed.styles[2]).toEqual(target);
    expect(parsed.styles[3]).toEqual(secondTarget);
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows[0]!.cells[0]!.style).toBe(2);
    expect(sheet.rows[1]!.cells[0]!.style).toBe(2);
    expect(sheet.rows[2]!.cells[0]!.style).toBe(3);
    expect(first.report.level).toBe('R2');
    const generatedXmlBytes = first.report.parts
      .filter((part) => part.disposition === 'patch')
      .reduce((total, part) => total + part.byteLength, 0);
    await expect(
      writeXlsxRoundTrip(edited, {
        limits: { maxGeneratedXmlBytes: generatedXmlBytes },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(edited, {
            limits: { maxGeneratedXmlBytes: generatedXmlBytes - 1 },
          }),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: generatedXmlBytes,
      limit: generatedXmlBytes - 1,
      limitName: 'maxGeneratedXmlBytes',
    });
    const stylePatchBytes = appendXlsxStylesPart(
      await zipPart(source, 'xl/styles.xml'),
      [target, secondTarget],
      defaultXlsxWriteLimits(),
      'xl/styles.xml',
    ).patchBytes;
    const aggregatePatchError = await capture(() =>
      writeXlsxRoundTrip(edited, {
        limits: { maxPatchBytes: stylePatchBytes },
      }),
    );
    expect(aggregatePatchError.diagnostic).toMatchObject({
      code: 'resource-limit-exceeded',
      limit: stylePatchBytes,
      limitName: 'maxPatchBytes',
    });
    expect(aggregatePatchError.diagnostic.actual).toBeGreaterThan(
      stylePatchBytes,
    );
    await expect(
      writeXlsxRoundTrip(edited, {
        limits: {
          maxDependencyEdges: 5,
          maxDirtyParts: 2,
          maxPatchCount: 12,
          maxPatchedParts: 2,
        },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    for (const [limitName, actual] of [
      ['maxDependencyEdges', 5],
      ['maxDirtyParts', 2],
      ['maxPatchCount', 12],
      ['maxPatchedParts', 2],
    ] as const) {
      expect(
        (
          await capture(() =>
            writeXlsxRoundTrip(edited, {
              limits: { [limitName]: actual - 1 },
            }),
          )
        ).diagnostic,
      ).toMatchObject({
        actual,
        code: 'resource-limit-exceeded',
        limit: actual - 1,
        limitName,
      });
    }
    await expect(
      writeXlsxRoundTrip(edited, { readerLimits: { maxStyles: 18 } }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(edited, { readerLimits: { maxStyles: 17 } }),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 18,
      code: 'resource-limit-exceeded',
      limit: 17,
      limitName: 'maxStyles',
      message: 'XLSX appended style records exceed the reader limit',
    });
  });

  it('preserves the authored ZIP timestamp when patching the styles part', async () => {
    const original = await createIndependentXlsx({
      'xl/styles.xml': TWO_STYLE_XML,
    });
    const archive = await JSZip.loadAsync(original);
    const entry = archive.file('xl/styles.xml')!;
    const authoredDate = new Date('2002-03-04T05:06:08.000Z');
    archive.file('xl/styles.xml', await entry.async('uint8array'), {
      date: authoredDate,
    });
    const source = await archive.generateAsync({
      compression: 'DEFLATE',
      type: 'uint8array',
    });
    const snapshot = await readXlsxRoundTrip(source);
    const edited = await applyXlsxEdits(snapshot, [
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'append-timestamp-style',
        sheetKey: snapshot.document.sheets[0]!.key,
        style: { font: { italic: true } },
      },
    ]);
    const result = await writeXlsxRoundTrip(edited);
    expect(
      (await JSZip.loadAsync(result.data))
        .file('xl/styles.xml')!
        .date.toISOString(),
    ).toBe(authoredDate.toISOString());
  });

  it('keeps style-only formula content independent and blocks date interpretation changes', async () => {
    const formulaSource = await createIndependentXlsx({
      'xl/styles.xml': TWO_STYLE_XML,
      'xl/worksheets/sheet1.xml': independentWorksheet(
        '<row r="1"><c r="A1"><f t="array" ref="A1">1+1</f><v>2</v></c></row>',
      ),
    });
    const formulaSnapshot = await readXlsxRoundTrip(formulaSource);
    const formulaEdited = await applyXlsxEdits(formulaSnapshot, [
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'style-formula',
        sheetKey: formulaSnapshot.document.sheets[0]!.key,
        style: formulaSnapshot.document.styles[1]!,
      },
    ]);
    await expect(writeXlsxRoundTrip(formulaEdited)).resolves.toMatchObject({
      report: { diagnostics: [], level: 'R2' },
    });

    const dateStyles = TWO_STYLE_XML.replace(
      '<cellXfs count="2">',
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><cellXfs count="2">',
    ).replace('<xf numFmtId="0" fontId="1"', '<xf numFmtId="164" fontId="1"');
    const dateSource = await createIndependentXlsx({
      'xl/styles.xml': dateStyles,
      'xl/worksheets/sheet1.xml': independentWorksheet(
        '<row r="1"><c r="A1" s="1"><v>2</v></c></row>',
      ),
    });
    const dateSnapshot = await readXlsxRoundTrip(dateSource);
    const dateEdited = await applyXlsxEdits(dateSnapshot, [
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'style-date',
        sheetKey: dateSnapshot.document.sheets[0]!.key,
        style: dateSnapshot.document.styles[0]!,
      },
    ]);
    expect(
      (await capture(() => writeXlsxRoundTrip(dateEdited))).diagnostic,
    ).toMatchObject({
      cell: 'A1',
      code: 'preservation-conflict',
      featureClass: 'date-style-conversion',
      message: 'XLSX style edit changes the cell date-value interpretation',
      operationId: 'style-date',
    });

    const sameDateStyles = dateStyles
      .replace('<fonts count="2">', '<fonts count="3">')
      .replace(
        '</fonts>',
        '<font><i/><sz val="11"/><name val="Calibri"/></font></fonts>',
      )
      .replace('<cellXfs count="2">', '<cellXfs count="3">')
      .replace(
        '</cellXfs>',
        '<xf numFmtId="164" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/></cellXfs>',
      );
    const sameDateSource = await createIndependentXlsx({
      'xl/styles.xml': sameDateStyles,
      'xl/worksheets/sheet1.xml': independentWorksheet(
        '<row r="1"><c r="A1" s="1"><v>2</v></c></row>',
      ),
    });
    const sameDateSnapshot = await readXlsxRoundTrip(sameDateSource);
    const sameDateEdited = await applyXlsxEdits(sameDateSnapshot, [
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'style-date-font',
        sheetKey: sameDateSnapshot.document.sheets[0]!.key,
        style: sameDateSnapshot.document.styles[2]!,
      },
    ]);
    await expect(writeXlsxRoundTrip(sameDateEdited)).resolves.toMatchObject({
      report: { level: 'R2' },
    });
  });

  it('does not treat a style target as a content formula-closure target', async () => {
    const source = await createIndependentXlsx({
      'xl/styles.xml': TWO_STYLE_XML,
      'xl/worksheets/sheet1.xml': independentWorksheet(
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>1+1</f><v>2</v></c></row>',
      ),
    });
    const snapshot = await readXlsxRoundTrip(source);
    const edited = await applyXlsxEdits(snapshot, [
      operation(snapshot, {
        content: { kind: 'value', value: { kind: 'number', value: 3 } },
      }),
      {
        cell: 'B1',
        kind: 'set-cell-style',
        operationId: 'style-formula',
        sheetKey: snapshot.document.sheets[0]!.key,
        style: snapshot.document.styles[1]!,
      },
    ]);
    expect(
      (await capture(() => writeXlsxRoundTrip(edited))).diagnostic,
    ).toMatchObject({
      cell: 'B1',
      code: 'formula-rewrite-unsupported',
      featureClass: 'formula-dependency',
    });
  });

  it('is deterministic and isolated across repeated concurrent writes', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const edited = await applyXlsxEdits(snapshot, [operation(snapshot)]);
    const before = portable(edited);
    const results = await Promise.all(
      Array.from({ length: 6 }, () => writeXlsxRoundTrip(edited)),
    );
    for (const result of results) {
      expect(result.data).toEqual(results[0]!.data);
      expect(result.report).toEqual(results[0]!.report);
    }
    expect(edited).toEqual(before);
    expect(new Set(results.map((result) => result.data)).size).toBe(6);
  });

  it('replays sequential operations and validates edited JSON independently', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const first = await applyXlsxEdits(snapshot, [
      operation(snapshot, {
        content: { kind: 'value', value: { kind: 'number', value: 1 } },
      }),
    ]);
    const second = await applyXlsxEdits(first, [
      operation(snapshot, {
        content: { kind: 'value', value: { kind: 'number', value: 2 } },
        operationId: 'edit-2',
      }),
    ]);
    expect(second.operations.map((item) => item.operationId)).toEqual([
      'edit-1',
      'edit-2',
    ]);
    await expect(validateXlsxRoundTripJson(portable(second))).resolves.toEqual(
      second,
    );
    const result = await writeXlsxRoundTrip(second);
    const parsed = await parseXlsx(result.data);
    const sheet = parsed.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.rows[0]!.cells[0]!.content : null,
    ).toEqual({ kind: 'value', value: { kind: 'number', value: 2 } });
  });

  it('patches distinct cells on the same worksheet in one atomic closure', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const edited = await applyXlsxEdits(snapshot, [
      operation(snapshot, {
        content: { kind: 'value', value: { kind: 'number', value: 10 } },
      }),
      operation(snapshot, {
        cell: 'B2',
        content: { kind: 'value', value: { kind: 'number', value: 20 } },
        operationId: 'edit-2',
      }),
    ]);
    const result = await writeXlsxRoundTrip(edited);
    const parsed = await parseXlsx(result.data);
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows[0]!.cells[0]!.content).toEqual({
      kind: 'value',
      value: { kind: 'number', value: 10 },
    });
    expect(sheet.rows[1]!.cells[0]!.content).toEqual({
      kind: 'value',
      value: { kind: 'number', value: 20 },
    });
  });

  it('preserves the source ZIP timestamp for patched worksheet parts', async () => {
    const original = await createIndependentXlsx();
    const archive = await JSZip.loadAsync(original);
    const entry = archive.file('xl/worksheets/sheet1.xml')!;
    const authoredDate = new Date('2001-02-03T04:05:06.000Z');
    archive.file('xl/worksheets/sheet1.xml', await entry.async('uint8array'), {
      date: authoredDate,
    });
    const source = await archive.generateAsync({
      compression: 'DEFLATE',
      type: 'uint8array',
    });
    const snapshot = await readXlsxRoundTrip(source);
    const edited = await applyXlsxEdits(snapshot, [operation(snapshot)]);
    const result = await writeXlsxRoundTrip(edited);
    const outputEntry = (await JSZip.loadAsync(result.data)).file(
      'xl/worksheets/sheet1.xml',
    )!;
    expect(outputEntry.date.toISOString()).toBe(authoredDate.toISOString());
  });

  it('rejects R3 requests and exact validation-pass/output boundaries', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const edited = await applyXlsxEdits(snapshot, [operation(snapshot)]);
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(edited, { minimumEditedFidelity: 'R3' }),
        )
      ).diagnostic,
    ).toMatchObject({
      code: 'producer-verification-failed',
      fidelity: 'R3',
      message: 'The XLSX cell-edit profile has no producer R3 evidence',
    });
    await expect(
      writeXlsxRoundTrip(edited, {
        limits: { maxValidationPasses: 4 },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(edited, {
            limits: { maxValidationPasses: 3 },
          }),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 4,
      code: 'resource-limit-exceeded',
      limit: 3,
      limitName: 'maxValidationPasses',
    });
    const successful = await writeXlsxRoundTrip(edited);
    await expect(
      writeXlsxRoundTrip(edited, {
        limits: {
          maxOutputBytes: successful.data.byteLength,
          maxSourcePackageBytes: snapshot.source.byteLength,
        },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(edited, {
            limits: {
              maxOutputBytes: successful.data.byteLength - 1,
              maxSourcePackageBytes: snapshot.source.byteLength,
            },
          }),
        )
      ).diagnostic.limitName,
    ).toBe('maxOutputBytes');
  });

  it('enforces dirty-part, patched-part, and dependency budgets exactly', async () => {
    const snapshot = await readXlsxRoundTrip(await createTwoSheetXlsx());
    const edited = await applyXlsxEdits(snapshot, [
      operation(snapshot, {
        content: { kind: 'value', value: { kind: 'number', value: 10 } },
      }),
      operation(snapshot, {
        content: { kind: 'value', value: { kind: 'number', value: 20 } },
        operationId: 'edit-2',
        sheetKey: snapshot.document.sheets[1]!.key,
      }),
    ]);
    const successful = await writeXlsxRoundTrip(edited, {
      limits: {
        maxDependencyEdges: 2,
        maxDirtyParts: 2,
        maxPatchCount: 2,
        maxPatchedParts: 2,
      },
    });
    expect(successful.report.level).toBe('R2');
    expect(
      successful.report.parts
        .filter((part) => part.disposition === 'patch')
        .map((part) => part.name),
    ).toEqual(['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']);
    for (const limitName of [
      'maxDependencyEdges',
      'maxDirtyParts',
      'maxPatchCount',
      'maxPatchedParts',
    ] as const) {
      const error = await capture(() =>
        writeXlsxRoundTrip(edited, { limits: { [limitName]: 1 } }),
      );
      expect(error.diagnostic).toMatchObject({
        actual: 2,
        code: 'resource-limit-exceeded',
        limit: 1,
        limitName,
        ...(limitName === 'maxPatchCount'
          ? { part: 'xl/worksheets/sheet2.xml' }
          : {}),
      });
    }
    const patchedParts = successful.report.parts.filter(
      (part) => part.disposition === 'patch',
    );
    const generatedXmlBytes = patchedParts.reduce(
      (total, part) => total + part.byteLength,
      0,
    );
    await expect(
      writeXlsxRoundTrip(edited, {
        limits: { maxGeneratedXmlBytes: generatedXmlBytes },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(edited, {
            limits: { maxGeneratedXmlBytes: generatedXmlBytes - 1 },
          }),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: generatedXmlBytes,
      limit: generatedXmlBytes - 1,
      limitName: 'maxGeneratedXmlBytes',
      part: 'xl/worksheets/sheet2.xml',
    });
    const replacementBytes = new TextEncoder().encode(
      '<c r="A1"><v>10</v></c>',
    ).byteLength;
    await expect(
      writeXlsxRoundTrip(edited, {
        limits: { maxPatchBytes: replacementBytes * 2 },
      }),
    ).resolves.toMatchObject({ report: { level: 'R2' } });
    expect(
      (
        await capture(() =>
          writeXlsxRoundTrip(edited, {
            limits: { maxPatchBytes: replacementBytes * 2 - 1 },
          }),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: replacementBytes * 2,
      limit: replacementBytes * 2 - 1,
      limitName: 'maxPatchBytes',
      part: 'xl/worksheets/sheet2.xml',
    });
  });

  it('blocks date-formatted numbers, grouped formulas, and target extensions', async () => {
    const styles = `<styleSheet xmlns="${XLSX_SPREADSHEET_NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    const dateSnapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/styles.xml': styles,
        'xl/worksheets/sheet1.xml': independentWorksheet(
          '<row r="1"><c r="A1" s="1"><v>2</v></c></row>',
        ),
      }),
    );
    const dateEdited = await applyXlsxEdits(dateSnapshot, [
      operation(dateSnapshot, {
        content: { kind: 'value', value: { kind: 'number', value: 3 } },
      }),
    ]);
    expect(
      (await capture(() => writeXlsxRoundTrip(dateEdited))).diagnostic,
    ).toMatchObject({
      code: 'preservation-conflict',
      featureClass: 'date-formatted-cell',
      message: 'XLSX number edit targets a date-formatted cell',
    });
    const numberSnapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/styles.xml': styles.replace('yyyy-mm-dd', '0.00'),
        'xl/worksheets/sheet1.xml': independentWorksheet(
          '<row r="1"><c r="A1" s="1"><v>2</v></c></row>',
        ),
      }),
    );
    const numberEdited = await applyXlsxEdits(numberSnapshot, [
      operation(numberSnapshot, {
        content: { kind: 'value', value: { kind: 'number', value: 3 } },
      }),
    ]);
    await expect(writeXlsxRoundTrip(numberEdited)).resolves.toMatchObject({
      report: { level: 'R2' },
    });

    const groupedSnapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': independentWorksheet(
          '<row r="1"><c r="A1"><f t="array" ref="A1">1+1</f></c></row>',
        ),
      }),
    );
    const groupedEdited = await applyXlsxEdits(groupedSnapshot, [
      operation(groupedSnapshot, {
        content: { expression: '2+2', kind: 'formula' },
      }),
    ]);
    expect(
      (await capture(() => writeXlsxRoundTrip(groupedEdited))).diagnostic,
    ).toMatchObject({
      code: 'formula-rewrite-unsupported',
      featureClass: 'formula-group',
    });

    const extensionSnapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': independentWorksheet(
          '<row r="1"><c r="A1" custom="x"><v>1</v></c></row>',
        ),
      }),
    );
    const extensionEdited = await applyXlsxEdits(extensionSnapshot, [
      operation(extensionSnapshot),
    ]);
    expect(
      (await capture(() => writeXlsxRoundTrip(extensionEdited))).diagnostic,
    ).toMatchObject({
      code: 'preservation-conflict',
      featureClass: 'cell-extension',
    });
  });

  it('validates options consistently across edit, JSON, and write APIs', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const invalid = { unknown: true } as never;
    await expect(applyXlsxEdits(snapshot, [], invalid)).rejects.toThrow(
      'Unknown XLSX write option unknown',
    );
    await expect(validateXlsxRoundTripJson(snapshot, invalid)).rejects.toThrow(
      'Unknown XLSX write option unknown',
    );
    await expect(writeXlsxRoundTrip(snapshot, invalid)).rejects.toThrow(
      'Unknown XLSX write option unknown',
    );
  });

  it('edits Strict workbooks and preserves safe external hyperlinks without fetching', async () => {
    const strictSheetNs = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictRelNs =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const strict = await createIndependentXlsx({
      '[Content_Types].xml': `<Types xmlns="${XLSX_CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelNs}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${strictRelNs}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheetNs}" xmlns:r="${strictRelNs}"><s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets></s:workbook>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheetNs}"><s:sheetData><s:row r="1"><s:c r="A1"><s:v>1</s:v></s:c></s:row></s:sheetData></s:worksheet>`,
    });
    const strictSnapshot = await readXlsxRoundTrip(strict);
    const strictEdited = await applyXlsxEdits(strictSnapshot, [
      operation(strictSnapshot, {
        content: { kind: 'value', value: { kind: 'number', value: 2 } },
      }),
    ]);
    const strictResult = await writeXlsxRoundTrip(strictEdited);
    expect(strictResult.report.level).toBe('R2');
    expect(
      (await readXlsxRoundTrip(strictResult.data)).source.conformance,
    ).toBe('strict');
    const missingStylePart = await applyXlsxEdits(strictSnapshot, [
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'append-without-styles',
        sheetKey: strictSnapshot.document.sheets[0]!.key,
        style: { font: { bold: true } },
      },
    ]);
    expect(
      (await capture(() => writeXlsxRoundTrip(missingStylePart))).diagnostic,
    ).toMatchObject({
      code: 'preservation-conflict',
      featureClass: 'missing-styles-part',
      message: 'XLSX cannot append styles without an existing styles part',
    });
    const strictWithStyles = await createIndependentXlsx({
      '[Content_Types].xml': `<Types xmlns="${XLSX_CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelNs}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${strictRelNs}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="styles" Type="${strictRelNs}/styles" Target="styles.xml"/></Relationships>`,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': `<s:styleSheet xmlns:s="${strictSheetNs}"><s:fonts count="1"><s:font/></s:fonts><s:fills count="1"><s:fill><s:patternFill patternType="none"/></s:fill></s:fills><s:borders count="1"><s:border/></s:borders><s:cellStyleXfs count="1"><s:xf/></s:cellStyleXfs><s:cellXfs count="1"><s:xf/></s:cellXfs></s:styleSheet>`,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheetNs}" xmlns:r="${strictRelNs}"><s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets></s:workbook>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheetNs}"><s:sheetData><s:row r="1"><s:c r="A1"><s:v>1</s:v></s:c></s:row></s:sheetData></s:worksheet>`,
    });
    const strictStyleSnapshot = await readXlsxRoundTrip(strictWithStyles);
    const strictStyleEdited = await applyXlsxEdits(strictStyleSnapshot, [
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'strict-append-style',
        sheetKey: strictStyleSnapshot.document.sheets[0]!.key,
        style: { font: { bold: true } },
      },
    ]);
    const strictStyleResult = await writeXlsxRoundTrip(strictStyleEdited);
    expect(strictStyleResult.report.level).toBe('R2');
    expect(
      (await readXlsxRoundTrip(strictStyleResult.data)).source.conformance,
    ).toBe('strict');

    const external = await createIndependentXlsx({
      'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="link" Type="${XLSX_OFFICE_REL_TYPE}hyperlink" Target="https://example.invalid/never-fetched" TargetMode="External"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><hyperlinks><hyperlink ref="A1" r:id="link"/></hyperlinks></worksheet>`,
    });
    const externalSnapshot = await readXlsxRoundTrip(external);
    const externalEdited = await applyXlsxEdits(externalSnapshot, [
      operation(externalSnapshot, {
        content: { kind: 'value', value: { kind: 'number', value: 2 } },
      }),
    ]);
    const externalResult = await writeXlsxRoundTrip(externalEdited);
    expect(externalResult.report.level).toBe('R2');
    const reparsed = await parseXlsx(externalResult.data);
    const sheet = reparsed.sheets[0]!;
    expect(sheet.kind === 'worksheet' ? sheet.hyperlinks : []).toEqual([
      expect.objectContaining({
        target: {
          kind: 'external',
          url: 'https://example.invalid/never-fetched',
        },
      }),
    ]);
  });

  it('blocks unaffected formulas, defined names, and external-capable formulas', async () => {
    const formulaSource = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': independentWorksheet(
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1+1</f><v>2</v></c></row>',
      ),
    });
    const formulaSnapshot = await readXlsxRoundTrip(formulaSource);
    const formulaEdited = await applyXlsxEdits(formulaSnapshot, [
      operation(formulaSnapshot),
    ]);
    expect(
      (await capture(() => writeXlsxRoundTrip(formulaEdited))).diagnostic,
    ).toMatchObject({
      code: 'formula-rewrite-unsupported',
      featureClass: 'formula-dependency',
      message:
        'XLSX cell edit dependency closure contains an unaffected formula',
    });

    const namesSource = await createIndependentXlsx({
      'xl/workbook.xml': `<?xml version="1.0"?><workbook xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/></sheets><definedNames><definedName name="Value">Sheet1!$A$1</definedName></definedNames></workbook>`,
    });
    const namesSnapshot = await readXlsxRoundTrip(namesSource);
    const namesEdited = await applyXlsxEdits(namesSnapshot, [
      operation(namesSnapshot),
    ]);
    expect(
      (await capture(() => writeXlsxRoundTrip(namesEdited))).diagnostic,
    ).toMatchObject({
      code: 'formula-rewrite-unsupported',
      featureClass: 'defined-name',
      message:
        'XLSX cell edit dependency closure contains defined-name formulas',
    });

    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const edited = await applyXlsxEdits(snapshot, [
      operation(snapshot, {
        content: {
          expression: 'WEBSERVICE("https://example.invalid")',
          kind: 'formula',
        },
      }),
    ]);
    expect(
      (await capture(() => writeXlsxRoundTrip(edited))).diagnostic,
    ).toMatchObject({
      code: 'formula-rewrite-unsupported',
      featureClass: 'external-formula',
      message:
        'XLSX cell edit formula uses an external-capable function or reference',
      operationId: 'edit-1',
    });
  });

  it('blocks signed, opaque, active, unknown-part, and relationship closures', async () => {
    const activeTypes = `<?xml version="1.0"?><Types xmlns="${XLSX_CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/embeddings/item.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>`;
    const cases: Array<{
      expected: string;
      expectedMessage?: string;
      options?: Parameters<typeof readXlsxRoundTrip>[1];
      overrides: Record<string, string | Uint8Array | null>;
      writeOptions?: Parameters<typeof writeXlsxRoundTrip>[1];
    }> = [
      {
        expected: 'signed-package-conflict',
        overrides: { '_xmlsignatures/sig1.xml': '<Signature/>' },
      },
      {
        expected: 'preservation-conflict',
        options: { securityMode: 'preserve-opaque' },
        overrides: {
          '[Content_Types].xml': activeTypes,
          'xl/embeddings/item.bin': new Uint8Array([1, 2, 3]),
        },
      },
      {
        expected: 'opaque-content-conflict',
        expectedMessage:
          'Opaque XLSX content has no proven independent cell-edit closure',
        options: { securityMode: 'preserve-opaque' },
        overrides: {
          '[Content_Types].xml': activeTypes.replace(
            '/xl/embeddings/item.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject',
            '/xl/opaque/item.bin" ContentType="application/octet-stream',
          ),
          'xl/opaque/item.bin': new Uint8Array([1, 2, 3]),
        },
        writeOptions: { acknowledgeOpaqueContent: true },
      },
      {
        expected: 'opaque-content-conflict',
        expectedMessage:
          'Opaque XLSX content requires acknowledgement and a proven independent closure',
        options: { securityMode: 'preserve-opaque' },
        overrides: {
          '[Content_Types].xml': activeTypes.replace(
            '/xl/embeddings/item.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject',
            '/xl/opaque/item.bin" ContentType="application/octet-stream',
          ),
          'xl/opaque/item.bin': new Uint8Array([1, 2, 3]),
        },
      },
      {
        expected: 'opaque-content-conflict',
        overrides: { 'xl/custom.xml': '<custom/>' },
      },
      {
        expected: 'opaque-content-conflict',
        overrides: {
          'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="custom" Type="${XLSX_OFFICE_REL_TYPE}customXml" Target="https://example.invalid/data" TargetMode="External"/></Relationships>`,
        },
      },
    ];
    for (const item of cases) {
      const snapshot = await readXlsxRoundTrip(
        await createIndependentXlsx(item.overrides),
        item.options,
      );
      const edited = await applyXlsxEdits(snapshot, [operation(snapshot)]);
      const error = await capture(() =>
        writeXlsxRoundTrip(edited, item.writeOptions),
      );
      expect(error.diagnostic.code).toBe(item.expected);
      if (item.expectedMessage !== undefined) {
        expect(error.diagnostic.message).toBe(item.expectedMessage);
      }
    }
  });
});
