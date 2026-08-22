import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  parseXlsxWithDiagnostics,
  XlsxParseError,
} from '../../src/formats/xlsx';
import {
  createIndependentXlsx,
  independentWorkbook,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

function arrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
}

async function createChartWorkbook(
  chart: string | null,
  worksheet: string | null | undefined = undefined,
): Promise<Uint8Array> {
  return createIndependentXlsx({
    '[Content_Types].xml': `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/chartsheets/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    </Types>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
      <Relationship Id="rIdSheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rIdChart1" Type="${XLSX_OFFICE_REL_TYPE}chartsheet" Target="chartsheets/chart1.xml"/>
      <Relationship Id="rIdStyles" Type="${XLSX_OFFICE_REL_TYPE}styles" Target="styles.xml"/>
      <Relationship Id="rIdSharedStrings" Type="${XLSX_OFFICE_REL_TYPE}sharedStrings" Target="sharedStrings.xml"/>
    </Relationships>`,
    'xl/chartsheets/chart1.xml': chart,
    'xl/workbook.xml': independentWorkbook(`
      <sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/>
      <sheet name="Chart" sheetId="2" state="hidden" r:id="rIdChart1"/>`),
    ...(worksheet === undefined
      ? {}
      : { 'xl/worksheets/sheet1.xml': worksheet }),
  });
}

describe('public XLSX parser', () => {
  it('returns deterministic portable workbook metadata', async () => {
    const bytes = await createIndependentXlsx();
    const result = await parseXlsxWithDiagnostics(bytes);

    expect(result).toEqual({
      diagnostics: [],
      document: {
        differentialStyles: [],
        namedStyles: [
          {
            builtinId: 0,
            name: 'Normal',
            style: { font: { name: 'Calibri', size: 11 } },
          },
        ],
        sheets: [
          {
            columns: [],
            comments: [],
            conditionalFormattings: [],
            dataValidations: [],
            protectedRanges: [],
            declaredDimension: {
              end: { column: 3, row: 3 },
              reference: 'A1:C3',
              start: { column: 1, row: 1 },
            },
            drawings: [],
            hyperlinks: [],
            index: 0,
            kind: 'worksheet',
            mergedRanges: [],
            name: 'Sheet1',
            payload: 'full-sheet',
            rows: [
              {
                cells: [
                  {
                    address: 'A1',
                    column: 1,
                    content: {
                      kind: 'value',
                      value: { kind: 'text', text: 'Black box' },
                    },
                  },
                ],
                index: 1,
              },
              {
                cells: [
                  {
                    address: 'B2',
                    column: 2,
                    content: {
                      kind: 'value',
                      value: { kind: 'number', value: 42 },
                    },
                  },
                ],
                index: 2,
              },
              {
                cells: [
                  {
                    address: 'C3',
                    column: 3,
                    content: {
                      kind: 'value',
                      value: { kind: 'boolean', value: true },
                    },
                  },
                ],
                index: 3,
              },
            ],
            state: 'visible',
            tables: [],
            views: [],
          },
        ],
        styles: [{ font: { name: 'Calibri', size: 11 } }],
        workbook: {
          calculation: {
            calculationCompleted: true,
            calculateOnSave: true,
            concurrentCalculation: true,
            forceFullCalculation: false,
            fullCalculationOnLoad: false,
            fullPrecision: true,
            iteration: {
              enabled: false,
              maxChange: 0.001,
              maxIterations: 100,
            },
            mode: 'automatic',
            referenceMode: 'A1',
          },
          dateSystem: '1900',
          definedNames: [],
          views: [
            {
              activeSheetIndex: 0,
              autoFilterDateGrouping: true,
              firstVisibleSheetIndex: 0,
              minimized: false,
              showHorizontalScroll: true,
              showSheetTabs: true,
              showVerticalScroll: true,
              tabRatio: 600,
              visibility: 'visible',
            },
          ],
        },
      },
    });
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });

  it('returns normalized cell styles and serial dates without exposing raw XF indexes', async () => {
    const workbook = independentWorkbook(
      '<sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/>',
    ).replace('date1904="0"', 'date1904="1"');
    const bytes = await createIndependentXlsx({
      'xl/styles.xml': `<styleSheet xmlns="${XLSX_SPREADSHEET_NS}">
        <fonts count="1"><font/></fonts>
        <fills count="1"><fill/></fills>
        <borders count="1"><border/></borders>
        <cellStyleXfs count="1"><xf/></cellStyleXfs>
        <cellXfs count="3">
          <xf/>
          <xf numFmtId="14"/>
          <xf numFmtId="14"/>
        </cellXfs>
        <dxfs count="1"><dxf><font><b/><color rgb="FFFF0000"/></font></dxf></dxfs>
      </styleSheet>`,
      'xl/workbook.xml': workbook,
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}">
        <sheetPr><tabColor rgb="FF00AABB"/><outlinePr summaryBelow="0"/></sheetPr>
        <dimension ref="A1:C1"/>
        <cols><col min="1" max="2" width="14" hidden="1" style="1"/></cols>
        <sheetViews><sheetView workbookViewId="0" rightToLeft="1" zoomScale="125">
          <pane xSplit="1" state="frozen" topLeftCell="B1"/>
          <selection pane="topRight" activeCell="B1" sqref="B1:B2"/>
        </sheetView></sheetViews>
        <sheetFormatPr defaultRowHeight="18" defaultColWidth="12"/>
        <sheetData><row s="1" customFormat="1" collapsed="1">
          <c r="A1" s="1"><v>0</v></c>
          <c r="B1" s="2"><f>A1+1</f><v>1</v></c>
          <c r="C1"><v>2</v></c>
        </row></sheetData>
        <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
      </worksheet>`,
    });

    const document = await parseXlsx(bytes);

    expect(document.styles).toEqual([{}, { numberFormat: 'mm-dd-yy' }]);
    expect(document.differentialStyles).toEqual([
      {
        font: {
          bold: true,
          color: { argb: 'FFFF0000', kind: 'rgb' },
        },
      },
    ]);
    expect(document.sheets[0]).toMatchObject({
      columns: [{ end: 2, hidden: true, start: 1, style: 1, width: 14 }],
      declaredDimension: {
        end: { column: 3, row: 1 },
        reference: 'A1:C1',
        start: { column: 1, row: 1 },
      },
      mergedRanges: [
        {
          end: { column: 2, row: 1 },
          reference: 'A1:B1',
          start: { column: 1, row: 1 },
        },
      ],
      outline: {
        applyStyles: false,
        showOutlineSymbols: true,
        summaryBelow: false,
        summaryRight: true,
      },
      rows: [
        {
          cells: [
            {
              address: 'A1',
              content: {
                kind: 'value',
                value: {
                  kind: 'date',
                  normalized: '1904-01-01',
                  precision: 'date',
                  source: {
                    dateSystem: '1904',
                    kind: 'serial',
                    value: 0,
                  },
                },
              },
              style: 1,
            },
            {
              address: 'B1',
              content: {
                cached: {
                  kind: 'date',
                  normalized: '1904-01-02',
                  precision: 'date',
                  source: {
                    dateSystem: '1904',
                    kind: 'serial',
                    value: 1,
                  },
                },
                formula: { expression: 'A1+1', kind: 'normal' },
                kind: 'formula',
              },
              style: 1,
            },
            {
              address: 'C1',
              content: {
                kind: 'value',
                value: { kind: 'number', value: 2 },
              },
            },
          ],
          collapsed: true,
          style: 1,
        },
      ],
      sheetFormat: {
        baseColumnWidth: 8,
        customHeight: false,
        defaultColumnWidth: 12,
        defaultRowHeight: 18,
        outlineColumnLevel: 0,
        outlineRowLevel: 0,
        thickBottom: false,
        thickTop: false,
        zeroHeight: false,
      },
      tabColor: { argb: 'FF00AABB', kind: 'rgb' },
      views: [
        {
          kind: 'normal',
          pane: {
            activePane: 'top-right',
            state: 'frozen',
            topLeftCell: 'B1',
            xSplit: 1,
            ySplit: 0,
          },
          rightToLeft: true,
          selections: [
            {
              activeCell: 'B1',
              pane: 'top-right',
              ranges: [
                {
                  end: { column: 2, row: 2 },
                  reference: 'B1:B2',
                  start: { column: 2, row: 1 },
                },
              ],
            },
          ],
          showGridLines: true,
          showRowColumnHeaders: true,
          tabSelected: false,
          workbookViewId: 0,
          zoomScale: 125,
        },
      ],
    });
    expect(JSON.parse(JSON.stringify(document))).toEqual(document);
  });

  it('maps style-table limits to the public structured error', async () => {
    const bytes = await createIndependentXlsx({
      'xl/styles.xml': `<styleSheet xmlns="${XLSX_SPREADSHEET_NS}">
        <fonts count="1"><font/></fonts>
        <fills count="1"><fill/></fills>
        <borders count="1"><border/></borders>
        <cellStyleXfs count="1"><xf/></cellStyleXfs>
        <cellXfs count="2"><xf/><xf/></cellXfs>
      </styleSheet>`,
    });

    await expect(
      parseXlsx(bytes, { limits: { maxStyles: 1 } }),
    ).rejects.toMatchObject({
      cause: {
        actual: 6,
        limit: 1,
        limitName: 'maxStyles',
        name: 'XlsxResourceLimitError',
        part: 'xl/styles.xml',
      },
      diagnostic: {
        actual: 6,
        code: 'resource-limit-exceeded',
        limit: 1,
        limitName: 'maxStyles',
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });

  it('validates worksheet view references against workbook windows', async () => {
    const bytes = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}">
        <sheetViews><sheetView workbookViewId="1"/></sheetViews><sheetData/>
      </worksheet>`,
    });

    await expect(parseXlsx(bytes)).rejects.toMatchObject({
      diagnostic: {
        code: 'invalid-document-value',
        message: 'Worksheet workbook view reference is out of range',
        part: 'xl/worksheets/sheet1.xml',
      },
      name: 'XlsxParseError',
    });
  });

  it('returns safe worksheet hyperlinks without exposing relationship IDs', async () => {
    const worksheet = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_TYPE.slice(0, -1)}">
      <sheetData/>
      <hyperlinks>
        <hyperlink ref="A1" location="Sheet1!B2" display="Internal"/>
        <hyperlink ref="C3:D4" r:id="external" tooltip="Website"/>
      </hyperlinks>
    </worksheet>`;
    const relationshipPart = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
      <Relationship Id="external" Type="${XLSX_OFFICE_REL_TYPE}hyperlink"
        Target="https://user:secret@example.com/path" TargetMode="External"/>
    </Relationships>`;
    const bytes = await createIndependentXlsx({
      'xl/worksheets/_rels/sheet1.xml.rels': relationshipPart,
      'xl/worksheets/sheet1.xml': worksheet,
    });

    const document = await parseXlsx(bytes);
    const sheet = document.sheets[0]!;
    expect(sheet).toMatchObject({
      hyperlinks: [
        {
          display: 'Internal',
          range: { reference: 'A1' },
          selectionRelation: 'full-sheet',
          target: { kind: 'internal', location: 'Sheet1!B2' },
        },
        {
          range: { reference: 'C3:D4' },
          selectionRelation: 'full-sheet',
          target: { kind: 'external', url: 'https://example.com/path' },
          tooltip: 'Website',
        },
      ],
    });
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect('id' in sheet.hyperlinks[1]!.target).toBe(false);
    expect(JSON.stringify(document)).not.toContain('secret');

    const unsafe = await createIndependentXlsx({
      'xl/worksheets/_rels/sheet1.xml.rels': relationshipPart.replace(
        'https://user:secret@example.com/path',
        'javascript:alert(1)',
      ),
      'xl/worksheets/sheet1.xml': worksheet,
    });
    await expect(parseXlsx(unsafe)).rejects.toMatchObject({
      diagnostic: {
        code: 'security-rejected-content',
        relationshipType: 'hyperlink',
      },
      name: 'XlsxParseError',
    });
  });

  it('returns portable defined names and shares their formula and text budgets', async () => {
    const workbook = independentWorkbook(
      '<sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/>',
    ).replace(
      '<calcPr',
      `<definedNames>
        <definedName name="Sum" comment="note">ABC</definedName>
        <definedName name="Local" localSheetId="0">Sheet1!$A$1</definedName>
      </definedNames><calcPr`,
    );
    const bytes = await createIndependentXlsx({
      'xl/workbook.xml': workbook,
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}">
        <sheetData><row><c><f>DEF</f><v>1</v></c></row></sheetData>
      </worksheet>`,
    });

    const document = await parseXlsx(bytes);
    expect(document.workbook.definedNames).toEqual([
      {
        comment: 'note',
        expression: 'ABC',
        hidden: false,
        name: 'Sum',
      },
      {
        expression: 'Sheet1!$A$1',
        hidden: false,
        name: 'Local',
        sheetIndex: 0,
      },
    ]);
    expect(JSON.parse(JSON.stringify(document.workbook.definedNames))).toEqual(
      document.workbook.definedNames,
    );

    await expect(
      parseXlsx(bytes, {
        limits: {
          maxFormulaCharacters: 13,
          maxTotalFormulaCharacters: 16,
        },
      }),
    ).rejects.toMatchObject({
      cause: {
        actual: 17,
        limit: 16,
        limitName: 'maxTotalFormulaCharacters',
        name: 'XlsxResourceLimitError',
      },
      diagnostic: {
        code: 'resource-limit-exceeded',
        limitName: 'maxTotalFormulaCharacters',
      },
      name: 'XlsxParseError',
    });

    await expect(
      parseXlsx(bytes, {
        limits: { maxTextCharacters: 21 },
        selection: {},
      }),
    ).resolves.toMatchObject({ workbook: { definedNames: [{}, {}] } });
    await expect(
      parseXlsx(bytes, {
        limits: { maxTextCharacters: 20 },
        selection: {},
      }),
    ).rejects.toMatchObject({
      cause: {
        actual: 21,
        limit: 20,
        limitName: 'maxTextCharacters',
        name: 'XlsxResourceLimitError',
      },
      diagnostic: {
        code: 'resource-limit-exceeded',
        limitName: 'maxTextCharacters',
      },
      name: 'XlsxParseError',
    });
  });

  it('treats ArrayBuffer, Uint8Array, subarray, and Blob inputs identically', async () => {
    const bytes = await createIndependentXlsx();
    const padded = new Uint8Array(bytes.byteLength + 2);
    padded.set(bytes, 1);
    const inputs = [
      bytes,
      padded.subarray(1, padded.byteLength - 1),
      arrayBuffer(bytes),
      new Blob([arrayBuffer(bytes)]),
    ];

    const outputs = await Promise.all(inputs.map((input) => parseXlsx(input)));
    for (const output of outputs) expect(output).toEqual(outputs[0]);
  });

  it('does not mutate caller bytes or options', async () => {
    const bytes = await createIndependentXlsx();
    const before = bytes.slice();
    const limits = { maxInputBytes: bytes.byteLength };
    const options = { limits } as const;

    await parseXlsx(bytes, options);

    expect(bytes).toEqual(before);
    expect(options).toEqual({ limits: { maxInputBytes: bytes.byteLength } });
  });

  it('isolates concurrent parses and returns literal-equal results', async () => {
    const bytes = await createIndependentXlsx();
    const outputs = await Promise.all(
      Array.from({ length: 12 }, () => parseXlsx(bytes)),
    );

    for (const output of outputs) expect(output).toEqual(outputs[0]);
    expect(new Set(outputs).size).toBe(outputs.length);
    expect(new Set(outputs.map((output) => output.sheets)).size).toBe(
      outputs.length,
    );
  });

  it('returns the document-only convenience result', async () => {
    const bytes = await createIndependentXlsx();

    await expect(parseXlsx(bytes)).resolves.toEqual(
      (await parseXlsxWithDiagnostics(bytes)).document,
    );
  });

  it.each(['tolerant', 'strict'] as const)(
    'rejects invalid ZIP data with a structured error in %s mode',
    async (errorMode) => {
      const action = parseXlsx(new Uint8Array([1, 2, 3, 4]), { errorMode });
      await expect(action).rejects.toMatchObject({
        diagnostic: {
          code: 'invalid-package',
          message: 'Failed to open XLSX OPC package',
          severity: 'error',
        },
        name: 'XlsxParseError',
      });
      const error = await action.catch((reason: unknown) => reason);
      expect(error).toBeInstanceOf(XlsxParseError);
      expect((error as XlsxParseError).cause).toBeInstanceOf(Error);
    },
  );

  it.each(['tolerant', 'strict'] as const)(
    'enforces input limits with structured metadata in %s mode',
    async (errorMode) => {
      const bytes = await createIndependentXlsx();
      await expect(
        parseXlsx(bytes, {
          errorMode,
          limits: { maxInputBytes: bytes.byteLength - 1 },
        }),
      ).rejects.toMatchObject({
        cause: {
          actual: bytes.byteLength,
          limit: bytes.byteLength - 1,
          limitName: 'maxInputBytes',
          name: 'XlsxResourceLimitError',
        },
        diagnostic: {
          actual: bytes.byteLength,
          code: 'resource-limit-exceeded',
          limit: bytes.byteLength - 1,
          limitName: 'maxInputBytes',
          severity: 'error',
        },
        name: 'XlsxParseError',
      });
    },
  );

  it('maps archive preflight limits to the public structured error', async () => {
    const bytes = await createIndependentXlsx();

    await expect(
      parseXlsx(bytes, { limits: { maxEntries: 1 } }),
    ).rejects.toMatchObject({
      cause: {
        limit: 1,
        limitName: 'maxEntries',
        name: 'XlsxResourceLimitError',
      },
      diagnostic: {
        code: 'resource-limit-exceeded',
        limit: 1,
        limitName: 'maxEntries',
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });

  it('maps streamed workbook-table limits to the public structured error', async () => {
    const bytes = await createIndependentXlsx({
      'xl/sharedStrings.xml': `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si/><si/></sst>`,
    });

    await expect(
      parseXlsx(bytes, { limits: { maxSharedStrings: 1 } }),
    ).rejects.toMatchObject({
      cause: {
        actual: 2,
        limit: 1,
        limitName: 'maxSharedStrings',
        name: 'XlsxResourceLimitError',
        part: 'xl/sharedStrings.xml',
      },
      diagnostic: {
        actual: 2,
        code: 'resource-limit-exceeded',
        limit: 1,
        limitName: 'maxSharedStrings',
        part: 'xl/sharedStrings.xml',
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });

  it('rejects malformed shared strings through the public parser', async () => {
    const bytes = await createIndependentXlsx({
      'xl/sharedStrings.xml': `<sst xmlns="urn:wrong"><si><t>bad</t></si></sst>`,
    });

    await expect(parseXlsx(bytes)).rejects.toMatchObject({
      diagnostic: {
        code: 'invalid-document-structure',
        message: 'Shared-string element has an unsupported namespace',
        part: 'xl/sharedStrings.xml',
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });

  it.each([
    [
      null,
      'missing-required-part',
      'Required XLSX part is missing: xl/worksheets/sheet1.xml',
    ],
    [
      '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
      'invalid-document-structure',
      'Worksheet root is missing',
    ],
    [
      '<worksheet xmlns="urn:wrong"><sheetData/></worksheet>',
      'invalid-document-structure',
      'Worksheet element has an unsupported namespace',
    ],
    [
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
      'invalid-document-structure',
      'Worksheet sheetData is missing',
    ],
  ] as const)(
    'validates streamed worksheet payload %#',
    async (worksheetXml, code, message) => {
      const bytes = await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': worksheetXml,
      });
      await expect(parseXlsx(bytes)).rejects.toMatchObject({
        diagnostic: {
          code,
          message,
          part: 'xl/worksheets/sheet1.xml',
          severity: 'error',
        },
        name: 'XlsxParseError',
      });
    },
  );

  it('maps aggregate streamed-cell limits to structured public errors', async () => {
    const bytes = await createIndependentXlsx();
    await expect(
      parseXlsx(bytes, { limits: { maxReturnedCells: 2 } }),
    ).rejects.toMatchObject({
      cause: {
        actual: 3,
        limit: 2,
        limitName: 'maxReturnedCells',
        name: 'XlsxResourceLimitError',
        part: 'xl/worksheets/sheet1.xml',
      },
      diagnostic: {
        actual: 3,
        code: 'resource-limit-exceeded',
        limit: 2,
        limitName: 'maxReturnedCells',
        part: 'xl/worksheets/sheet1.xml',
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });

  it('counts serialized shared-string expansion in the text budget', async () => {
    const bytes = await createIndependentXlsx();
    await expect(
      parseXlsx(bytes, { limits: { maxTextCharacters: 18 } }),
    ).resolves.toMatchObject({ sheets: [{ name: 'Sheet1' }] });
    await expect(
      parseXlsx(bytes, { limits: { maxTextCharacters: 17 } }),
    ).rejects.toMatchObject({
      cause: {
        actual: 18,
        limit: 17,
        limitName: 'maxTextCharacters',
        name: 'XlsxResourceLimitError',
      },
      diagnostic: {
        actual: 18,
        code: 'resource-limit-exceeded',
        limit: 17,
        limitName: 'maxTextCharacters',
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });

  it('preserves worksheet cell diagnostics through the public parser', async () => {
    const bytes = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': `
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData><row><c t="b"><v>2</v></c></row></sheetData>
        </worksheet>`,
    });
    await expect(parseXlsx(bytes)).rejects.toMatchObject({
      diagnostic: {
        cell: 'A1',
        code: 'invalid-document-value',
        message: 'Cell boolean is invalid',
        part: 'xl/worksheets/sheet1.xml',
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });

  it('returns shared formula expressions and cached state through the public API', async () => {
    const bytes = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': `
        <worksheet xmlns="${XLSX_SPREADSHEET_NS}">
          <sheetData><row>
            <c r="A1"><f t="shared" si="0" ref="A1:B1">A1+1</f><v>2</v></c>
            <c r="B1"><f t="shared" si="0"/></c>
          </row></sheetData>
        </worksheet>`,
    });

    await expect(parseXlsx(bytes)).resolves.toMatchObject({
      sheets: [
        {
          rows: [
            {
              cells: [
                {
                  address: 'A1',
                  content: {
                    cached: { kind: 'number', value: 2 },
                    formula: { expression: 'A1+1', kind: 'normal' },
                    kind: 'formula',
                  },
                },
                {
                  address: 'B1',
                  content: {
                    cached: { kind: 'missing' },
                    formula: { expression: 'B1+1', kind: 'normal' },
                    kind: 'formula',
                  },
                },
              ],
            },
          ],
        },
      ],
    });
  });

  it('returns sparse selected ranges with explicit payload metadata', async () => {
    const bytes = await createIndependentXlsx();
    const document = await parseXlsx(bytes, {
      selection: { ranges: { sheet1: ['B2:C3'] } },
    });

    expect(document.sheets).toEqual([
      {
        columns: [],
        comments: [],
        conditionalFormattings: [],
        dataValidations: [],
        protectedRanges: [],
        declaredDimension: {
          end: { column: 3, row: 3 },
          reference: 'A1:C3',
          start: { column: 1, row: 1 },
        },
        drawings: [],
        hyperlinks: [],
        index: 0,
        kind: 'worksheet',
        mergedRanges: [],
        name: 'Sheet1',
        payload: 'selected-ranges',
        rows: [
          {
            cells: [
              {
                address: 'B2',
                column: 2,
                content: {
                  kind: 'value',
                  value: { kind: 'number', value: 42 },
                },
              },
            ],
            index: 2,
          },
          {
            cells: [
              {
                address: 'C3',
                column: 3,
                content: {
                  kind: 'value',
                  value: { kind: 'boolean', value: true },
                },
              },
            ],
            index: 3,
          },
        ],
        state: 'visible',
        tables: [],
        views: [],
      },
    ]);
  });

  it('lets a whole-sheet selection win over validated ranges', async () => {
    const bytes = await createIndependentXlsx();
    const selected = await parseXlsx(bytes, {
      selection: {
        ranges: { SHEET1: ['B2'] },
        sheetNames: ['sheet1'],
      },
    });

    expect(selected.sheets[0]).toMatchObject({
      payload: 'full-sheet',
      rows: [{ index: 1 }, { index: 2 }, { index: 3 }],
    });
  });

  it('preserves unselected sheet metadata without reading its worksheet part', async () => {
    const bytes = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': null,
    });

    await expect(parseXlsx(bytes, { selection: {} })).resolves.toMatchObject({
      sheets: [
        {
          kind: 'worksheet',
          name: 'Sheet1',
          payload: 'not-selected',
          rows: [],
        },
      ],
    });
  });

  it('reads only selected worksheet and chart-sheet payloads', async () => {
    const chart = `<chartsheet xmlns="${XLSX_SPREADSHEET_NS}"/>`;
    const chartOnly = await createChartWorkbook(chart, null);
    await expect(
      parseXlsx(chartOnly, { selection: { sheetNames: ['chart'] } }),
    ).resolves.toMatchObject({
      sheets: [
        {
          kind: 'worksheet',
          name: 'Sheet1',
          payload: 'not-selected',
          rows: [],
        },
        { kind: 'chart-sheet', name: 'Chart', payload: 'full-sheet' },
      ],
    });

    const noPayloads = await createChartWorkbook(null, null);
    await expect(
      parseXlsx(noPayloads, { selection: {} }),
    ).resolves.toMatchObject({
      sheets: [
        { kind: 'worksheet', payload: 'not-selected', rows: [] },
        { kind: 'chart-sheet', payload: 'not-selected' },
      ],
    });

    const missingChart = await createChartWorkbook(null);
    await expect(parseXlsx(missingChart)).rejects.toMatchObject({
      diagnostic: {
        code: 'missing-required-part',
        message: 'Required XLSX part is missing: xl/chartsheets/chart1.xml',
        part: 'xl/chartsheets/chart1.xml',
      },
      name: 'XlsxParseError',
    });
  });

  it('rejects range selection on chart sheets before reading their payload', async () => {
    const bytes = await createChartWorkbook(null, null);
    await expect(
      parseXlsx(bytes, { selection: { ranges: { Chart: ['A1'] } } }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'invalid-selection',
        message: 'XLSX selection ranges require a worksheet',
        severity: 'error',
      },
      name: 'XlsxParseError',
    });
  });

  it.each([
    [{ sheetNames: ['Missing'] }, 'XLSX selection references an unknown sheet'],
    [
      { ranges: { Sheet1: ['Sheet1!A1'] } },
      'XLSX selection contains an invalid range',
    ],
  ] as const)(
    'reports invalid public selection %#',
    async (selection, message) => {
      const bytes = await createIndependentXlsx();
      await expect(parseXlsx(bytes, { selection })).rejects.toMatchObject({
        diagnostic: {
          code: 'invalid-selection',
          message,
          severity: 'error',
        },
        name: 'XlsxParseError',
      });
    },
  );

  it('accounts scanned and returned cells independently through the public parser', async () => {
    const bytes = await createIndependentXlsx();
    const selection = { ranges: { Sheet1: ['B2'] } } as const;

    await expect(
      parseXlsx(bytes, {
        limits: { maxReturnedCells: 1, maxScannedCells: 3 },
        selection,
      }),
    ).resolves.toMatchObject({
      sheets: [{ payload: 'selected-ranges', rows: [{ index: 2 }] }],
    });
    await expect(
      parseXlsx(bytes, {
        limits: { maxReturnedCells: 1, maxScannedCells: 2 },
        selection,
      }),
    ).rejects.toMatchObject({
      cause: {
        actual: 3,
        limit: 2,
        limitName: 'maxScannedCells',
        part: 'xl/worksheets/sheet1.xml',
      },
      diagnostic: {
        actual: 3,
        code: 'resource-limit-exceeded',
        limit: 2,
        limitName: 'maxScannedCells',
        part: 'xl/worksheets/sheet1.xml',
      },
      name: 'XlsxParseError',
    });
    await expect(
      parseXlsx(bytes, {
        limits: { maxReturnedCells: 1, maxScannedCells: 3 },
        selection: { ranges: { Sheet1: ['B2:C3'] } },
      }),
    ).rejects.toMatchObject({
      cause: {
        actual: 2,
        limit: 1,
        limitName: 'maxReturnedCells',
        part: 'xl/worksheets/sheet1.xml',
      },
      diagnostic: {
        actual: 2,
        code: 'resource-limit-exceeded',
        limit: 1,
        limitName: 'maxReturnedCells',
        part: 'xl/worksheets/sheet1.xml',
      },
      name: 'XlsxParseError',
    });
  });

  it('validates unreturned worksheet cells before discarding their values', async () => {
    const bytes = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': `
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData><row><c><v>1</v></c><c t="b"><v>2</v></c></row></sheetData>
        </worksheet>`,
    });

    await expect(
      parseXlsx(bytes, {
        selection: { ranges: { Sheet1: ['A1'] } },
      }),
    ).rejects.toMatchObject({
      diagnostic: {
        cell: 'B1',
        code: 'invalid-document-value',
        message: 'Cell boolean is invalid',
        part: 'xl/worksheets/sheet1.xml',
      },
      name: 'XlsxParseError',
    });
  });

  it.each([
    { errorMode: 'strict' },
    { errorMode: 'tolerant' },
    { displayTextMode: 'none' },
    { displayTextMode: 'supported' },
    { imageMode: 'base64' },
    { imageMode: 'blob' },
    { imageMode: 'both' },
    { imageMode: 'none' },
    { pivotCacheMode: 'metadata' },
    { pivotCacheMode: 'none' },
    { pivotCacheMode: 'records' },
  ] as const)('accepts valid runtime option %#', async (options) => {
    const bytes = await createIndependentXlsx();
    await expect(parseXlsx(bytes, options)).resolves.toMatchObject({
      sheets: [{ name: 'Sheet1' }],
    });
  });

  it.each([
    ['errorMode', 'recover', 'XLSX errorMode is invalid'],
    ['displayTextMode', 'all', 'XLSX displayTextMode is invalid'],
    ['imageMode', 'url', 'XLSX imageMode is invalid'],
    ['pivotCacheMode', 'all', 'XLSX pivotCacheMode is invalid'],
  ] as const)(
    'rejects invalid runtime option %s',
    async (name, value, message) => {
      const bytes = await createIndependentXlsx();
      await expect(parseXlsx(bytes, { [name]: value })).rejects.toThrow(
        message,
      );
    },
  );
});
