import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
  XlsxParseError,
} from '../../src/formats/xlsx';
import type { XlsxBlackBoxOverrides } from '../black-box/xlsx-package';
import {
  createIndependentXlsx,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const TABLE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
const TABLE_PART = 'xl/tables/table1.xml';
const STYLES_WITH_DXF = `<styleSheet xmlns="${XLSX_SPREADSHEET_NS}">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="1"><dxf><font><b/></font></dxf></dxfs>
</styleSheet>`;

const CONTENT_TYPES = `<?xml version="1.0"?><Types xmlns="${XLSX_CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/${TABLE_PART}" ContentType="${TABLE_CONTENT_TYPE}"/>
</Types>`;

function worksheet(
  tableParts = '<tableParts count="1"><tablePart r:id="table"/></tableParts>',
): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}">
    <sheetData>
      <row r="1"><c r="A1" t="str"><v>Product</v></c><c r="B1" t="str"><v>Amount</v></c><c r="C1" t="str"><v>Tax</v></c></row>
      <row r="2"><c r="A2" t="str"><v>One</v></c><c r="B2"><v>10</v></c><c r="C2"><v>1</v></c></row>
      <row r="3"><c r="A3" t="str"><v>Two</v></c><c r="B3"><v>20</v></c><c r="C3"><v>2</v></c></row>
      <row r="4"><c r="A4" t="str"><v>Total</v></c><c r="B4"><v>30</v></c><c r="C4"><v>3</v></c></row>
    </sheetData>
    ${tableParts}
  </worksheet>`;
}

function tableXml(overrides = ''): string {
  return `<table xmlns="${XLSX_SPREADSHEET_NS}" id="1" name="Sales_Internal" displayName="Sales" ref="A1:C4" totalsRowCount="1" totalsRowShown="1" insertRow="1" published="1" ${overrides}>
    <tableColumns count="3">
      <tableColumn id="1" name="Product" totalsRowLabel="Total"/>
      <tableColumn id="2" name="Amount" totalsRowFunction="sum">
        <calculatedColumnFormula array="1">SUBTOTAL(109,[Amount])</calculatedColumnFormula>
      </tableColumn>
      <tableColumn id="3" name="Tax" totalsRowFunction="custom">
        <totalsRowFormula>SUM([Tax])</totalsRowFormula>
      </tableColumn>
    </tableColumns>
    <tableStyleInfo name="TableStyleMedium2" showFirstColumn="1" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
  </table>`;
}

function relationship(
  target = '../tables/table1.xml',
  targetMode = '',
  type = `${XLSX_OFFICE_REL_TYPE}table`,
): string {
  return `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="table" Type="${type}" Target="${target}" ${targetMode}/></Relationships>`;
}

async function createTableWorkbook(
  overrides: XlsxBlackBoxOverrides = {},
): Promise<Uint8Array> {
  return createIndependentXlsx({
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/worksheets/_rels/sheet1.xml.rels': relationship(),
    'xl/worksheets/sheet1.xml': worksheet(),
    [TABLE_PART]: tableXml(),
    ...overrides,
  });
}

async function capture(
  overrides: XlsxBlackBoxOverrides,
  options: Parameters<typeof parseXlsx>[1] = {},
): Promise<XlsxParseError> {
  try {
    await parseXlsx(await createTableWorkbook(overrides), options);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected XLSX table parsing to fail');
}

describe('XLSX tables', () => {
  it('parses complete table columns, formulas, totals, and style metadata', async () => {
    const document = await parseXlsx(await createTableWorkbook());
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    expect(sheet.kind === 'worksheet' ? sheet.tables : []).toEqual([
      {
        columns: [
          {
            id: 1,
            name: 'Product',
            totalsFunction: 'none',
            totalsLabel: 'Total',
          },
          {
            calculatedFormula: {
              array: true,
              expression: 'SUBTOTAL(109,[Amount])',
            },
            id: 2,
            name: 'Amount',
            totalsFunction: 'sum',
          },
          {
            id: 3,
            name: 'Tax',
            totalsFormula: { array: false, expression: 'SUM([Tax])' },
            totalsFunction: 'custom',
          },
        ],
        displayName: 'Sales',
        headerRow: true,
        id: 1,
        insertRow: true,
        insertRowShift: false,
        name: 'Sales_Internal',
        published: true,
        range: {
          end: { column: 3, row: 4 },
          reference: 'A1:C4',
          start: { column: 1, row: 1 },
        },
        selectionRelation: 'full-sheet',
        style: {
          name: 'TableStyleMedium2',
          showColumnStripes: false,
          showFirstColumn: true,
          showLastColumn: false,
          showRowStripes: true,
        },
        tableType: 'worksheet',
        totalsRow: true,
        totalsRowShown: true,
      },
    ]);
    expect(JSON.parse(JSON.stringify(document))).toEqual(document);
  });

  it('preserves table and column connection, style, and identity metadata', async () => {
    const xml = tableXml(
      'comment="Quarter &amp; region" connectionId="0" dataCellStyle="Data Style" dataDxfId="0" headerRowBorderDxfId="0" headerRowCellStyle="Header Style" headerRowDxfId="0" insertRowShift="true" tableBorderDxfId="0" tableType="queryTable" totalsRowBorderDxfId="0" totalsRowCellStyle="Totals Style" totalsRowDxfId="0"',
    ).replace(
      'id="1" name="Product"',
      'id="1" name="Product" uniqueName="Product Unique" queryTableFieldId="0" headerRowCellStyle="Column Header" dataCellStyle="Column Data" totalsRowCellStyle="Column Totals" headerRowDxfId="0" dataDxfId="0" totalsRowDxfId="0"',
    );
    const document = await parseXlsx(
      await createTableWorkbook({
        'xl/styles.xml': STYLES_WITH_DXF,
        [TABLE_PART]: xml,
      }),
    );
    const sheet = document.sheets[0]!;
    const table = sheet.kind === 'worksheet' ? sheet.tables[0] : undefined;
    expect(table).toMatchObject({
      comment: 'Quarter & region',
      connectionId: 0,
      dataCellStyle: 'Data Style',
      dataDifferentialStyle: 0,
      headerCellStyle: 'Header Style',
      headerDifferentialStyle: 0,
      headerRowBorderDifferentialStyle: 0,
      insertRowShift: true,
      tableBorderDifferentialStyle: 0,
      tableType: 'query-table',
      totalsCellStyle: 'Totals Style',
      totalsDifferentialStyle: 0,
      totalsRowBorderDifferentialStyle: 0,
    });
    expect(table?.columns[0]).toMatchObject({
      dataCellStyle: 'Column Data',
      dataDifferentialStyle: 0,
      headerCellStyle: 'Column Header',
      headerDifferentialStyle: 0,
      queryTableFieldId: 0,
      totalsCellStyle: 'Column Totals',
      totalsDifferentialStyle: 0,
      uniqueName: 'Product Unique',
    });
  });

  it('decodes every XML entity form in table-owned text', async () => {
    const xml = `<table xmlns="${XLSX_SPREADSHEET_NS}" id="2" name="Entity_Table" displayName="Entity_Table" comment="&amp;&apos;&gt;&lt;&quot;&#65;&#x1F600;" ref="A1:A1"><tableColumns count="1"><tableColumn id="1" name="&amp;&apos;&gt;&lt;&quot;&#65;&#x1F600;"/></tableColumns></table>`;
    const document = await parseXlsx(
      await createTableWorkbook({ [TABLE_PART]: xml }),
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.tables[0] : undefined,
    ).toMatchObject({
      columns: [{ name: '&\'><"A😀' }],
      comment: '&\'><"A😀',
    });
  });

  it('returns only tables intersecting selected ranges while validating all parts', async () => {
    const bytes = await createTableWorkbook();
    const included = await parseXlsx(bytes, {
      selection: { ranges: { Sheet1: ['B2'] } },
    });
    const includedSheet = included.sheets[0]!;
    expect(
      includedSheet.kind === 'worksheet' ? includedSheet.tables : [],
    ).toMatchObject([{ selectionRelation: 'intersects-selection' }]);
    const excluded = await parseXlsx(bytes, {
      selection: { ranges: { Sheet1: ['D1'] } },
    });
    const excludedSheet = excluded.sheets[0]!;
    expect(
      excludedSheet.kind === 'worksheet' ? excludedSheet.tables : [],
    ).toEqual([]);
  });

  it('parses table-owned auto-filter and sort metadata', async () => {
    const xml = tableXml().replace(
      '<tableColumns count="3">',
      `<autoFilter ref="A1:C3"><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="10"/></customFilters></filterColumn><sortState ref="A2:C3" sortMethod="stroke"><sortCondition ref="B2:B3" descending="1"/></sortState></autoFilter><tableColumns count="3">`,
    );
    const document = await parseXlsx(
      await createTableWorkbook({ [TABLE_PART]: xml }),
    );
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    expect(
      sheet.kind === 'worksheet' ? sheet.tables[0]?.autoFilter : null,
    ).toEqual({
      columns: [
        {
          columnId: 1,
          hiddenButton: false,
          rule: {
            and: false,
            conditions: [{ operator: 'greater-than', value: '10' }],
            kind: 'custom',
          },
          showButton: true,
        },
      ],
      range: {
        end: { column: 3, row: 3 },
        reference: 'A1:C3',
        start: { column: 1, row: 1 },
      },
      selectionRelation: 'full-sheet',
      sort: {
        caseSensitive: false,
        columnSort: false,
        conditions: [
          {
            descending: true,
            range: {
              end: { column: 2, row: 3 },
              reference: 'B2:B3',
              start: { column: 2, row: 2 },
            },
            sortBy: 'value',
          },
        ],
        range: {
          end: { column: 3, row: 3 },
          reference: 'A2:C3',
          start: { column: 1, row: 2 },
        },
        sortMethod: 'stroke',
      },
    });
    const outside = xml.replace('ref="A1:C3"', 'ref="A1:D3"');
    expect((await capture({ [TABLE_PART]: outside })).diagnostic.message).toBe(
      'Table auto-filter range is outside the table',
    );
  });

  it('round-trips a table and filter workbook through portable exact R0', async () => {
    const xml = tableXml().replace(
      '<tableColumns count="3">',
      '<autoFilter ref="A1:C3"><filterColumn colId="1"><top10 val="2"/></filterColumn></autoFilter><tableColumns count="3">',
    );
    const bytes = await createTableWorkbook({ [TABLE_PART]: xml });
    const snapshot = await readXlsxRoundTrip(bytes);
    const portable = await validateXlsxRoundTripJson(
      JSON.parse(JSON.stringify(snapshot)) as unknown,
    );
    const result = await writeXlsxRoundTrip(portable);
    expect(result.data).toEqual(bytes);
    expect(result.report.level).toBe('R0');
    expect(result.report.outputSha256).toBe(result.report.sourceSha256);
    expect(
      result.report.parts.every((part) => part.disposition === 'copy'),
    ).toBe(true);
    expect(snapshot.document.sheets[0]).toMatchObject({
      tables: [{ autoFilter: { columns: [{ columnId: 1 }] } }],
    });
  });

  it('checks every table auto-filter containment boundary', async () => {
    const source = tableXml().replace('ref="A1:C4"', 'ref="B2:D4"');
    const withFilter = (reference: string) =>
      source.replace(
        '<tableColumns count="3">',
        `<autoFilter ref="${reference}"/><tableColumns count="3">`,
      );
    await expect(
      parseXlsx(
        await createTableWorkbook({ [TABLE_PART]: withFilter('B2:D4') }),
      ),
    ).resolves.toBeDefined();
    for (const reference of ['B1:D4', 'B2:D5', 'A2:D4', 'B2:E4']) {
      expect(
        (await capture({ [TABLE_PART]: withFilter(reference) })).diagnostic
          .message,
      ).toBe('Table auto-filter range is outside the table');
    }
    const duplicate = source.replace(
      '<tableColumns count="3">',
      '<autoFilter ref="B2:D4"/><autoFilter ref="B2:D4"/><tableColumns count="3">',
    );
    expect(
      (await capture({ [TABLE_PART]: duplicate })).diagnostic.message,
    ).toBe('Table auto-filter range is outside the table');
  });

  it('accepts the exact table limit and rejects one over', async () => {
    await expect(
      parseXlsx(await createTableWorkbook(), { limits: { maxTables: 1 } }),
    ).resolves.toBeDefined();
    const error = await capture(
      {
        'xl/worksheets/sheet1.xml': worksheet(
          '<tableParts count="2"><tablePart r:id="table"/><tablePart r:id="other"/></tableParts>',
        ),
      },
      { limits: { maxTables: 1 } },
    );
    expect(error.diagnostic).toMatchObject({
      actual: 2,
      code: 'resource-limit-exceeded',
      limit: 1,
      limitName: 'maxTables',
      part: 'xl/worksheets/sheet1.xml',
    });
  });

  it.each([
    [
      '<tableParts count="2"><tablePart r:id="table"/></tableParts>',
      'Worksheet table-part count does not match',
    ],
    [
      '<tableParts count="1"><tablePart r:id="table"/></tableParts><tableParts count="1"><tablePart r:id="table"/></tableParts>',
      'Worksheet contains duplicate tableParts elements',
    ],
    [
      '<tableParts count="01"><tablePart r:id="table"/></tableParts>',
      'Worksheet table-part count is invalid',
    ],
    ['<tableParts count="0"/>', 'Worksheet table-part count is invalid'],
    [
      '<tableParts count="1"><wrong r:id="table"/></tableParts>',
      'Worksheet element nesting is invalid',
    ],
    [
      '<tableParts count="1"><tablePart r:id=""/></tableParts>',
      'Worksheet table relationship reference is invalid',
    ],
    [
      '<tableParts count="2"><tablePart r:id="table"/><tablePart r:id="table"/></tableParts>',
      'Worksheet contains duplicate table relationship references',
    ],
  ] as const)(
    'validates worksheet table-part container %#',
    async (tableParts, message) => {
      const error = await capture({
        'xl/worksheets/sheet1.xml': worksheet(tableParts),
      });
      expect(error.diagnostic.message).toBe(message);
    },
  );

  it('ignores a tablePart element outside its owning container', async () => {
    const document = await parseXlsx(
      await createTableWorkbook({
        'xl/worksheets/sheet1.xml': worksheet('<tablePart r:id="table"/>'),
      }),
    );
    const sheet = document.sheets[0]!;
    expect(sheet.kind === 'worksheet' ? sheet.tables : []).toEqual([]);
  });

  it('rejects a tablePart nested under another owning collection', async () => {
    const error = await capture({
      'xl/worksheets/sheet1.xml': worksheet(
        '<hyperlinks><tablePart r:id="table"/></hyperlinks>',
      ),
    });
    expect(error.diagnostic.message).toBe(
      'Worksheet element nesting is invalid',
    );
  });

  it('enforces the table limit across worksheet owners', async () => {
    const contentTypes = CONTENT_TYPES.replace(
      '</Types>',
      `<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table2.xml" ContentType="${TABLE_CONTENT_TYPE}"/></Types>`,
    );
    const bytes = await createIndependentXlsx({
      '[Content_Types].xml': contentTypes,
      'xl/workbook.xml': `<workbook xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="sheet1"/><sheet name="Sheet2" sheetId="2" r:id="sheet2"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="sheet2" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
      'xl/worksheets/_rels/sheet1.xml.rels': relationship(),
      'xl/worksheets/_rels/sheet2.xml.rels': relationship(
        '../tables/table2.xml',
      ),
      'xl/worksheets/sheet1.xml': worksheet(),
      'xl/worksheets/sheet2.xml': worksheet(),
      'xl/tables/table1.xml': tableXml(),
      'xl/tables/table2.xml': tableXml()
        .replace('id="1"', 'id="2"')
        .replace('name="Sales_Internal"', 'name="Other_Internal"')
        .replace('displayName="Sales"', 'displayName="Other"'),
    });
    await expect(
      parseXlsx(bytes, { limits: { maxTables: 2 } }),
    ).resolves.toBeDefined();
    try {
      await parseXlsx(bytes, { limits: { maxTables: 1 } });
    } catch (error) {
      expect(error).toBeInstanceOf(XlsxParseError);
      expect((error as XlsxParseError).diagnostic).toMatchObject({
        actual: 2,
        limit: 1,
        limitName: 'maxTables',
        part: 'xl/tables/table2.xml',
      });
      return;
    }
    throw new Error('Expected the workbook-global table bound to fail');
  });

  it('enforces workbook-global table IDs, names, and part ownership', async () => {
    const contentTypes = CONTENT_TYPES.replace(
      '</Types>',
      `<Override PartName="/xl/tables/table2.xml" ContentType="${TABLE_CONTENT_TYPE}"/></Types>`,
    );
    const worksheetWithTwo = worksheet(
      '<tableParts count="2"><tablePart r:id="table"/><tablePart r:id="other"/></tableParts>',
    );
    const relationships = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="table" Type="${XLSX_OFFICE_REL_TYPE}table" Target="../tables/table1.xml"/><Relationship Id="other" Type="${XLSX_OFFICE_REL_TYPE}table" Target="../tables/table2.xml"/></Relationships>`;
    const baseOverrides = {
      '[Content_Types].xml': contentTypes,
      'xl/worksheets/_rels/sheet1.xml.rels': relationships,
      'xl/worksheets/sheet1.xml': worksheetWithTwo,
    };
    const duplicateId = await capture({
      ...baseOverrides,
      'xl/tables/table2.xml': tableXml()
        .replace('name="Sales_Internal"', 'name="Other_Internal"')
        .replace('displayName="Sales"', 'displayName="Other"'),
    });
    expect(duplicateId.diagnostic.message).toBe(
      'Workbook contains duplicate table identities',
    );
    const duplicateName = await capture({
      ...baseOverrides,
      'xl/tables/table2.xml': tableXml().replace('id="1"', 'id="2"'),
    });
    expect(duplicateName.diagnostic.message).toBe(
      'Workbook contains duplicate table identities',
    );
    const duplicatePart = await capture({
      ...baseOverrides,
      'xl/worksheets/_rels/sheet1.xml.rels': relationships.replace(
        '../tables/table2.xml',
        '../tables/table1.xml',
      ),
      'xl/tables/table2.xml': tableXml()
        .replace('id="1"', 'id="2"')
        .replace('name="Sales_Internal"', 'name="Other_Internal"')
        .replace('displayName="Sales"', 'displayName="Other"'),
    });
    expect(duplicatePart.diagnostic.message).toBe(
      'Workbook references a table part more than once',
    );
    const unicodeFold = await capture({
      ...baseOverrides,
      [TABLE_PART]: tableXml()
        .replace('name="Sales_Internal"', 'name="Straße"')
        .replace('displayName="Sales"', 'displayName="First_Table"'),
      'xl/tables/table2.xml': tableXml()
        .replace('id="1"', 'id="2"')
        .replace('name="Sales_Internal"', 'name="STRASSE"')
        .replace('displayName="Sales"', 'displayName="Second_Table"'),
    });
    expect(unicodeFold.diagnostic.message).toBe(
      'Workbook contains duplicate table identities',
    );
    const unicodeDisplayFold = await capture({
      ...baseOverrides,
      [TABLE_PART]: tableXml()
        .replace('name="Sales_Internal"', 'name="First_Internal"')
        .replace('displayName="Sales"', 'displayName="Straße"'),
      'xl/tables/table2.xml': tableXml()
        .replace('id="1"', 'id="2"')
        .replace('name="Sales_Internal"', 'name="Second_Internal"')
        .replace('displayName="Sales"', 'displayName="STRASSE"'),
    });
    expect(unicodeDisplayFold.diagnostic.message).toBe(
      'Workbook contains duplicate table identities',
    );
  });

  it('parses Strict table namespaces and relationship types', async () => {
    const strictSheet = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictRelationship =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const bytes = await createIndependentXlsx({
      '[Content_Types].xml': CONTENT_TYPES,
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelationship}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${strictRelationship}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheet}" xmlns:r="${strictRelationship}"><s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets></s:workbook>`,
      'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="table" Type="${strictRelationship}/table" Target="../tables/table1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheet}" xmlns:r="${strictRelationship}"><s:sheetData/><s:tableParts count="1"><s:tablePart r:id="table"/></s:tableParts></s:worksheet>`,
      [TABLE_PART]: `<s:table xmlns:s="${strictSheet}" id="1" name="StrictTable" displayName="StrictTable" ref="A1:A1"><s:tableColumns count="1"><s:tableColumn id="1" name="Value"/></s:tableColumns></s:table>`,
    });
    const document = await parseXlsx(bytes);
    const sheet = document.sheets[0]!;
    expect(sheet.kind === 'worksheet' ? sheet.tables : []).toMatchObject([
      { displayName: 'StrictTable', name: 'StrictTable' },
    ]);
  });

  it('enforces table formula budgets at exact boundaries', async () => {
    const calculated = 'SUBTOTAL(109,[Amount])';
    const totals = 'SUM([Tax])';
    const total = calculated.length + totals.length;
    await expect(
      parseXlsx(await createTableWorkbook(), {
        limits: {
          maxFormulaCharacters: calculated.length,
          maxTotalFormulaCharacters: total,
        },
      }),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(
          {},
          {
            limits: {
              maxFormulaCharacters: calculated.length - 1,
              maxTotalFormulaCharacters: total,
            },
          },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: calculated.length,
      limit: calculated.length - 1,
      limitName: 'maxFormulaCharacters',
      part: TABLE_PART,
    });
    expect(
      (
        await capture(
          {},
          {
            limits: {
              maxFormulaCharacters: calculated.length,
              maxTotalFormulaCharacters: total - 1,
            },
          },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: total,
      limit: total - 1,
      limitName: 'maxTotalFormulaCharacters',
      part: TABLE_PART,
    });
  });

  it('enforces table-owned text accounting at the exact boundary', async () => {
    const xml = `<table xmlns="${XLSX_SPREADSHEET_NS}" id="2" name="Text_Table" displayName="Text_Table" comment="abc" ref="A1:A1"><tableColumns count="1"><tableColumn id="1" name="Value" uniqueName="u" totalsRowLabel="l" headerRowCellStyle="h" dataCellStyle="d" totalsRowCellStyle="t"/></tableColumns><tableStyleInfo name="s"/></table>`;
    const emptyWorksheet = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheetData/><tableParts count="1"><tablePart r:id="table"/></tableParts></worksheet>`;
    const publicText =
      'Black box'.length + 'Text_Table'.length * 2 + 'abcValueulhdts'.length;
    await expect(
      parseXlsx(
        await createTableWorkbook({
          'xl/worksheets/sheet1.xml': emptyWorksheet,
          [TABLE_PART]: xml,
        }),
        {
          limits: { maxTextCharacters: publicText },
        },
      ),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(
          {
            'xl/worksheets/sheet1.xml': emptyWorksheet,
            [TABLE_PART]: xml,
          },
          { limits: { maxTextCharacters: publicText - 1 } },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: publicText,
      limit: publicText - 1,
      limitName: 'maxTextCharacters',
      part: TABLE_PART,
    });
  });

  it('accepts the maximum unsigned table and column identities', async () => {
    const xml = `<table xmlns="${XLSX_SPREADSHEET_NS}" id="4294967295" name="Maximum_Table" displayName="Maximum_Table" connectionId="4294967295" ref="A1:A1"><tableColumns count="1"><tableColumn id="4294967295" name="Value" queryTableFieldId="4294967295"/></tableColumns></table>`;
    const document = await parseXlsx(
      await createTableWorkbook({ [TABLE_PART]: xml }),
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.tables[0] : undefined,
    ).toMatchObject({
      connectionId: 0xffff_ffff,
      id: 0xffff_ffff,
      columns: [{ id: 0xffff_ffff, queryTableFieldId: 0xffff_ffff }],
    });
  });

  it.each([
    ['average', 'average'],
    ['count', 'count'],
    ['countNums', 'count-numbers'],
    ['max', 'maximum'],
    ['min', 'minimum'],
    ['none', 'none'],
    ['stdDev', 'standard-deviation'],
    ['sum', 'sum'],
    ['var', 'variance'],
  ] as const)(
    'normalizes table totals function %s',
    async (source, expected) => {
      const xml = `<table xmlns="${XLSX_SPREADSHEET_NS}" id="2" name="Totals_${source}" displayName="Totals_${source}" ref="A1:A2" totalsRowCount="1"><tableColumns count="1"><tableColumn id="1" name="Value" totalsRowFunction="${source}"/></tableColumns></table>`;
      const document = await parseXlsx(
        await createTableWorkbook({ [TABLE_PART]: xml }),
      );
      const sheet = document.sheets[0]!;
      expect(
        sheet.kind === 'worksheet'
          ? sheet.tables[0]?.columns[0]?.totalsFunction
          : undefined,
      ).toBe(expected);
    },
  );

  it('accepts Unicode table identifiers and default flags', async () => {
    const xml = `<table xmlns="${XLSX_SPREADSHEET_NS}" id="2" name="Đơn_Hàng" displayName="Đơn_Hàng" ref="A1:A1" headerRowCount="0"><tableColumns count="1"><tableColumn id="1" name="Sản phẩm"/></tableColumns><tableStyleInfo/></table>`;
    const document = await parseXlsx(
      await createTableWorkbook({ [TABLE_PART]: xml }),
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.tables[0] : undefined,
    ).toMatchObject({
      displayName: 'Đơn_Hàng',
      headerRow: false,
      insertRow: false,
      published: false,
      style: {
        showColumnStripes: false,
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: false,
      },
      totalsRow: false,
      totalsRowShown: true,
    });
  });

  it('parses every lexical table boolean without truthiness coercion', async () => {
    const xml = `<table xmlns="${XLSX_SPREADSHEET_NS}" id="2" name="Boolean_Table" displayName="Boolean_Table" ref="A1:A2" headerRowCount="1" insertRow="false" insertRowShift="0" published="true" totalsRowCount="1" totalsRowShown="1"><tableColumns count="1"><tableColumn id="1" name="Value" totalsRowFunction="custom"><totalsRowFormula array="false">SUM([Value])</totalsRowFormula></tableColumn></tableColumns><tableStyleInfo showColumnStripes="false" showFirstColumn="0" showLastColumn="true" showRowStripes="1"/></table>`;
    const document = await parseXlsx(
      await createTableWorkbook({ [TABLE_PART]: xml }),
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.tables[0] : undefined,
    ).toMatchObject({
      insertRow: false,
      insertRowShift: false,
      published: true,
      style: {
        showColumnStripes: false,
        showFirstColumn: false,
        showLastColumn: true,
        showRowStripes: true,
      },
      totalsRowShown: true,
      columns: [{ totalsFormula: { array: false } }],
    });
  });

  it.each([
    ['queryTable', 'query-table'],
    ['worksheet', 'worksheet'],
    ['xml', 'xml'],
  ] as const)('normalizes table type %s', async (source, expected) => {
    const document = await parseXlsx(
      await createTableWorkbook({
        [TABLE_PART]: tableXml(`tableType="${source}"`),
      }),
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.tables[0]?.tableType : undefined,
    ).toBe(expected);
  });

  it('enforces configured table row and column bounds exactly', async () => {
    const emptyWorksheet = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheetData/><tableParts count="1"><tablePart r:id="table"/></tableParts></worksheet>`;
    await expect(
      parseXlsx(await createTableWorkbook(), {
        limits: { maxColumnsPerWorksheet: 3, maxRowsPerWorksheet: 4 },
      }),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(
          { 'xl/worksheets/sheet1.xml': emptyWorksheet },
          { limits: { maxColumnsPerWorksheet: 2 } },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 3,
      limit: 2,
      limitName: 'maxColumnsPerWorksheet',
      part: TABLE_PART,
    });
    expect(
      (
        await capture(
          { 'xl/worksheets/sheet1.xml': emptyWorksheet },
          { limits: { maxRowsPerWorksheet: 3 } },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 4,
      limit: 3,
      limitName: 'maxRowsPerWorksheet',
      part: TABLE_PART,
    });
  });

  it('validates all table differential-style references', async () => {
    const valid = tableXml(
      'dataDxfId="0" headerRowDxfId="0" totalsRowDxfId="0" headerRowBorderDxfId="0" tableBorderDxfId="0" totalsRowBorderDxfId="0"',
    ).replace(
      'id="1" name="Product"',
      'id="1" name="Product" dataDxfId="0" headerRowDxfId="0" totalsRowDxfId="0"',
    );
    await expect(
      parseXlsx(
        await createTableWorkbook({
          'xl/styles.xml': STYLES_WITH_DXF,
          [TABLE_PART]: valid,
        }),
      ),
    ).resolves.toBeDefined();
    for (const attribute of [
      'dataDxfId',
      'headerRowDxfId',
      'totalsRowDxfId',
      'headerRowBorderDxfId',
      'tableBorderDxfId',
      'totalsRowBorderDxfId',
    ]) {
      const error = await capture({
        'xl/styles.xml': STYLES_WITH_DXF,
        [TABLE_PART]: valid.replace(`${attribute}="0"`, `${attribute}="1"`),
      });
      expect(error.diagnostic.message).toContain(
        'differential-style reference is invalid',
      );
    }
    for (const [attribute, message] of [
      ['dataDxfId', 'Table data differential-style reference is invalid'],
      [
        'headerRowDxfId',
        'Table header differential-style reference is invalid',
      ],
      [
        'totalsRowDxfId',
        'Table totals differential-style reference is invalid',
      ],
    ] as const) {
      const columnXml = tableXml().replace(
        'id="1" name="Product"',
        `id="1" name="Product" ${attribute}="1"`,
      );
      const error = await capture({
        'xl/styles.xml': STYLES_WITH_DXF,
        [TABLE_PART]: columnXml,
      });
      expect(error.diagnostic.message).toBe(message);
    }
  });

  it('distinguishes every table-selection rectangle boundary', async () => {
    const bytes = await createTableWorkbook({
      [TABLE_PART]: tableXml()
        .replace('ref="A1:C4"', 'ref="B2:C3"')
        .replace('<tableColumns count="3">', '<tableColumns count="2">')
        .replace(
          '<tableColumn id="3" name="Tax" totalsRowFunction="custom">\n        <totalsRowFormula>SUM([Tax])</totalsRowFormula>\n      </tableColumn>',
          '',
        ),
    });
    for (const reference of ['A1', 'B1', 'D1', 'A2', 'D2', 'A4', 'B4', 'D4']) {
      const document = await parseXlsx(bytes, {
        selection: { ranges: { Sheet1: [reference] } },
      });
      const sheet = document.sheets[0]!;
      expect(sheet.kind === 'worksheet' ? sheet.tables : []).toEqual([]);
    }
    for (const reference of ['B2', 'C2', 'B3', 'C3']) {
      const document = await parseXlsx(bytes, {
        selection: { ranges: { Sheet1: [reference] } },
      });
      const sheet = document.sheets[0]!;
      expect(sheet.kind === 'worksheet' ? sheet.tables : []).toHaveLength(1);
    }
  });

  it.each([
    [
      {
        'xl/worksheets/_rels/sheet1.xml.rels': relationship(
          'https://example.invalid',
          'TargetMode="External"',
        ),
      },
      'invalid-relationship-target',
      'Worksheet table relationship must be internal',
    ],
    [
      {
        'xl/worksheets/_rels/sheet1.xml.rels': relationship(
          '../tables/table1.xml',
          '',
          `${XLSX_OFFICE_REL_TYPE}hyperlink`,
        ),
      },
      'invalid-document-structure',
      'Worksheet table relationship is missing',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          TABLE_CONTENT_TYPE,
          'application/xml',
        ),
      },
      'invalid-document-structure',
      'Worksheet table target has the wrong content type',
    ],
    [
      {
        'xl/worksheets/sheet1.xml': worksheet(
          '<tableParts count="1"><tablePart r:id="missing"/></tableParts>',
        ),
      },
      'invalid-document-structure',
      'Worksheet table relationship is missing',
    ],
  ] as const)(
    'enforces owner-scoped table relationship %#',
    async (overrides, code, message) => {
      const error = await capture(overrides);
      expect(error.diagnostic).toMatchObject({ code, message });
    },
  );

  it.each([
    [`<notTable xmlns="${XLSX_SPREADSHEET_NS}"/>`, 'Table root is missing'],
    [
      '<table xmlns="urn:wrong" id="1" name="Wrong" displayName="Wrong" ref="A1:A1"><tableColumns count="1"><tableColumn id="1" name="Value"/></tableColumns></table>',
      'Table root has the wrong namespace',
    ],
    [
      `<table xmlns="${XLSX_SPREADSHEET_NS}" id="1" name="Empty" displayName="Empty" ref="A1:A1"><tableColumns count="0"/></table>`,
      'Table columns collection is invalid',
    ],
    [
      tableXml().replace(
        '<tableColumns count="3">',
        '<tableColumns count="1"><tableColumn id="1" name="Duplicate"/></tableColumns><tableColumns count="3">',
      ),
      'Table columns collection is missing',
    ],
  ] as const)(
    'rejects invalid table XML structure %#',
    async (xml, message) => {
      const error = await capture({ [TABLE_PART]: xml });
      expect(error.diagnostic).toMatchObject({
        code: 'invalid-document-structure',
        message,
        part: TABLE_PART,
      });
    },
  );

  it.each([
    ['insertRow="1"', 'insertRow="bad"', 'Table insert-row flag is invalid'],
    ['published="1"', 'published="bad"', 'Table published flag is invalid'],
    [
      'totalsRowShown="1"',
      'totalsRowShown="bad"',
      'Table totals-row shown flag is invalid',
    ],
    ['array="1"', 'array="bad"', 'Table formula array flag is invalid'],
    [
      'showFirstColumn="1"',
      'showFirstColumn="bad"',
      'Table first-column flag is invalid',
    ],
    [
      'showLastColumn="0"',
      'showLastColumn="bad"',
      'Table last-column flag is invalid',
    ],
    [
      'showRowStripes="1"',
      'showRowStripes="bad"',
      'Table row-stripe flag is invalid',
    ],
    [
      'showColumnStripes="0"',
      'showColumnStripes="bad"',
      'Table column-stripe flag is invalid',
    ],
  ] as const)(
    'rejects invalid table boolean %#',
    async (source, replacement, message) => {
      const error = await capture({
        [TABLE_PART]: tableXml().replace(source, replacement),
      });
      expect(error.diagnostic.message).toBe(message);
    },
  );

  it.each([
    [tableXml().replace('name="Sales_Internal"', ''), 'Table name is invalid'],
    [tableXml().replace('displayName="Sales"', ''), 'Table name is invalid'],
    [
      tableXml().replace('count="3"', 'count="01"'),
      'Table column count is invalid',
    ],
    [
      tableXml().replace('id="1" name="Product"', 'id="01" name="Product"'),
      'Table column ID is invalid',
    ],
    [
      tableXml().replace('id="1" name="Product"', 'id="1" name=""'),
      'Table column name is invalid',
    ],
    [
      tableXml().replace('totalsRowFunction="sum"', 'totalsRowFunction="bad"'),
      'Table totals-row function is invalid',
    ],
    [
      tableXml().replace(
        '<calculatedColumnFormula array="1">SUBTOTAL(109,[Amount])</calculatedColumnFormula>',
        '<calculatedColumnFormula array="1"/>',
      ),
      'Table calculated-column formula is invalid',
    ],
    [
      tableXml().replace(
        '<calculatedColumnFormula array="1">SUBTOTAL(109,[Amount])</calculatedColumnFormula>',
        '<calculatedColumnFormula>A</calculatedColumnFormula><calculatedColumnFormula>B</calculatedColumnFormula>',
      ),
      'Table calculated-column formula is invalid',
    ],
    [tableXml('connectionId="01"'), 'Table connection ID is invalid'],
    [
      tableXml().replace(
        'totalsRowCount="1"',
        'headerRowCount="01" totalsRowCount="1"',
      ),
      'Table header-row count is invalid',
    ],
    [
      tableXml().replace('totalsRowCount="1"', 'totalsRowCount="01"'),
      'Table totals-row count is invalid',
    ],
    [
      tableXml().replace(
        'id="1" name="Product"',
        'id="1" name="Product" queryTableFieldId="01"',
      ),
      'Table column query-field ID is invalid',
    ],
    [
      tableXml().replace(
        '<totalsRowFormula>SUM([Tax])</totalsRowFormula>',
        '<totalsRowFormula/>',
      ),
      'Table totals-row formula is invalid',
    ],
  ] as const)('rejects exact table value contract %#', async (xml, message) => {
    const error = await capture({ [TABLE_PART]: xml });
    expect(error.diagnostic).toMatchObject({ message, part: TABLE_PART });
  });

  it('classifies duplicate table formulas as invalid structure', async () => {
    const xml = tableXml().replace(
      '<calculatedColumnFormula array="1">SUBTOTAL(109,[Amount])</calculatedColumnFormula>',
      '<calculatedColumnFormula>A</calculatedColumnFormula><calculatedColumnFormula>B</calculatedColumnFormula>',
    );
    const error = await capture({ [TABLE_PART]: xml });
    expect(error.diagnostic).toMatchObject({
      code: 'invalid-document-structure',
      message: 'Table calculated-column formula is invalid',
      part: TABLE_PART,
    });
  });

  it('classifies duplicate table style info as invalid structure', async () => {
    const xml = tableXml().replace(
      '<tableStyleInfo ',
      '<tableStyleInfo/><tableStyleInfo ',
    );
    const error = await capture({ [TABLE_PART]: xml });
    expect(error.diagnostic).toMatchObject({
      code: 'invalid-document-structure',
      message: 'Table style info is invalid',
      part: TABLE_PART,
    });
  });

  it.each([
    [tableXml().replace('id="1"', 'id="0"'), 'Table ID is invalid'],
    [tableXml().replace('id="1"', 'id="01"'), 'Table ID is invalid'],
    [tableXml().replace('id="1"', 'id="4294967296"'), 'Table ID is invalid'],
    [
      tableXml().replace('name="Sales_Internal"', 'name="A1"'),
      'Table name is invalid',
    ],
    [
      tableXml().replace('displayName="Sales"', 'displayName="Bad Name"'),
      'Table name is invalid',
    ],
    [
      tableXml().replace('ref="A1:C4"', 'ref="$A$1:$C$4"'),
      'Table range is invalid',
    ],
    [
      tableXml().replace('count="3"', 'count="2"'),
      'Table column count does not match',
    ],
    [
      tableXml().replace('ref="A1:C4"', 'ref="A1:B4"'),
      'Table column count does not match its range',
    ],
    [
      tableXml().replace('id="3" name="Tax"', 'id="2" name="Tax"'),
      'Table contains duplicate column IDs',
    ],
    [
      tableXml().replace('id="3" name="Tax"', 'id="3" name="amount"'),
      'Table contains duplicate column names',
    ],
    [
      tableXml()
        .replace('name="Product"', 'name="Straße"')
        .replace('name="Amount"', 'name="STRASSE"'),
      'Table contains duplicate column names',
    ],
    [
      tableXml().replace('totalsRowCount="1"', 'totalsRowCount="2"'),
      'Table header or totals row count is invalid',
    ],
    [
      tableXml().replace(
        'totalsRowCount="1"',
        'headerRowCount="2" totalsRowCount="1"',
      ),
      'Table header or totals row count is invalid',
    ],
    [
      tableXml().replace('ref="A1:C4"', 'ref="A1:C1"'),
      'Table range is too short for its header and totals rows',
    ],
    [
      tableXml().replace(
        /<tableColumns[\s\S]*<\/tableColumns>/u,
        '<tableColumns count="0"/>',
      ),
      'Table columns collection is invalid',
    ],
    [tableXml('tableType="bad"'), 'Table type is invalid'],
    [
      tableXml('insertRowShift="bad"'),
      'Table insert-row-shift flag is invalid',
    ],
    [
      tableXml().replace(
        'totalsRowFunction="custom"',
        'totalsRowFunction="sum"',
      ),
      'Table totals formula requires the custom function',
    ],
    [
      tableXml().replace('<totalsRowFormula>SUM([Tax])</totalsRowFormula>', ''),
      'Custom table totals formula is missing',
    ],
  ] as const)('rejects invalid table contract %#', async (xml, message) => {
    const error = await capture({ [TABLE_PART]: xml });
    expect(error.diagnostic.message).toBe(message);
  });

  it('returns no tables for a workbook without tableParts', async () => {
    const document = await parseXlsx(await createIndependentXlsx());
    const sheet = document.sheets[0]!;
    expect(sheet.kind === 'worksheet' ? sheet.tables : []).toEqual([]);
  });
});
