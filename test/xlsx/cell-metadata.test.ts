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
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const DYNAMIC_NS =
  'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
const RICH_NS =
  'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';
const METADATA_REL = `${XLSX_OFFICE_REL_TYPE}sheetMetadata`;
const STRICT_SPREADSHEET_NS = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const STRICT_REL_NS = 'http://purl.oclc.org/ooxml/officeDocument/relationships';

function contentTypes(metadataType: string): string {
  return `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/xl/metadata.xml" ContentType="${metadataType}"/>
  </Types>`;
}

function workbookRelationships(
  metadataRelationship = `<Relationship Id="metadata" Type="${METADATA_REL}" Target="metadata.xml"/>`,
): string {
  return `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
    <Relationship Id="rIdSheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rIdStyles" Type="${XLSX_OFFICE_REL_TYPE}styles" Target="styles.xml"/>
    <Relationship Id="rIdSharedStrings" Type="${XLSX_OFFICE_REL_TYPE}sharedStrings" Target="sharedStrings.xml"/>
    ${metadataRelationship}
  </Relationships>`;
}

function metadataXml(body: string): string {
  return `<metadata xmlns="${XLSX_SPREADSHEET_NS}" xmlns:xda="${DYNAMIC_NS}" xmlns:xlrd="${RICH_NS}">${body}</metadata>`;
}

const COMPLETE_METADATA = metadataXml(`
  <metadataTypes count="2">
    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="true"/>
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="0"/>
  </metadataTypes>
  <futureMetadata name="XLDAPR" count="2">
    <bk><extLst><ext uri="{dynamic}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>
    <bk><extLst><ext uri="{dynamic}"><xda:dynamicArrayProperties fDynamic="false" fCollapsed="true"/></ext></extLst></bk>
  </futureMetadata>
  <futureMetadata name="XLRICHVALUE" count="2">
    <bk><extLst><ext uri="{rich}"><xlrd:rvb i="4"/></ext></extLst></bk>
    <bk><extLst><ext uri="{rich}"><xlrd:rvb i="9"/></ext></extLst></bk>
  </futureMetadata>
  <cellMetadata count="2">
    <bk><rc t="1" v="0"/></bk>
    <bk><rc t="1" v="1"/></bk>
  </cellMetadata>
  <valueMetadata count="2">
    <bk><rc t="2" v="0"/></bk>
    <bk><rc t="2" v="1"/><rc t="1" v="1"/></bk>
  </valueMetadata>
`);

function worksheet(cells: string): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData>${cells}</sheetData></worksheet>`;
}

const COMPLETE_CELLS = `
  <row r="1"><c r="A1" cm="1" vm="1"><v>1</v></c></row>
  <row r="2"><c r="B2" cm="2" vm="2" t="b"><v>1</v></c></row>
`;

async function source(
  overrides: Record<string, string | Uint8Array | null> = {},
): Promise<Uint8Array> {
  return createIndependentXlsx({
    '[Content_Types].xml': contentTypes(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml',
    ),
    'xl/_rels/workbook.xml.rels': workbookRelationships(),
    'xl/metadata.xml': COMPLETE_METADATA,
    'xl/worksheets/sheet1.xml': worksheet(COMPLETE_CELLS),
    ...overrides,
  });
}

describe('XLSX modern cell metadata', () => {
  it('resolves dynamic-array and rich-value metadata onto selected cells', async () => {
    const document = await parseXlsx(await source(), { errorMode: 'strict' });
    const sheet = document.sheets[0];
    expect(sheet?.kind).toBe('worksheet');
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows[0]?.cells[0]?.metadata).toStrictEqual({
      cell: [{ collapsed: false, dynamic: true, kind: 'dynamic-array' }],
      value: [{ kind: 'rich-value', valueIndex: 4 }],
    });
    expect(sheet.rows[1]?.cells[0]?.metadata).toStrictEqual({
      cell: [{ collapsed: true, dynamic: false, kind: 'dynamic-array' }],
      value: [
        { kind: 'rich-value', valueIndex: 9 },
        { collapsed: true, dynamic: false, kind: 'dynamic-array' },
      ],
    });
    expect(sheet.rows[0]?.cells[0]?.metadata?.cell?.[0]).not.toBe(
      sheet.rows[1]?.cells[0]?.metadata?.value?.[1],
    );
    expect(JSON.parse(JSON.stringify(document))).toStrictEqual(document);
  });

  it('validates complete metadata while emitting only selected cells', async () => {
    const document = await parseXlsx(await source(), {
      errorMode: 'strict',
      selection: { ranges: { Sheet1: ['B2'] } },
    });
    const sheet = document.sheets[0];
    expect(sheet?.kind).toBe('worksheet');
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows).toHaveLength(1);
    expect(sheet.rows[0]?.cells[0]?.address).toBe('B2');
  });

  it('resolves value metadata without inventing cell metadata', async () => {
    const document = await parseXlsx(
      await source({
        'xl/worksheets/sheet1.xml': worksheet(
          '<row r="1"><c r="A1" vm="1"><v>1</v></c></row>',
        ),
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0];
    expect(sheet?.kind).toBe('worksheet');
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows[0]?.cells[0]?.metadata).toStrictEqual({
      value: [{ kind: 'rich-value', valueIndex: 4 }],
    });
  });

  it('preserves modern metadata through standalone JSON and exact R0', async () => {
    const input = await source();
    const snapshot = await readXlsxRoundTrip(input);
    expect(snapshot.document.sheets[0]?.kind).toBe('worksheet');
    expect(snapshot.supportProfile.domains).toContainEqual({
      domain: 'modern-cell-metadata',
      level: 'preservation-only',
    });
    const output = await writeXlsxRoundTrip(
      await validateXlsxRoundTripJson(
        JSON.parse(JSON.stringify(snapshot)) as unknown,
      ),
    );
    expect(output.data).toStrictEqual(input);
    expect(output.report.level).toBe('R0');
  });

  it('parses Strict workbook-owned cell metadata', async () => {
    const input = await source({
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${STRICT_REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
        <Relationship Id="sheet" Type="${STRICT_REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="styles" Type="${STRICT_REL_NS}/styles" Target="styles.xml"/>
        <Relationship Id="strings" Type="${STRICT_REL_NS}/sharedStrings" Target="sharedStrings.xml"/>
        <Relationship Id="metadata" Type="${STRICT_REL_NS}/sheetMetadata" Target="metadata.xml"/>
      </Relationships>`,
      'xl/workbook.xml': `<s:workbook xmlns:s="${STRICT_SPREADSHEET_NS}" xmlns:r="${STRICT_REL_NS}"><s:bookViews><s:workbookView/></s:bookViews><s:sheets><s:sheet name="Sheet1" sheetId="1" r:id="sheet"/></s:sheets><s:calcPr/></s:workbook>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${STRICT_SPREADSHEET_NS}"><s:sheetData><s:row r="1"><s:c r="A1" cm="1"><s:v>1</s:v></s:c></s:row></s:sheetData></s:worksheet>`,
      'xl/sharedStrings.xml': `<s:sst xmlns:s="${STRICT_SPREADSHEET_NS}" count="0" uniqueCount="0"/>`,
      'xl/styles.xml': `<s:styleSheet xmlns:s="${STRICT_SPREADSHEET_NS}"><s:fonts count="1"><s:font/></s:fonts><s:fills count="1"><s:fill><s:patternFill patternType="none"/></s:fill></s:fills><s:borders count="1"><s:border/></s:borders><s:cellStyleXfs count="1"><s:xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></s:cellStyleXfs><s:cellXfs count="1"><s:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></s:cellXfs><s:cellStyles count="1"><s:cellStyle name="Normal" xfId="0"/></s:cellStyles></s:styleSheet>`,
      'xl/metadata.xml': `<s:metadata xmlns:s="${STRICT_SPREADSHEET_NS}" xmlns:xda="${DYNAMIC_NS}">
        <s:metadataTypes count="1"><s:metadataType name="XLDAPR" minSupportedVersion="1"/></s:metadataTypes>
        <s:futureMetadata name="XLDAPR" count="1"><s:bk><s:extLst><s:ext uri="{dynamic}"><xda:dynamicArrayProperties fDynamic="1"/></s:ext></s:extLst></s:bk></s:futureMetadata>
        <s:cellMetadata count="1"><s:bk><s:rc t="1" v="0"/></s:bk></s:cellMetadata>
      </s:metadata>`,
    });
    const document = await parseXlsx(input, { errorMode: 'strict' });
    expect(document.sheets[0]?.kind).toBe('worksheet');
    if (document.sheets[0]?.kind !== 'worksheet') {
      throw new Error('Expected worksheet');
    }
    expect(document.sheets[0].rows[0]?.cells[0]?.metadata).toStrictEqual({
      cell: [{ collapsed: false, dynamic: true, kind: 'dynamic-array' }],
    });
  });

  it('diagnoses unsupported future metadata without exposing its payload', async () => {
    const unknown = metadataXml(`
      <metadataTypes count="1"><metadataType name="XLFUTURE" minSupportedVersion="1"/></metadataTypes>
      <futureMetadata name="XLFUTURE" count="1"><bk><extLst><ext uri="urn:secret"><x:secret xmlns:x="urn:x">hidden</x:secret></ext></extLst></bk></futureMetadata>
      <cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>
    `);
    const input = await source({
      'xl/metadata.xml': unknown,
      'xl/worksheets/sheet1.xml': worksheet(
        '<row r="1"><c r="A1" cm="1"><v>1</v></c></row>',
      ),
    });
    const result = await parseXlsxWithDiagnostics(input);
    expect(result.diagnostics).toStrictEqual([
      {
        code: 'unsupported-feature',
        message: 'Worksheet modern metadata content was omitted',
        part: 'xl/worksheets/sheet1.xml',
        severity: 'warning',
        sheet: 'Sheet1',
      },
    ]);
    expect(JSON.stringify(result)).not.toContain('urn:secret');
    expect(JSON.stringify(result)).not.toContain('hidden');
    await expect(
      parseXlsx(input, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: { code: 'unsupported-feature', severity: 'error' },
    });
    const snapshot = await readXlsxRoundTrip(input);
    expect((await writeXlsxRoundTrip(snapshot)).data).toStrictEqual(input);
  });

  it('preserves supported entries while diagnosing an unsupported record in the same block', async () => {
    const mixed = metadataXml(`
      <metadataTypes count="2"><metadataType name="XLDAPR" minSupportedVersion="1"/><metadataType name="XLFUTURE" minSupportedVersion="1"/></metadataTypes>
      <futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{dynamic}"><xda:dynamicArrayProperties fDynamic="1"/></ext></extLst></bk></futureMetadata>
      <futureMetadata name="XLFUTURE" count="1"><bk><extLst><ext uri="urn:future"/></extLst></bk></futureMetadata>
      <cellMetadata count="1"><bk><rc t="1" v="0"/><rc t="2" v="0"/></bk></cellMetadata>
    `);
    const result = await parseXlsxWithDiagnostics(
      await source({
        'xl/metadata.xml': mixed,
        'xl/worksheets/sheet1.xml': worksheet(
          '<row r="1"><c r="A1" cm="1"><v>1</v></c></row>',
        ),
      }),
    );
    expect(result.diagnostics).toHaveLength(1);
    const sheet = result.document.sheets[0];
    expect(sheet?.kind).toBe('worksheet');
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows[0]?.cells[0]?.metadata).toStrictEqual({
      cell: [{ collapsed: false, dynamic: true, kind: 'dynamic-array' }],
    });
  });

  it('accepts omitted optional collection counts', async () => {
    const input = await source({
      'xl/metadata.xml': COMPLETE_METADATA.replaceAll(/ count="\d+"/gu, ''),
    });
    await expect(
      parseXlsx(input, { errorMode: 'strict' }),
    ).resolves.toBeDefined();
  });

  it('requires metadata.xml whenever a cell authors cm or vm', async () => {
    const input = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': worksheet(
        '<row r="1"><c r="A1" cm="1"><v>1</v></c></row>',
      ),
    });
    await expect(parseXlsx(input)).rejects.toMatchObject({
      diagnostic: {
        cell: 'A1',
        code: 'missing-required-part',
        message: 'Worksheet cell metadata part is missing',
      },
    });
  });

  it('recovers a malformed optional metadata part only when no cell depends on it', async () => {
    const input = await source({
      'xl/metadata.xml': '<metadata xmlns="urn:wrong"/>',
      'xl/worksheets/sheet1.xml': worksheet(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
      ),
    });
    const tolerant = await parseXlsxWithDiagnostics(input);
    expect(tolerant.document.sheets[0]?.kind).toBe('worksheet');
    if (tolerant.document.sheets[0]?.kind !== 'worksheet') {
      throw new Error('Expected worksheet');
    }
    expect(tolerant.document.sheets[0].rows[0]?.cells[0]).not.toHaveProperty(
      'metadata',
    );
    expect(tolerant.diagnostics).toStrictEqual([
      {
        code: 'invalid-document-structure',
        message: 'Cell metadata root has the wrong namespace',
        part: 'xl/metadata.xml',
        severity: 'warning',
      },
    ]);
    await expect(
      parseXlsx(input, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: { message: 'Cell metadata root has the wrong namespace' },
    });
  });

  it.each([
    ['cm="0"', 'Worksheet cell metadata reference is invalid'],
    ['cm="3"', 'Worksheet cell metadata reference is invalid'],
    ['vm="0"', 'Worksheet cell metadata reference is invalid'],
    ['vm="3"', 'Worksheet cell metadata reference is invalid'],
    ['cm="01"', 'Worksheet cell metadata index is invalid'],
    ['vm="x"', 'Worksheet value metadata index is invalid'],
  ] as const)(
    'rejects invalid cell metadata reference %s',
    async (attribute, message) => {
      await expect(
        parseXlsx(
          await source({
            'xl/worksheets/sheet1.xml': worksheet(
              `<row r="1"><c r="A1" ${attribute}><v>1</v></c></row>`,
            ),
          }),
          { errorMode: 'strict' },
        ),
      ).rejects.toMatchObject({ diagnostic: { cell: 'A1', message } });
    },
  );

  it.each([
    [
      { '[Content_Types].xml': contentTypes('application/xml') },
      'Cell metadata target has the wrong content type',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': workbookRelationships(
          `<Relationship Id="metadata" Type="${METADATA_REL}" Target="https://example.invalid/metadata.xml" TargetMode="External"/>`,
        ),
      },
      'Workbook cell-metadata relationship must be internal',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': workbookRelationships(
          `<Relationship Id="metadata" Type="${METADATA_REL}" Target="metadata.xml"/><Relationship Id="metadata2" Type="${METADATA_REL}" Target="metadata2.xml"/>`,
        ),
      },
      'Workbook cell-metadata relationship is duplicated',
    ],
    [
      { 'xl/metadata.xml': '<metadata xmlns="urn:wrong"/>' },
      'Cell metadata root has the wrong namespace',
    ],
    [
      { 'xl/metadata.xml': '<wrong xmlns="urn:wrong"/>' },
      'Cell metadata root is missing or duplicated',
    ],
    [
      { 'xl/metadata.xml': metadataXml('<metadataTypes count="0"/>') },
      'Cell metadata type collection is empty',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          'metadataTypes count="2"',
          'metadataTypes count="3"',
        ),
      },
      'Cell metadata type count does not match',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          'name="XLRICHVALUE"',
          'name="XLDAPR"',
        ),
      },
      'Cell metadata type name is invalid or duplicated',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          'minSupportedVersion="120000"',
          'minSupportedVersion="bad"',
        ),
      },
      'Cell metadata minimum version is invalid',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace('copy="1"', 'copy="bad"'),
      },
      'Cell metadata type flag is invalid',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<futureMetadata name="XLDAPR" count="2">',
          '<futureMetadata count="2">',
        ),
      },
      'Future metadata name is invalid or duplicated',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<futureMetadata name="XLRICHVALUE" count="2">',
          '<futureMetadata name="XLDAPR" count="2">',
        ),
      },
      'Future metadata name is invalid or duplicated',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<futureMetadata name="XLDAPR" count="2">',
          '<futureMetadata name="XLDAPR" count="3">',
        ),
      },
      'Future metadata block count does not match',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<futureMetadata name="XLDAPR" count="2">',
          '<futureMetadata name="XLDAPR" count="bad">',
        ),
      },
      'Future metadata block count does not match',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>',
          '<x:other xmlns:x="urn:x"/>',
        ),
      },
      'Dynamic-array metadata block is invalid',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>',
          '<xda:dynamicArrayProperties/><xda:dynamicArrayProperties/>',
        ),
      },
      'Dynamic-array metadata block is invalid',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(DYNAMIC_NS, 'urn:wrong'),
      },
      'Dynamic-array metadata has the wrong namespace',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>',
          '<xda:dynamicArrayProperties fDynamic="bad" fCollapsed="0"/>',
        ),
      },
      'Dynamic-array dynamic flag is invalid',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          'fDynamic="1" fCollapsed="0"',
          'fDynamic="1" fCollapsed="bad"',
        ),
      },
      'Dynamic-array collapsed flag is invalid',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<xlrd:rvb i="4"/>',
          '<x:other xmlns:x="urn:x"/>',
        ),
      },
      'Rich-value metadata block is invalid',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<xlrd:rvb i="4"/>',
          '<xlrd:rvb i="4"/><xlrd:rvb i="5"/>',
        ),
      },
      'Rich-value metadata block is invalid',
    ],
    [
      { 'xl/metadata.xml': COMPLETE_METADATA.replace(RICH_NS, 'urn:wrong') },
      'Rich-value metadata has the wrong namespace',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<xlrd:rvb i="4"/>',
          '<xlrd:rvb i="bad"/>',
        ),
      },
      'Rich-value metadata index is invalid',
    ],
    [
      {
        'xl/metadata.xml': metadataXml(
          '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="1"/></metadataTypes><cellMetadata count="0"/>',
        ),
      },
      'Cell metadata cellMetadata is empty',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<cellMetadata count="2">',
          '<cellMetadata count="3">',
        ),
      },
      'Cell metadata cellMetadata count does not match',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<bk><rc t="1" v="0"/></bk>',
          '<bk/>',
        ),
      },
      'Cell metadata block is empty',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<rc t="1" v="0"/>',
          '<rc t="bad" v="0"/>',
        ),
      },
      'Cell metadata type index is invalid',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<rc t="1" v="0"/>',
          '<rc t="1" v="bad"/>',
        ),
      },
      'Cell metadata value index is invalid',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<rc t="1" v="0"/>',
          '<rc t="0" v="0"/>',
        ),
      },
      'Cell metadata type reference is invalid',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<rc t="1" v="0"/>',
          '<rc t="3" v="0"/>',
        ),
      },
      'Cell metadata type reference is invalid',
    ],
    [
      {
        'xl/metadata.xml': COMPLETE_METADATA.replace(
          '<rc t="1" v="0"/>',
          '<rc t="1" v="2"/>',
        ),
      },
      'Cell metadata value reference is invalid',
    ],
  ] as const)(
    'rejects malformed modern metadata %#',
    async (overrides, message) => {
      await expect(
        parseXlsx(await source(overrides), { errorMode: 'strict' }),
      ).rejects.toMatchObject({ diagnostic: { message } });
    },
  );

  it('enforces metadata records at exact parsed-and-resolved boundaries', async () => {
    const input = await source();
    await expect(
      parseXlsx(input, {
        errorMode: 'strict',
        limits: { maxMetadataRecords: 10 },
      }),
    ).resolves.toBeDefined();
    await expect(
      parseXlsx(input, {
        errorMode: 'strict',
        limits: { maxMetadataRecords: 9 },
      }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'resource-limit-exceeded',
        limitName: 'maxMetadataRecords',
      },
    });
    await expect(
      parseXlsx(input, {
        errorMode: 'strict',
        limits: { maxMetadataRecords: 4 },
      }),
    ).rejects.toMatchObject({
      diagnostic: {
        actual: 5,
        code: 'resource-limit-exceeded',
        limitName: 'maxMetadataRecords',
      },
    });
  });
});
