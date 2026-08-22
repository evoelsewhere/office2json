import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  parseXlsxWithDiagnostics,
  readXlsxRoundTrip,
  writeXlsxRoundTrip,
} from '../../src/formats/xlsx';
import {
  createIndependentXlsx,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const METADATA_REL = `${XLSX_OFFICE_REL_TYPE}sheetMetadata`;
const RICH_DATA_NS =
  'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';
const RICH_REL =
  'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue';
const RICH_STRUCTURE_REL = `${RICH_REL}Structure`;

const CONTENT_TYPES = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
  <Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
  <Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
</Types>`;

const RELATIONSHIPS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
  <Relationship Id="rIdSheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="styles" Type="${XLSX_OFFICE_REL_TYPE}styles" Target="styles.xml"/>
  <Relationship Id="strings" Type="${XLSX_OFFICE_REL_TYPE}sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="metadata" Type="${METADATA_REL}" Target="metadata.xml"/>
  <Relationship Id="richStructure" Type="${RICH_STRUCTURE_REL}" Target="richData/rdrichvaluestructure.xml"/>
  <Relationship Id="richData" Type="${RICH_REL}" Target="richData/rdrichvalue.xml"/>
</Relationships>`;

const METADATA = `<metadata xmlns="${XLSX_SPREADSHEET_NS}" xmlns:xlrd="${RICH_DATA_NS}">
  <metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="2">
    <bk><extLst><ext uri="{rich}"><xlrd:rvb i="0"/></ext></extLst></bk>
    <bk><extLst><ext uri="{rich}"><xlrd:rvb i="1"/></ext></extLst></bk>
  </futureMetadata>
  <valueMetadata count="2">
    <bk><rc t="1" v="0"/></bk>
    <bk><rc t="1" v="1"/></bk>
  </valueMetadata>
</metadata>`;

const STRUCTURES = `<rvStructures xmlns="${RICH_DATA_NS}" count="2">
  <s t="_linkedentity">
    <k n="_DisplayString" t="s"/>
    <k n="Population" t="d"/>
    <k n="Count" t="i"/>
    <k n="Related" t="r"/>
    <k n="Array" t="a"/>
    <k n="Available" t="b"/>
    <k n="Failure" t="e"/>
    <k n="%EntityId" t="spb"/>
    <k n="sourceUrl" t="s"/>
  </s>
  <s t="_formattednumber"><k n="Text"/><k n="Value" t="d"/></s>
</rvStructures>`;

const DATA = `<rvData xmlns="${RICH_DATA_NS}" count="2">
  <rv s="0">
    <fb t="s">Seattle fallback</fb>
    <v>Seattle &amp; region</v><v>737015</v><v>4</v><v>1</v><v>2</v>
    <v>true</v><v>#N/A</v><v>9</v><v>https://secret.example/token</v>
  </rv>
  <rv s="1"><fb>12.5 display</fb><v>12.5</v><v>12.5</v></rv>
</rvData>`;

interface SourceOverrides {
  contentTypes?: string;
  data?: string;
  metadata?: string;
  relationships?: string;
  structures?: string;
  worksheet?: string;
}

async function source(overrides: SourceOverrides = {}): Promise<Uint8Array> {
  return createIndependentXlsx({
    '[Content_Types].xml': overrides.contentTypes ?? CONTENT_TYPES,
    'xl/_rels/workbook.xml.rels': overrides.relationships ?? RELATIONSHIPS,
    'xl/metadata.xml': overrides.metadata ?? METADATA,
    'xl/richData/rdrichvalue.xml': overrides.data ?? DATA,
    'xl/richData/rdrichvaluestructure.xml': overrides.structures ?? STRUCTURES,
    'xl/worksheets/sheet1.xml':
      overrides.worksheet ??
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData>
        <row r="1"><c r="A1" vm="1"><v>1</v></c></row>
        <row r="2"><c r="A2" vm="2"><v>2</v></c></row>
      </sheetData></worksheet>`,
  });
}

async function firstCell(overrides: SourceOverrides = {}) {
  const document = await parseXlsx(await source(overrides), {
    errorMode: 'strict',
  });
  const sheet = document.sheets[0];
  if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
  return sheet.rows[0]!.cells[0]!;
}

describe('XLSX rich values', () => {
  it('resolves the typed rich-value graph without exposing sensitive source data', async () => {
    const input = await source();
    const document = await parseXlsx(input, { errorMode: 'strict' });
    const sheet = document.sheets[0];
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');

    expect(sheet.rows[0]?.cells[0]?.metadata?.value).toStrictEqual([
      {
        data: {
          fallback: { kind: 'text', value: 'Seattle fallback' },
          fields: [
            {
              name: '_DisplayString',
              type: 's',
              value: { kind: 'text', value: 'Seattle & region' },
            },
            {
              name: 'Population',
              type: 'd',
              value: { kind: 'number', value: 737015 },
            },
            {
              name: 'Count',
              type: 'i',
              value: { kind: 'integer', value: 4 },
            },
            {
              name: 'Related',
              type: 'r',
              value: { kind: 'rich-value-index', value: 1 },
            },
            {
              name: 'Array',
              type: 'a',
              value: { kind: 'array-index', value: 2 },
            },
            {
              name: 'Available',
              type: 'b',
              value: { kind: 'boolean', value: true },
            },
            {
              name: 'Failure',
              type: 'e',
              value: { code: '#N/A', kind: 'error' },
            },
            {
              name: '%EntityId',
              type: 'spb',
              value: { kind: 'omitted' },
            },
            {
              name: 'sourceUrl',
              type: 's',
              value: { kind: 'omitted' },
            },
          ],
          sourceDataOmitted: true,
          type: '_linkedentity',
        },
        kind: 'rich-value',
        valueIndex: 0,
      },
    ]);
    expect(sheet.rows[1]?.cells[0]?.metadata?.value?.[0]).toMatchObject({
      data: {
        fallback: { kind: 'text', value: '12.5 display' },
        sourceDataOmitted: false,
        type: '_formattednumber',
      },
      valueIndex: 1,
    });
    expect(JSON.stringify(document)).not.toContain('secret.example');
    expect(JSON.parse(JSON.stringify(document))).toStrictEqual(document);
  });

  it('copies rich-value output per selected cell', async () => {
    const cell = await firstCell({
      worksheet: `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1">
        <c r="A1" vm="1"><v>1</v></c><c r="B1" vm="1"><v>2</v></c>
      </row></sheetData></worksheet>`,
    });
    const document = await parseXlsx(
      await source({
        worksheet: `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1">
          <c r="A1" vm="1"><v>1</v></c><c r="B1" vm="1"><v>2</v></c>
        </row></sheetData></worksheet>`,
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0];
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
    const second = sheet.rows[0]!.cells[1]!;
    const firstEntry = cell.metadata?.value?.[0];
    const secondEntry = second.metadata?.value?.[0];
    if (
      firstEntry?.kind !== 'rich-value' ||
      secondEntry?.kind !== 'rich-value'
    ) {
      throw new Error('Expected rich-value metadata');
    }
    expect(firstEntry).not.toBe(secondEntry);
    expect(firstEntry.data).not.toBe(secondEntry.data);
    expect(firstEntry.data?.fields[0]).not.toBe(secondEntry.data?.fields[0]);
  });

  it('preserves rich-value parts byte-for-byte through portable round-trip JSON', async () => {
    const input = await source();
    const snapshot = await readXlsxRoundTrip(input);
    const output = await writeXlsxRoundTrip(
      JSON.parse(JSON.stringify(snapshot)) as typeof snapshot,
    );
    expect(output.data).toStrictEqual(input);
    expect(output.report.level).toBe('R0');
  });

  it('honors selection while validating the complete rich-value graph', async () => {
    const document = await parseXlsx(await source(), {
      errorMode: 'strict',
      selection: { ranges: { Sheet1: ['A2'] } },
    });
    const sheet = document.sheets[0];
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows).toHaveLength(1);
    expect(sheet.rows[0]?.cells[0]?.metadata?.value?.[0]).toMatchObject({
      kind: 'rich-value',
      valueIndex: 1,
    });
  });

  it('does not claim rich data when both optional workbook relationships are absent', async () => {
    const relationships = RELATIONSHIPS.replace(
      /\s*<Relationship Id="rich(?:Structure|Data)"[^>]+\/>/gu,
      '',
    );
    const cell = await firstCell({ relationships });
    expect(cell.metadata?.value).toStrictEqual([
      { kind: 'rich-value', valueIndex: 0 },
    ]);
  });

  it.each([
    [
      {
        relationships: RELATIONSHIPS.replace(
          /\s*<Relationship Id="richData"[^>]+\/>/u,
          '',
        ),
      },
      'Rich-value data and structure parts must both exist',
    ],
    [
      {
        relationships: RELATIONSHIPS.replace(
          '</Relationships>',
          `<Relationship Id="richData2" Type="${RICH_REL}" Target="richData/rdrichvalue.xml"/></Relationships>`,
        ),
      },
      'Rich-value relationship is duplicated',
    ],
    [
      {
        relationships: RELATIONSHIPS.replace(
          'Target="richData/rdrichvalue.xml"',
          'Target="https://example.test/data" TargetMode="External"',
        ),
      },
      'Rich-value relationship must be internal',
    ],
    [
      {
        contentTypes: CONTENT_TYPES.replace(
          'application/vnd.ms-excel.rdrichvalue+xml',
          'application/xml',
        ),
      },
      'Rich-value target has the wrong content type',
    ],
    [
      {
        structures: STRUCTURES.replace(
          `xmlns="${RICH_DATA_NS}"`,
          'xmlns="urn:wrong"',
        ),
      },
      'Rich-value rvStructures root has the wrong namespace',
    ],
    [
      { structures: STRUCTURES.replace('count="2"', 'count="3"') },
      'Rich-value structure count does not match',
    ],
    [
      { structures: STRUCTURES.replace('t="_linkedentity"', '') },
      'Rich-value structure type is invalid',
    ],
    [
      { structures: STRUCTURES.replace('n="Population"', '') },
      'Rich-value key name is invalid',
    ],
    [
      {
        structures: STRUCTURES.replace(
          'n="Population" t="d"',
          'n="Population" t="x"',
        ),
      },
      'Rich-value key type is invalid',
    ],
    [
      { data: DATA.replace('<rv s="0">', '<rv s="9">') },
      'Rich-value structure reference is invalid',
    ],
    [
      { data: DATA.replace('<v>https://secret.example/token</v>', '') },
      'Rich-value field count does not match',
    ],
    [
      { data: DATA.replace('<v>true</v>', '<v>yes</v>') },
      'Rich-value boolean is invalid',
    ],
    [
      { data: DATA.replace('<v>737015</v>', '<v>NaN</v>') },
      'Rich-value number is invalid',
    ],
    [
      {
        data: DATA.replace(
          '<v>4</v><v>1</v><v>2</v>',
          '<v>4</v><v>8</v><v>2</v>',
        ),
      },
      'Rich-value reference is invalid',
    ],
    [
      { data: DATA.replace('<v>#N/A</v>', '<v>bad</v>') },
      'Rich-value error is invalid',
    ],
    [
      { metadata: METADATA.replace('xlrd:rvb i="1"', 'xlrd:rvb i="9"') },
      'Rich-value metadata index is invalid',
    ],
    [
      {
        data: DATA.replace(
          '<fb t="s">Seattle fallback</fb>',
          '<fb t="x">Seattle fallback</fb>',
        ),
      },
      'Rich-value fallback type is invalid',
    ],
  ] as const)(
    'rejects malformed rich-value package %#',
    async (overrides, message) => {
      await expect(
        parseXlsx(await source(overrides), { errorMode: 'strict' }),
      ).rejects.toThrow(message);
    },
  );

  it('diagnoses malformed optional rich data in tolerant mode without leaking it', async () => {
    const input = await source({
      data: DATA.replace('<v>true</v>', '<v>secret-invalid</v>'),
    });
    const result = await parseXlsxWithDiagnostics(input);
    expect(result.diagnostics).toContainEqual({
      code: 'invalid-document-value',
      message: 'Rich-value boolean is invalid',
      part: 'xl/richData/rdrichvalue.xml',
      severity: 'warning',
    });
    expect(JSON.stringify(result)).not.toContain('secret-invalid');
    const sheet = result.document?.sheets[0];
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows[0]?.cells[0]?.metadata?.value).toStrictEqual([
      { kind: 'rich-value', valueIndex: 0 },
    ]);
  });

  it('enforces aggregate metadata and output text limits', async () => {
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        limits: { maxMetadataRecords: 21 },
      }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'resource-limit-exceeded',
        limitName: 'maxMetadataRecords',
        part: 'xl/richData/rdrichvalue.xml',
      },
    });
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        limits: { maxMetadataRecords: 37 },
      }),
    ).resolves.toBeDefined();
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        limits: { maxMetadataRecords: 36 },
      }),
    ).rejects.toMatchObject({
      diagnostic: { code: 'resource-limit-exceeded' },
    });
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        limits: { maxTextCharacters: 174 },
      }),
    ).resolves.toBeDefined();
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        limits: { maxTextCharacters: 173 },
      }),
    ).rejects.toMatchObject({
      diagnostic: { code: 'resource-limit-exceeded' },
    });
  });
});
