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
  type XlsxBlackBoxPart,
} from '../black-box/xlsx-package';

const RICH_NS =
  'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';
const RICH_TYPES_NS =
  'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2';
const RICH_REL_NS =
  'http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel';
const RICH_REL =
  'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue';
const OFFICE_REL_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const CONTENT_TYPES = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
  <Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
  <Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
  <Override PartName="/xl/richData/rdRichValueTypes.xml" ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>
  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
</Types>`;

const WORKBOOK_RELATIONSHIPS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
  <Relationship Id="rIdSheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="styles" Type="${XLSX_OFFICE_REL_TYPE}styles" Target="styles.xml"/>
  <Relationship Id="strings" Type="${XLSX_OFFICE_REL_TYPE}sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="metadata" Type="${XLSX_OFFICE_REL_TYPE}sheetMetadata" Target="metadata.xml"/>
  <Relationship Id="rich" Type="${RICH_REL}" Target="richData/rdrichvalue.xml"/>
  <Relationship Id="structure" Type="${RICH_REL}Structure" Target="richData/rdrichvaluestructure.xml"/>
  <Relationship Id="types" Type="${RICH_REL}Types" Target="richData/rdRichValueTypes.xml"/>
  <Relationship Id="imageRels" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>
</Relationships>`;

const METADATA = `<metadata xmlns="${XLSX_SPREADSHEET_NS}" xmlns:xlrd="${RICH_NS}">
  <metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{rich}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata>
  <valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>
</metadata>`;

const STRUCTURES = `<rvStructures xmlns="${RICH_NS}" count="1"><s t="_localImage">
  <k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/>
</s></rvStructures>`;

const DATA = `<rvData xmlns="${RICH_NS}" count="1"><rv s="0"><v>0</v><v>5</v></rv></rvData>`;
const TYPES = `<rvTypesInfo xmlns="${RICH_TYPES_NS}"/>`;
const IMAGE_REFS = `<richValueRels xmlns="${RICH_REL_NS}" xmlns:r="${OFFICE_REL_NS}"><rel r:id="rId1"/></richValueRels>`;
const IMAGE_RELATIONSHIPS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="rId1" Type="${OFFICE_REL_NS}/image" Target="../media/image1.png"/></Relationships>`;

async function source(
  overrides: Record<string, XlsxBlackBoxPart> = {},
): Promise<Uint8Array> {
  return createIndependentXlsx({
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/_rels/workbook.xml.rels': WORKBOOK_RELATIONSHIPS,
    'xl/metadata.xml': METADATA,
    'xl/richData/_rels/richValueRel.xml.rels': IMAGE_RELATIONSHIPS,
    'xl/richData/rdRichValueTypes.xml': TYPES,
    'xl/richData/rdrichvalue.xml': DATA,
    'xl/richData/rdrichvaluestructure.xml': STRUCTURES,
    'xl/richData/richValueRel.xml': IMAGE_REFS,
    'xl/media/image1.png': new Uint8Array([1, 2, 3]),
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData>
      <row r="1"><c r="A1" vm="1"><v>1</v></c><c r="B1" vm="1"><v>2</v></c></row>
    </sheetData></worksheet>`,
    ...overrides,
  });
}

function expectedImage(payload = false) {
  return {
    ...(payload ? { base64: 'data:image/png;base64,AQID', byteLength: 3 } : {}),
    contentType: 'image/png',
    kind: 'local-image',
    part: 'xl/media/image1.png',
  };
}

describe('XLSX rich-value local images', () => {
  it('returns safe image metadata without reading bytes by default', async () => {
    const document = await parseXlsx(await source(), { errorMode: 'strict' });
    const sheet = document.sheets[0];
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
    const entry = sheet.rows[0]?.cells[0]?.metadata?.value?.[0];
    expect(entry).toMatchObject({
      data: {
        fields: [
          {
            name: '_rvRel:LocalImageIdentifier',
            type: 'i',
            value: { kind: 'omitted' },
          },
          {
            name: 'CalcOrigin',
            type: 'i',
            value: { kind: 'integer', value: 5 },
          },
        ],
        image: expectedImage(),
        sourceDataOmitted: true,
        type: '_localImage',
      },
      kind: 'rich-value',
      valueIndex: 0,
    });
    expect(JSON.parse(JSON.stringify(document))).toStrictEqual(document);
  });

  it('hydrates selected local images through canonical Base64 and repeated-byte limits', async () => {
    const document = await parseXlsx(await source(), {
      errorMode: 'strict',
      imageMode: 'base64',
      limits: { maxMediaBytes: 6 },
    });
    const sheet = document.sheets[0];
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
    for (const cell of sheet.rows[0]!.cells) {
      expect(cell.metadata?.value?.[0]).toMatchObject({
        data: { image: expectedImage(true) },
      });
    }
    expect(sheet.rows[0]?.cells[0]?.metadata?.value?.[0]).not.toBe(
      sheet.rows[0]?.cells[1]?.metadata?.value?.[0],
    );
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        imageMode: 'base64',
        limits: { maxMediaBytes: 5 },
      }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'resource-limit-exceeded',
        limitName: 'maxMediaBytes',
      },
    });
  });

  it('hydrates only selected image cells while validating the complete graph', async () => {
    const document = await parseXlsx(await source(), {
      errorMode: 'strict',
      imageMode: 'base64',
      limits: { maxMediaBytes: 3 },
      selection: { ranges: { Sheet1: ['B1'] } },
    });
    const sheet = document.sheets[0];
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows[0]?.cells).toHaveLength(1);
    expect(sheet.rows[0]?.cells[0]?.metadata?.value?.[0]).toMatchObject({
      data: { image: expectedImage(true) },
    });
  });

  it('enforces companion records and returned image metadata text exactly', async () => {
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        limits: { maxMetadataRecords: 12, maxTextCharacters: 161 },
      }),
    ).resolves.toBeDefined();
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        limits: { maxMetadataRecords: 11 },
      }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'resource-limit-exceeded',
        limitName: 'maxMetadataRecords',
      },
    });
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        limits: { maxTextCharacters: 160 },
      }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'resource-limit-exceeded',
        limitName: 'maxTextCharacters',
      },
    });
  });

  it('preserves local-image graph and media bytes through exact R0', async () => {
    const input = await source();
    const snapshot = await readXlsxRoundTrip(input);
    const output = await writeXlsxRoundTrip(
      JSON.parse(JSON.stringify(snapshot)) as typeof snapshot,
    );
    expect(output.data).toStrictEqual(input);
    expect(output.report.level).toBe('R0');
  });

  it('accepts Strict relationship namespaces in the Microsoft companion graph', async () => {
    const strictRelationships = IMAGE_RELATIONSHIPS.replace(
      OFFICE_REL_NS,
      'http://purl.oclc.org/ooxml/officeDocument/relationships',
    );
    const strictReferences = IMAGE_REFS.replace(
      OFFICE_REL_NS,
      'http://purl.oclc.org/ooxml/officeDocument/relationships',
    );
    const document = await parseXlsx(
      await source({
        'xl/richData/_rels/richValueRel.xml.rels': strictRelationships,
        'xl/richData/richValueRel.xml': strictReferences,
      }),
      { errorMode: 'strict' },
    );
    expect(document.styles).toBeDefined();
  });

  it('preserves an authored local-image fallback', async () => {
    const data = DATA.replace('<rv s="0">', '<rv s="0"><fb>preview</fb>');
    const document = await parseXlsx(
      await source({ 'xl/richData/rdrichvalue.xml': data }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0];
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows[0]?.cells[0]?.metadata?.value?.[0]).toMatchObject({
      data: { fallback: { kind: 'text', value: 'preview' } },
    });
  });

  it.each([
    [
      {
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELATIONSHIPS.replace(
          /\s*<Relationship Id="types"[^>]+\/>/u,
          '',
        ),
      },
      'Rich-value image types and relationships must both exist',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELATIONSHIPS.replace(
          /\s*<Relationship Id="(?:types|imageRels)"[^>]+\/>/gu,
          '',
        ),
      },
      'Rich-value local-image relationship graph is missing',
    ],
    [
      {
        'xl/richData/rdRichValueTypes.xml': TYPES.replace(
          RICH_TYPES_NS,
          'urn:wrong',
        ),
      },
      'Rich-value rvTypesInfo root has the wrong namespace',
    ],
    [
      {
        'xl/richData/rdRichValueTypes.xml': `<wrong xmlns="${RICH_TYPES_NS}"/>`,
      },
      'Rich-value rvTypesInfo root is missing or duplicated',
    ],
    [
      {
        'xl/richData/richValueRel.xml': IMAGE_REFS.replace(
          RICH_REL_NS,
          'urn:wrong',
        ),
      },
      'Rich-value richValueRels root has the wrong namespace',
    ],
    [
      {
        'xl/richData/richValueRel.xml': `<wrong xmlns="${RICH_REL_NS}" xmlns:r="${OFFICE_REL_NS}"/>`,
      },
      'Rich-value richValueRels root is missing or duplicated',
    ],
    [
      {
        'xl/richData/richValueRel.xml': IMAGE_REFS.replace(
          OFFICE_REL_NS,
          'urn:wrong',
        ),
      },
      'Rich-value image relationship namespace is invalid',
    ],
    [
      {
        'xl/richData/richValueRel.xml': IMAGE_REFS.replace(
          '<rel r:id="rId1"/>',
          '<rel>text</rel>',
        ),
      },
      'Rich-value companion collection is invalid',
    ],
    [
      {
        'xl/richData/richValueRel.xml': IMAGE_REFS.replace(
          '<rel r:id="rId1"/>',
          '<x:rel xmlns:x="urn:wrong" r:id="rId1"/>',
        ),
      },
      'Rich-value companion element has the wrong namespace',
    ],
    [
      {
        'xl/richData/richValueRel.xml': IMAGE_REFS.replace(
          '<rel r:id="rId1"/>',
          '<rel/>',
        ),
      },
      'Rich-value image relationship is invalid',
    ],
    [
      { 'xl/richData/_rels/richValueRel.xml.rels': null },
      'Required XLSX part is missing: xl/richData/_rels/richValueRel.xml.rels',
    ],
    [
      {
        'xl/richData/_rels/richValueRel.xml.rels': IMAGE_RELATIONSHIPS.replace(
          `${OFFICE_REL_NS}/image`,
          `${OFFICE_REL_NS}/hyperlink`,
        ),
      },
      'Rich-value image relationship is invalid',
    ],
    [
      {
        'xl/richData/_rels/richValueRel.xml.rels': IMAGE_RELATIONSHIPS.replace(
          '/>',
          ' TargetMode="External"/>',
        ).replace('../media/image1.png', 'https://example.test/image.png'),
      },
      'External rich-value images are not loaded',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'ContentType="image/png"',
          'ContentType="image/svg+xml"',
        ),
      },
      'Rich-value image content type is not safely supported',
    ],
    [{ 'xl/media/image1.png': null }, 'Rich-value image part is missing'],
    [
      { 'xl/richData/rdrichvalue.xml': DATA.replace('<v>0</v>', '<v>1</v>') },
      'Rich-value local-image reference is invalid',
    ],
    [
      {
        'xl/richData/rdrichvalue.xml': DATA.replace('<v>0</v>', '<v>bad</v>'),
      },
      'Rich-value local-image reference is invalid',
    ],
    [
      {
        'xl/richData/rdrichvalue.xml': DATA.replace(
          '<v>0</v>',
          '<v><nested/></v>',
        ),
      },
      'Rich-value field is invalid',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELATIONSHIPS.replace(
          '</Relationships>',
          '<Relationship Id="imageRels2" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/></Relationships>',
        ),
      },
      'Rich-value relationship is duplicated',
    ],
    [
      {
        'xl/richData/rdrichvaluestructure.xml': STRUCTURES.replace(
          '_rvRel:LocalImageIdentifier',
          'Wrong',
        ),
      },
      'Rich-value local-image reference is invalid',
    ],
    [
      {
        'xl/richData/rdrichvaluestructure.xml': STRUCTURES.replace(
          'n="_rvRel:LocalImageIdentifier" t="i"',
          'n="_rvRel:LocalImageIdentifier" t="s"',
        ),
      },
      'Rich-value local-image reference is invalid',
    ],
  ] as const)(
    'rejects malformed rich-value image graph %#',
    async (overrides, message) => {
      await expect(
        parseXlsx(await source(overrides), { errorMode: 'strict' }),
      ).rejects.toThrow(message);
    },
  );

  it('never fetches an external image and safely recovers in tolerant mode', async () => {
    const input = await source({
      'xl/richData/_rels/richValueRel.xml.rels': IMAGE_RELATIONSHIPS.replace(
        'Target="../media/image1.png"',
        'Target="https://127.0.0.1:1/secret" TargetMode="External"',
      ),
    });
    const result = await parseXlsxWithDiagnostics(input);
    expect(result.diagnostics).toContainEqual({
      code: 'security-rejected-content',
      message: 'External rich-value images are not loaded',
      part: 'xl/richData/richValueRel.xml',
      severity: 'warning',
    });
    expect(JSON.stringify(result)).not.toContain('127.0.0.1');
  });
});
