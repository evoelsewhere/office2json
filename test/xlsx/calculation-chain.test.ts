import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common/xml/tree';
import {
  parseXlsx,
  parseXlsxWithDiagnostics,
  readXlsxRoundTrip,
  writeXlsxRoundTrip,
} from '../../src/formats/xlsx';
import { parseXlsxCalculationChainPart } from '../../src/formats/xlsx/internal/calculation-chain';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import {
  createIndependentXlsx,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
  type XlsxBlackBoxPart,
} from '../black-box/xlsx-package';

const CALC_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml';
const STRICT_NS = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const STRICT_REL_NS = 'http://purl.oclc.org/ooxml/officeDocument/relationships';

const CONTENT_TYPES = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="${CALC_CONTENT_TYPE}"/>
</Types>`;

const RELATIONSHIPS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
  <Relationship Id="sheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="sheet2" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="styles" Type="${XLSX_OFFICE_REL_TYPE}styles" Target="styles.xml"/>
  <Relationship Id="strings" Type="${XLSX_OFFICE_REL_TYPE}sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="chain" Type="${XLSX_OFFICE_REL_TYPE}calcChain" Target="calcChain.xml"/>
</Relationships>`;

const WORKBOOK = `<workbook xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="First" sheetId="3" r:id="sheet1"/><sheet name="Second" sheetId="9" r:id="sheet2"/></sheets><calcPr/>
</workbook>`;

const CHAIN = `<calcChain xmlns="${XLSX_SPREADSHEET_NS}">
  <c r="a1" i="3" l="1" a="true" s="0" t="false"/>
  <c r="$B$2"/>
  <c r="C3" i="9" l="0" a="0" s="true" t="1"/>
</calcChain>`;

async function source(
  overrides: Record<string, XlsxBlackBoxPart> = {},
): Promise<Uint8Array> {
  return createIndependentXlsx({
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/_rels/workbook.xml.rels': RELATIONSHIPS,
    'xl/calcChain.xml': CHAIN,
    'xl/workbook.xml': WORKBOOK,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>`,
    'xl/worksheets/sheet2.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="3"><c r="C3"><f>2+2</f><v>4</v></c></row></sheetData></worksheet>`,
    ...overrides,
  });
}

describe('XLSX calculation chain metadata', () => {
  it('preserves authored order, inherited sheet IDs, canonical cells, and flags', async () => {
    const document = await parseXlsx(await source(), { errorMode: 'strict' });
    expect(document.workbook.calculation.chain).toStrictEqual([
      {
        address: 'A1',
        arrayFormula: true,
        childChain: false,
        newDependencyLevel: true,
        newThread: false,
        sheetIndex: 0,
      },
      {
        address: 'B2',
        arrayFormula: false,
        childChain: false,
        newDependencyLevel: false,
        newThread: false,
        sheetIndex: 0,
      },
      {
        address: 'C3',
        arrayFormula: false,
        childChain: true,
        newDependencyLevel: false,
        newThread: true,
        sheetIndex: 1,
      },
    ]);
    expect(JSON.parse(JSON.stringify(document))).toStrictEqual(document);
  });

  it('keeps the complete chain independent of range selection', async () => {
    const document = await parseXlsx(await source(), {
      errorMode: 'strict',
      selection: { ranges: { Second: ['C3'] } },
    });
    expect(document.workbook.calculation.chain).toHaveLength(3);
    expect(document.sheets[0]?.payload).toBe('not-selected');
  });

  it('preserves calculation-chain bytes through standalone exact R0', async () => {
    const input = await source();
    const snapshot = await readXlsxRoundTrip(input);
    const output = await writeXlsxRoundTrip(
      JSON.parse(JSON.stringify(snapshot)) as typeof snapshot,
    );
    expect(output.data).toStrictEqual(input);
    expect(output.report.level).toBe('R0');
  });

  it('parses a Strict workbook-owned calculation chain', async () => {
    const strictRelationships = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
      <Relationship Id="sheet1" Type="${STRICT_REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="sheet2" Type="${STRICT_REL_NS}/worksheet" Target="worksheets/sheet2.xml"/>
      <Relationship Id="styles" Type="${STRICT_REL_NS}/styles" Target="styles.xml"/>
      <Relationship Id="strings" Type="${STRICT_REL_NS}/sharedStrings" Target="sharedStrings.xml"/>
      <Relationship Id="chain" Type="${STRICT_REL_NS}/calcChain" Target="calcChain.xml"/>
    </Relationships>`;
    const input = await source({
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="main" Type="${STRICT_REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': strictRelationships,
      'xl/calcChain.xml': CHAIN.replace(XLSX_SPREADSHEET_NS, STRICT_NS),
      'xl/sharedStrings.xml': `<s:sst xmlns:s="${STRICT_NS}" count="0" uniqueCount="0"/>`,
      'xl/styles.xml': `<s:styleSheet xmlns:s="${STRICT_NS}"><s:fonts count="1"><s:font/></s:fonts><s:fills count="1"><s:fill/></s:fills><s:borders count="1"><s:border/></s:borders><s:cellStyleXfs count="1"><s:xf/></s:cellStyleXfs><s:cellXfs count="1"><s:xf/></s:cellXfs></s:styleSheet>`,
      'xl/workbook.xml': `<s:workbook xmlns:s="${STRICT_NS}" xmlns:r="${STRICT_REL_NS}"><s:sheets><s:sheet name="First" sheetId="3" r:id="sheet1"/><s:sheet name="Second" sheetId="9" r:id="sheet2"/></s:sheets><s:calcPr/></s:workbook>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${STRICT_NS}"><s:sheetData/></s:worksheet>`,
      'xl/worksheets/sheet2.xml': `<s:worksheet xmlns:s="${STRICT_NS}"><s:sheetData/></s:worksheet>`,
    });
    const document = await parseXlsx(input, { errorMode: 'strict' });
    expect(
      document.workbook.calculation.chain?.map((entry) => entry.sheetIndex),
    ).toStrictEqual([0, 0, 1]);
  });

  it('omits the optional chain property when no relationship exists', async () => {
    const input = await source({
      'xl/_rels/workbook.xml.rels': RELATIONSHIPS.replace(
        /\s*<Relationship Id="chain"[^>]+\/>/u,
        '',
      ),
      'xl/calcChain.xml': null,
    });
    const document = await parseXlsx(input, { errorMode: 'strict' });
    expect(document.workbook.calculation).not.toHaveProperty('chain');
  });

  it.each([
    [
      {
        'xl/_rels/workbook.xml.rels': RELATIONSHIPS.replace(
          '</Relationships>',
          `<Relationship Id="chain2" Type="${XLSX_OFFICE_REL_TYPE}calcChain" Target="calcChain.xml"/></Relationships>`,
        ),
      },
      'Calculation chain relationship is duplicated',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': RELATIONSHIPS.replace(
          'Target="calcChain.xml"',
          'Target="https://example.test/chain" TargetMode="External"',
        ),
      },
      'Calculation chain relationship must be internal',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          CALC_CONTENT_TYPE,
          'application/xml',
        ),
      },
      'Calculation chain target has the wrong content type',
    ],
    [
      { 'xl/calcChain.xml': CHAIN.replace(XLSX_SPREADSHEET_NS, 'urn:wrong') },
      'Calculation chain root has the wrong namespace',
    ],
    [
      { 'xl/calcChain.xml': `<wrong xmlns="${XLSX_SPREADSHEET_NS}"/>` },
      'Calculation chain root is missing or duplicated',
    ],
    [
      {
        'xl/calcChain.xml': CHAIN.replace(
          '<c r="a1" i="3" l="1" a="true" s="0" t="false"/>',
          '<c>text</c>',
        ),
      },
      'Calculation chain entry collection is invalid',
    ],
    [
      {
        'xl/calcChain.xml': CHAIN.replace(
          '<c r="a1"',
          '<x:c xmlns:x="urn:wrong" r="a1"',
        ),
      },
      'Calculation chain entry has the wrong namespace',
    ],
    [
      { 'xl/calcChain.xml': CHAIN.replace(' i="3"', '') },
      'Calculation chain sheet reference is invalid',
    ],
    [
      { 'xl/calcChain.xml': CHAIN.replace('i="3"', 'i="03"') },
      'Calculation chain sheet reference is invalid',
    ],
    [
      { 'xl/calcChain.xml': CHAIN.replace('i="3"', 'i="4"') },
      'Calculation chain sheet reference is invalid',
    ],
    [
      { 'xl/calcChain.xml': CHAIN.replace('r="a1"', 'r="XFE1"') },
      'Calculation chain cell reference is invalid',
    ],
    [
      { 'xl/calcChain.xml': CHAIN.replace('<c r="$B$2"/>', '<c r="A1"/>') },
      'Calculation chain contains a duplicate cell',
    ],
    [
      { 'xl/calcChain.xml': CHAIN.replace('l="1"', 'l="bad"') },
      'Calculation chain dependency-level flag is invalid',
    ],
    [
      { 'xl/calcChain.xml': CHAIN.replace('a="true"', 'a="bad"') },
      'Calculation chain array flag is invalid',
    ],
    [
      { 'xl/calcChain.xml': CHAIN.replace('s="true"', 's="bad"') },
      'Calculation chain child flag is invalid',
    ],
    [
      { 'xl/calcChain.xml': CHAIN.replace('t="1"', 't="bad"') },
      'Calculation chain thread flag is invalid',
    ],
  ] as const)(
    'rejects malformed calculation-chain package %#',
    async (overrides, message) => {
      await expect(
        parseXlsx(await source(overrides), { errorMode: 'strict' }),
      ).rejects.toThrow(message);
    },
  );

  it('recovers a malformed optional chain only in tolerant mode', async () => {
    const input = await source({
      'xl/calcChain.xml': '<calcChain xmlns="urn:wrong"/>',
    });
    const result = await parseXlsxWithDiagnostics(input);
    expect(result.diagnostics).toContainEqual({
      code: 'invalid-document-structure',
      message: 'Calculation chain root has the wrong namespace',
      part: 'xl/calcChain.xml',
      severity: 'warning',
    });
    expect(result.document.workbook.calculation).not.toHaveProperty('chain');
  });

  it('enforces the calculation-chain entry limit exactly', async () => {
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        limits: { maxCalculationChainEntries: 3 },
      }),
    ).resolves.toBeDefined();
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        limits: { maxCalculationChainEntries: 2 },
      }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'resource-limit-exceeded',
        limitName: 'maxCalculationChainEntries',
      },
    });
  });

  it('parses reserved-prefix Strict normalized chains', () => {
    const result = parseXlsxCalculationChainPart(
      {
        'ns_s:calcChain': {
          attrs: { 'xmlns:s': 'http://purl.oclc.org/ooxml/spreadsheetml/main' },
          'ns_s:c': {
            attrs: {
              'xmlns:s': 'http://purl.oclc.org/ooxml/spreadsheetml/main',
              i: '12',
              r: 'd4',
              t: 'true',
            },
          },
        },
      } as unknown as XmlLookupValue,
      'strict',
      'calcChain.xml',
      new Map([[12, 1]]),
      defaultXlsxResourceLimits(),
    );
    expect(result).toStrictEqual([
      {
        address: 'D4',
        arrayFormula: false,
        childChain: false,
        newDependencyLevel: false,
        newThread: true,
        sheetIndex: 1,
      },
    ]);
  });

  it('accepts explicit empty normalized chains and the UInt32 sheet-ID boundary', () => {
    expect(
      parseXlsxCalculationChainPart(
        {
          calcChain: {
            attrs: { xmlns: XLSX_SPREADSHEET_NS },
            c: undefined,
          },
        } as unknown as XmlLookupValue,
        'transitional',
        'calcChain.xml',
        new Map(),
        defaultXlsxResourceLimits(),
      ),
    ).toStrictEqual([]);
    expect(
      parseXlsxCalculationChainPart(
        {
          calcChain: {
            attrs: { xmlns: XLSX_SPREADSHEET_NS },
            c: { attrs: { i: '4294967295', r: 'A1' } },
          },
        } as unknown as XmlLookupValue,
        'transitional',
        'calcChain.xml',
        new Map([[4294967295, 0]]),
        defaultXlsxResourceLimits(),
      ),
    ).toHaveLength(1);
    expect(() =>
      parseXlsxCalculationChainPart(
        {
          calcChain: {
            attrs: { xmlns: XLSX_SPREADSHEET_NS },
            c: { attrs: { i: '4294967296', r: 'A1' } },
          },
        } as unknown as XmlLookupValue,
        'transitional',
        'calcChain.xml',
        new Map([[4294967296, 0]]),
        defaultXlsxResourceLimits(),
      ),
    ).toThrow('Calculation chain sheet reference is invalid');
  });
});
