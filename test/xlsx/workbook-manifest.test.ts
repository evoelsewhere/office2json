import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { XlsxPartReader } from '../../src/formats/xlsx/internal/part-reader';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from '../../src/formats/xlsx/internal/resource-limits';
import { discoverXlsxWorkbook } from '../../src/formats/xlsx/internal/workbook-discovery';
import { parseXlsxWorkbookManifest } from '../../src/formats/xlsx/internal/workbook-manifest';
import {
  createIndependentXlsx,
  independentWorkbook,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
  type XlsxBlackBoxOverrides,
} from '../black-box/xlsx-package';

const STRICT_SPREADSHEET_NS = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const STRICT_OFFICE_REL_NS =
  'http://purl.oclc.org/ooxml/officeDocument/relationships';
const WORKSHEET_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const CHART_SHEET_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml';
const DEFAULT_WORKBOOK_VIEWS = [
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
] as const;

function workbookRelationships(entries: string): string {
  return `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">${entries}</Relationships>`;
}

function relationship(
  id: string,
  type: string,
  target: string,
  mode?: string,
): string {
  return `<Relationship Id="${id}" Type="${type}" Target="${target}"${
    mode === undefined ? '' : ` TargetMode="${mode}"`
  }/>`;
}

function contentTypes(entries: string): string {
  return `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    ${entries}
  </Types>`;
}

function worksheet(namespace = XLSX_SPREADSHEET_NS): string {
  return `<worksheet xmlns="${namespace}"><sheetData/></worksheet>`;
}

function chartSheet(namespace = XLSX_SPREADSHEET_NS): string {
  return `<chartsheet xmlns="${namespace}"><sheetViews/></chartsheet>`;
}

async function manifest(
  overrides: XlsxBlackBoxOverrides = {},
  limitOverrides: Partial<ResolvedXlsxResourceLimits> = {},
) {
  const zip = await JSZip.loadAsync(await createIndependentXlsx(overrides));
  const limits = { ...defaultXlsxResourceLimits(), ...limitOverrides };
  const reader = new XlsxPartReader(zip, [], limits);
  const discovery = await discoverXlsxWorkbook(reader, limits);
  return parseXlsxWorkbookManifest(discovery, reader, limits);
}

async function captureManifestError(
  overrides: XlsxBlackBoxOverrides,
): Promise<XlsxParseError> {
  try {
    await manifest(overrides);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected workbook manifest parsing to fail');
}

describe('XLSX workbook manifest', () => {
  it('returns the conventional worksheet and default workbook properties', async () => {
    await expect(manifest()).resolves.toEqual({
      pivotCaches: [],
      properties: {
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
        views: DEFAULT_WORKBOOK_VIEWS,
      },
      protectionTextCharacters: 0,
      sheetParts: ['xl/worksheets/sheet1.xml'],
      sheets: [
        {
          columns: [],
          comments: [],
          conditionalFormattings: [],
          dataValidations: [],
          protectedRanges: [],
          drawings: [],
          hyperlinks: [],
          index: 0,
          kind: 'worksheet',
          mergedRanges: [],
          name: 'Sheet1',
          payload: 'full-sheet',
          rows: [],
          state: 'visible',
          tables: [],
          views: [],
        },
      ],
      workbookRelationships: new Map([
        [
          'rIdSheet1',
          {
            id: 'rIdSheet1',
            mode: 'internal',
            target: 'xl/worksheets/sheet1.xml',
            type: `${XLSX_OFFICE_REL_TYPE}worksheet`,
          },
        ],
        [
          'rIdStyles',
          {
            id: 'rIdStyles',
            mode: 'internal',
            target: 'xl/styles.xml',
            type: `${XLSX_OFFICE_REL_TYPE}styles`,
          },
        ],
        [
          'rIdSharedStrings',
          {
            id: 'rIdSharedStrings',
            mode: 'internal',
            target: 'xl/sharedStrings.xml',
            type: `${XLSX_OFFICE_REL_TYPE}sharedStrings`,
          },
        ],
      ]),
    });
  });

  it('applies property defaults when workbookPr and calcPr are absent', async () => {
    const result = await manifest({
      'xl/workbook.xml': `<workbook xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}">
        <sheets><sheet name="Defaults" sheetId="1" r:id="rIdSheet1"/></sheets>
      </workbook>`,
    });

    expect(result.properties).toEqual({
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
      views: DEFAULT_WORKBOOK_VIEWS,
    });
  });

  it('parses workbook and sheet-scoped defined names from authored XML', async () => {
    const workbook = independentWorkbook(
      '<sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/>',
    ).replace(
      '<calcPr',
      `<definedNames>
        <definedName name="Global" hidden="1" comment="A &amp; B">SUM(Sheet1!$A$1:$A$2)</definedName>
        <definedName name="Local" localSheetId="0">Sheet1!$B$2</definedName>
        <definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$2</definedName>
      </definedNames><calcPr`,
    );
    const result = await manifest({ 'xl/workbook.xml': workbook });

    expect(result.properties.definedNames).toEqual([
      {
        comment: 'A & B',
        expression: 'SUM(Sheet1!$A$1:$A$2)',
        hidden: true,
        name: 'Global',
      },
      {
        expression: 'Sheet1!$B$2',
        hidden: false,
        name: 'Local',
        sheetIndex: 0,
      },
      {
        expression: 'Sheet1!$1:$2',
        hidden: false,
        name: '_xlnm.Print_Titles',
        sheetIndex: 0,
      },
    ]);
  });

  it('preserves authored order, sheet kinds, states, and calculation flags', async () => {
    const workbook = `
      <workbook xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}">
        <workbookPr date1904="true"/>
        <bookViews><workbookView activeTab="0" firstSheet="0"
          autoFilterDateGrouping="0" minimized="1"
          showHorizontalScroll="0" showSheetTabs="false"
          showVerticalScroll="true" tabRatio="750" visibility="hidden"
          windowHeight="800" windowWidth="1200" xWindow="-20" yWindow="30"/>
        </bookViews>
        <sheets>
          <sheet name="Visible" sheetId="7" r:id="rId1"/>
          <sheet name="Chart" sheetId="9" state="hidden" r:id="rId2"/>
          <sheet name="Archive" sheetId="12" state="veryHidden" r:id="rId3"/>
        </sheets>
        <calcPr calcId="4294967295" calcMode="manual" refMode="R1C1" iterate="1" iterateCount="250" iterateDelta="0.25" fullPrecision="0" calcCompleted="false" calcOnSave="0" concurrentCalc="false" concurrentManualCount="7" forceFullCalc="1" fullCalcOnLoad="true"/>
      </workbook>`;
    const relBase = XLSX_OFFICE_REL_TYPE;
    const result = await manifest({
      '[Content_Types].xml': contentTypes(`
        <Override PartName="/xl/worksheets/visible.xml" ContentType="${WORKSHEET_CONTENT_TYPE}"/>
        <Override PartName="/xl/chartsheets/chart.xml" ContentType="${CHART_SHEET_CONTENT_TYPE}"/>
        <Override PartName="/xl/worksheets/archive.xml" ContentType="${WORKSHEET_CONTENT_TYPE}"/>`),
      'xl/_rels/workbook.xml.rels': workbookRelationships(`
        ${relationship('rId1', `${relBase}worksheet`, 'worksheets/visible.xml')}
        ${relationship('rId2', `${relBase}chartsheet`, 'chartsheets/chart.xml')}
        ${relationship('rId3', `${relBase}worksheet`, 'worksheets/archive.xml')}`),
      'xl/chartsheets/chart.xml': chartSheet(),
      'xl/workbook.xml': workbook,
      'xl/worksheets/archive.xml': worksheet(),
      'xl/worksheets/sheet1.xml': null,
      'xl/worksheets/visible.xml': worksheet(),
    });

    expect(result.properties).toEqual({
      calculation: {
        calculationCompleted: false,
        calculationId: 4294967295,
        calculateOnSave: false,
        concurrentCalculation: false,
        concurrentManualCount: 7,
        forceFullCalculation: true,
        fullCalculationOnLoad: true,
        fullPrecision: false,
        iteration: {
          enabled: true,
          maxChange: 0.25,
          maxIterations: 250,
        },
        mode: 'manual',
        referenceMode: 'R1C1',
      },
      dateSystem: '1904',
      definedNames: [],
      views: [
        {
          activeSheetIndex: 0,
          autoFilterDateGrouping: false,
          firstVisibleSheetIndex: 0,
          minimized: true,
          showHorizontalScroll: false,
          showSheetTabs: false,
          showVerticalScroll: true,
          tabRatio: 750,
          visibility: 'hidden',
          windowHeight: 800,
          windowWidth: 1_200,
          xWindow: -20,
          yWindow: 30,
        },
      ],
    });
    expect(
      result.sheets.map(({ index, kind, name, state }) => ({
        index,
        kind,
        name,
        state,
      })),
    ).toEqual([
      { index: 0, kind: 'worksheet', name: 'Visible', state: 'visible' },
      { index: 1, kind: 'chart-sheet', name: 'Chart', state: 'hidden' },
      { index: 2, kind: 'worksheet', name: 'Archive', state: 'very-hidden' },
    ]);
    expect(result.sheetParts).toEqual([
      'xl/worksheets/visible.xml',
      'xl/chartsheets/chart.xml',
      'xl/worksheets/archive.xml',
    ]);
  });

  it('parses prefixed Strict workbook and sheet roots', async () => {
    const strictRelBase = `${STRICT_OFFICE_REL_NS}/`;
    const strictWorkbook = `<s:workbook xmlns:s="${STRICT_SPREADSHEET_NS}" xmlns:q="${STRICT_OFFICE_REL_NS}">
      <s:workbookPr date1904="false"/>
      <s:bookViews><s:workbookView activeTab="0" firstSheet="0" minimized="true"/></s:bookViews>
      <s:sheets>
        <s:sheet name="Strict data" sheetId="1" q:id="sheet"/>
        <s:sheet name="Strict chart" sheetId="2" q:id="chart"/>
      </s:sheets>
      <s:definedNames><s:definedName name="StrictName">'Strict data'!$A$1</s:definedName></s:definedNames>
      <s:calcPr calcMode="autoNoTable" forceFullCalc="false" fullCalcOnLoad="0"/>
    </s:workbook>`;
    const strictSheet = `<s:worksheet xmlns:s="${STRICT_SPREADSHEET_NS}"><s:sheetData/></s:worksheet>`;
    const strictChart = `<s:chartsheet xmlns:s="${STRICT_SPREADSHEET_NS}"/>`;
    const result = await manifest({
      '[Content_Types].xml': contentTypes(`
        <Override PartName="/xl/worksheets/strict.xml" ContentType="${WORKSHEET_CONTENT_TYPE}"/>
        <Override PartName="/xl/chartsheets/strict.xml" ContentType="${CHART_SHEET_CONTENT_TYPE}"/>`),
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">${relationship(
        'main',
        `${strictRelBase}officeDocument`,
        'xl/workbook.xml',
      )}</Relationships>`,
      'xl/_rels/workbook.xml.rels': workbookRelationships(`
        ${relationship('sheet', `${strictRelBase}worksheet`, 'worksheets/strict.xml')}
        ${relationship('chart', `${strictRelBase}chartsheet`, 'chartsheets/strict.xml')}`),
      'xl/chartsheets/strict.xml': strictChart,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/workbook.xml': strictWorkbook,
      'xl/worksheets/sheet1.xml': null,
      'xl/worksheets/strict.xml': strictSheet,
    });

    expect(result.properties.calculation.mode).toBe('automatic-except-tables');
    expect(result.properties.definedNames).toEqual([
      {
        expression: "'Strict data'!$A$1",
        hidden: false,
        name: 'StrictName',
      },
    ]);
    expect(result.properties.views).toMatchObject([
      { activeSheetIndex: 0, firstVisibleSheetIndex: 0, minimized: true },
    ]);
    expect(result.sheets.map((sheet) => sheet.kind)).toEqual([
      'worksheet',
      'chart-sheet',
    ]);
    expect(result.sheetParts).toEqual([
      'xl/worksheets/strict.xml',
      'xl/chartsheets/strict.xml',
    ]);
  });

  it('describes chart sheets without reading their unselected payloads', async () => {
    const result = await manifest({
      '[Content_Types].xml': contentTypes(
        `<Override PartName="/xl/chartsheets/chart.xml" ContentType="${CHART_SHEET_CONTENT_TYPE}"/>`,
      ),
      'xl/_rels/workbook.xml.rels': workbookRelationships(
        relationship(
          'rIdSheet1',
          `${XLSX_OFFICE_REL_TYPE}chartsheet`,
          'chartsheets/chart.xml',
        ),
      ),
      'xl/chartsheets/chart.xml': null,
    });

    expect(result.sheets).toEqual([
      {
        index: 0,
        kind: 'chart-sheet',
        name: 'Sheet1',
        payload: 'full-sheet',
        state: 'visible',
      },
    ]);
    expect(result.sheetParts).toEqual(['xl/chartsheets/chart.xml']);
  });

  it('accepts exactly maxWorksheets and rejects one over', async () => {
    const workbook = independentWorkbook(`
      <sheet name="One" sheetId="1" r:id="rId1"/>
      <sheet name="Two" sheetId="2" r:id="rId2"/>`);
    const overrides = {
      '[Content_Types].xml': contentTypes(`
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="${WORKSHEET_CONTENT_TYPE}"/>
        <Override PartName="/xl/worksheets/sheet2.xml" ContentType="${WORKSHEET_CONTENT_TYPE}"/>`),
      'xl/_rels/workbook.xml.rels': workbookRelationships(`
        ${relationship('rId1', `${XLSX_OFFICE_REL_TYPE}worksheet`, 'worksheets/sheet1.xml')}
        ${relationship('rId2', `${XLSX_OFFICE_REL_TYPE}worksheet`, 'worksheets/sheet2.xml')}`),
      'xl/workbook.xml': workbook,
      'xl/worksheets/sheet2.xml': worksheet(),
    };

    await expect(
      manifest(overrides, { maxWorksheets: 2 }),
    ).resolves.toMatchObject({
      sheets: [{ name: 'One' }, { name: 'Two' }],
    });
    await expect(
      manifest(overrides, { maxWorksheets: 1 }),
    ).rejects.toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxWorksheets',
      name: 'XlsxResourceLimitError',
      part: 'xl/workbook.xml',
    } satisfies Partial<XlsxResourceLimitError>);
  });

  it.each([
    'a'.repeat(31),
    'Space allowed',
    `High ${String.fromCodePoint(0x80)}`,
  ])('accepts valid sheet-name boundary %s', async (name) => {
    const result = await manifest({
      'xl/workbook.xml': independentWorkbook(
        `<sheet name="${name}" sheetId="4294967295" state="visible" r:id="rIdSheet1"/>`,
      ),
    });

    expect(result.sheets[0]).toMatchObject({ name, state: 'visible' });
  });

  it.each(['\\', '/', ':', '?', '*', '[', ']', String.fromCodePoint(0x7f)])(
    'rejects prohibited sheet-name character %#',
    async (character) => {
      const error = await captureManifestError({
        'xl/workbook.xml': independentWorkbook(
          `<sheet name="Bad${character}name" sheetId="1" r:id="rIdSheet1"/>`,
        ),
      });

      expect(error.diagnostic.message).toBe(
        'Workbook sheet has an invalid name',
      );
    },
  );

  it.each(['01', '+1', '1.0', '4294967296', '9007199254740992'])(
    'rejects invalid sheetId lexical form %s',
    async (sheetId) => {
      const error = await captureManifestError({
        'xl/workbook.xml': independentWorkbook(
          `<sheet name="Data" sheetId="${sheetId}" r:id="rIdSheet1"/>`,
        ),
      });

      expect(error.diagnostic.message).toBe(
        'Workbook sheet has an invalid sheetId',
      );
    },
  );

  it.each([
    [
      { 'xl/_rels/workbook.xml.rels': null },
      'Required XLSX part is missing: xl/_rels/workbook.xml.rels',
      'missing-required-part',
      'xl/_rels/workbook.xml.rels',
    ],
    [
      { 'xl/workbook.xml': `<workbook xmlns="${XLSX_SPREADSHEET_NS}"/>` },
      'Workbook sheets collection is missing',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': `<workbook xmlns="${XLSX_SPREADSHEET_NS}"><sheets/></workbook>`,
      },
      'Workbook must contain at least one sheet',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': `<workbook xmlns="${XLSX_SPREADSHEET_NS}"><sheets><sheet>text</sheet></sheets></workbook>`,
      },
      'Workbook must contain at least one sheet',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet sheetId="1" r:id="rIdSheet1"/>',
        ),
      },
      'Workbook sheet has an invalid name',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet name="" sheetId="1" r:id="rIdSheet1"/>',
        ),
      },
      'Workbook sheet has an invalid name',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          `<sheet name="${'a'.repeat(32)}" sheetId="1" r:id="rIdSheet1"/>`,
        ),
      },
      'Workbook sheet has an invalid name',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet name="Bad/name" sheetId="1" r:id="rIdSheet1"/>',
        ),
      },
      'Workbook sheet has an invalid name',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(`
          <sheet name="Data" sheetId="1" r:id="rIdSheet1"/>
          <sheet name="data" sheetId="2" r:id="rIdSheet1"/>`),
      },
      'Workbook contains duplicate sheet names',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(`
          <sheet name="σ" sheetId="1" r:id="rIdSheet1"/>
          <sheet name="ς" sheetId="2" r:id="rIdSheet1"/>`),
      },
      'Workbook contains duplicate sheet names',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet name="Data" r:id="rIdSheet1"/>',
        ),
      },
      'Workbook sheet has an invalid sheetId',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet name="Data" sheetId="1" r:id=""/>',
        ),
      },
      'Workbook sheet has an invalid relationship reference',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet name="Data" sheetId="0" r:id="rIdSheet1"/>',
        ),
      },
      'Workbook sheet has an invalid sheetId',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(`
          <sheet name="One" sheetId="1" r:id="rIdSheet1"/>
          <sheet name="Two" sheetId="1" r:id="rIdSheet1"/>`),
      },
      'Workbook contains duplicate sheetId values',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet name="Data" sheetId="1"/>',
        ),
      },
      'Workbook sheet has an invalid relationship reference',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet name="Data" sheetId="1" r:id="missing"/>',
        ),
      },
      'Workbook sheet relationship is missing or external',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': workbookRelationships(
          relationship(
            'rIdSheet1',
            `${XLSX_OFFICE_REL_TYPE}worksheet`,
            'https://example.com/sheet.xml',
            'External',
          ),
        ),
      },
      'Workbook sheet relationship is missing or external',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': workbookRelationships(
          relationship(
            'rIdSheet1',
            `${XLSX_OFFICE_REL_TYPE}dialogsheet`,
            'worksheets/sheet1.xml',
          ),
        ),
      },
      'Workbook sheet relationship has an unsupported type',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        '[Content_Types].xml': contentTypes(
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/xml"/>',
        ),
      },
      'Workbook sheet target has the wrong content type',
      'invalid-document-structure',
      'xl/worksheets/sheet1.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet name="Data" sheetId="1" state="shown" r:id="rIdSheet1"/>',
        ),
      },
      'Workbook sheet state is invalid',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet name="Data" sheetId="1" r:id="rIdSheet1"/>',
        ).replace('date1904="0"', 'date1904="yes"'),
      },
      'Workbook date1904 flag is invalid',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet name="Data" sheetId="1" r:id="rIdSheet1"/>',
        ).replace('calcMode="auto"', 'calcMode="automatic"'),
      },
      'Workbook calculation mode is invalid',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet name="Data" sheetId="1" r:id="rIdSheet1"/>',
        ).replace('forceFullCalc="0"', 'forceFullCalc="yes"'),
      },
      'Workbook force-full-calculation flag is invalid',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': independentWorkbook(
          '<sheet name="Data" sheetId="1" r:id="rIdSheet1"/>',
        ).replace('fullCalcOnLoad="0"', 'fullCalcOnLoad="yes"'),
      },
      'Workbook full-calculation-on-load flag is invalid',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
  ] as const)(
    'rejects invalid manifest %#',
    async (overrides, message, code, part) => {
      const error = await captureManifestError(overrides);
      expect(error.diagnostic).toEqual({
        code,
        message,
        part,
        severity: 'error',
      });
    },
  );

  it.each([
    ['calcId="-1"', 'Workbook calculation ID is invalid'],
    ['calcId="4294967296"', 'Workbook calculation ID is invalid'],
    [
      'concurrentManualCount="4294967296"',
      'Workbook concurrent manual count is invalid',
    ],
    ['iterateCount="bad"', 'Workbook iteration count is invalid'],
    ['iterateCount="4294967296"', 'Workbook iteration count is invalid'],
    ['iterateDelta="bad"', 'Workbook iteration delta is invalid'],
    ['iterateDelta="1e999"', 'Workbook iteration delta is invalid'],
    ['iterateDelta=" 1"', 'Workbook iteration delta is invalid'],
    ['iterateDelta="1 "', 'Workbook iteration delta is invalid'],
    ['refMode="bad"', 'Workbook calculation reference mode is invalid'],
    ['calcCompleted="bad"', 'Workbook calculation-completed flag is invalid'],
    ['calcOnSave="bad"', 'Workbook calculate-on-save flag is invalid'],
    ['concurrentCalc="bad"', 'Workbook concurrent-calculation flag is invalid'],
    ['fullPrecision="bad"', 'Workbook full-precision flag is invalid'],
    ['iterate="bad"', 'Workbook iteration flag is invalid'],
  ] as const)(
    'rejects invalid calcPr attribute %s',
    async (attribute, message) => {
      const workbook = independentWorkbook(
        '<sheet name="Data" sheetId="1" r:id="rIdSheet1"/>',
      ).replace('<calcPr ', `<calcPr ${attribute} `);
      const error = await captureManifestError({ 'xl/workbook.xml': workbook });
      expect(error.diagnostic).toEqual({
        code: 'invalid-document-structure',
        message,
        part: 'xl/workbook.xml',
        severity: 'error',
      });
    },
  );

  it.each([
    ['+1', 1],
    ['12', 12],
    ['1.', 1],
    ['.55', 0.55],
    ['1e3', 1000],
    ['1e30', 1e30],
    ['1e+3', 1000],
    ['1E-3', 0.001],
    ['-0', 0],
  ] as const)(
    'parses calculation iteration delta %s',
    async (value, expected) => {
      const workbook = independentWorkbook(
        '<sheet name="Data" sheetId="1" r:id="rIdSheet1"/>',
      ).replace('<calcPr ', `<calcPr iterateDelta="${value}" `);
      const result = await manifest({ 'xl/workbook.xml': workbook });
      expect(result.properties.calculation.iteration.maxChange).toBe(expected);
      if (value === '-0') {
        expect(
          Object.is(result.properties.calculation.iteration.maxChange, -0),
        ).toBe(false);
      }
    },
  );

  it('accepts UInt32 maxima for iteration and concurrent-manual counts', async () => {
    const workbook = independentWorkbook(
      '<sheet name="Data" sheetId="1" r:id="rIdSheet1"/>',
    ).replace(
      '<calcPr ',
      '<calcPr iterateCount="4294967295" concurrentManualCount="4294967295" ',
    );
    const result = await manifest({ 'xl/workbook.xml': workbook });
    expect(result.properties.calculation).toMatchObject({
      concurrentManualCount: 4294967295,
      iteration: { maxIterations: 4294967295 },
    });
  });
});
