import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
  XlsxParseError,
} from '../../src/formats/xlsx';
import {
  defaultXlsxResourceLimits,
  XlsxResourceLimitError,
} from '../../src/formats/xlsx/internal/resource-limits';
import type { XlsxXmlElement } from '../../src/formats/xlsx/internal/streaming-xml';
import {
  parseXlsxPageMargins,
  parseXlsxPageSetup,
  parseXlsxPageSetupProperties,
  parseXlsxPrintOptions,
  XlsxHeaderFooterCapture,
  XlsxPageBreaksCapture,
} from '../../src/formats/xlsx/internal/worksheet-print';
import type { XlsxWorksheetBudget } from '../../src/formats/xlsx/internal/worksheet';
import {
  createIndependentXlsx,
  independentWorkbook,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const PART = 'xl/worksheets/sheet1.xml';

function budget(): XlsxWorksheetBudget {
  return {
    conditionalFormattingRules: 0,
    formulaCharacters: 0,
    formulaGroups: 0,
    rangeAreas: 0,
    returnedCells: 0,
    richTextRuns: 0,
    scannedCells: 0,
    textCharacters: 0,
    validationRules: 0,
  };
}

function element(
  localName: string,
  attrs: Record<string, string> = {},
): XlsxXmlElement {
  return {
    attributes: new Map(
      Object.entries(attrs).map(([name, value]) => [`{}${name}`, value]),
    ),
    localName,
    namespace: 'worksheet',
  };
}

function worksheet(print: string): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}">${print.replace(
    '<!--sheetData-->',
    '<sheetData/>',
  )}</worksheet>`;
}

function workbook(definedNames = ''): string {
  return independentWorkbook(
    '<sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/>',
  ).replace('</sheets>', `</sheets>${definedNames}`);
}

async function bytes(print: string, definedNames = ''): Promise<Uint8Array> {
  return createIndependentXlsx({
    'xl/workbook.xml': workbook(definedNames),
    'xl/worksheets/sheet1.xml': worksheet(print),
  });
}

async function parsePrint(
  print: string,
  options: Parameters<typeof parseXlsx>[1] = {},
  definedNames = '',
) {
  return parseXlsx(await bytes(print, definedNames), options);
}

async function capture(
  print: string,
  options: Parameters<typeof parseXlsx>[1] = {},
): Promise<XlsxParseError> {
  try {
    await parsePrint(print, options);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected XLSX print parsing to fail');
}

function captureTree(action: () => unknown): XlsxParseError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected normalized XLSX print parsing to fail');
}

function captureLimit(action: () => unknown): XlsxResourceLimitError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxResourceLimitError);
    return error as XlsxResourceLimitError;
  }
  throw new Error('Expected normalized XLSX print resource failure');
}

const COMPLETE_PRINT = `<sheetPr><pageSetUpPr autoPageBreaks="0" fitToPage="1"/></sheetPr>
<!--sheetData-->
<printOptions gridLines="1" gridLinesSet="true" headings="1" horizontalCentered="true" verticalCentered="1"/>
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
<pageSetup blackAndWhite="1" cellComments="atEnd" copies="2" draft="1" errors="NA" firstPageNumber="3" fitToHeight="4" fitToWidth="2" horizontalDpi="600" orientation="landscape" pageOrder="overThenDown" paperHeight="297mm" paperSize="9" paperWidth="21cm" scale="85" useFirstPageNumber="1" usePrinterDefaults="1" verticalDpi="600"/>
<headerFooter alignWithMargins="0" differentFirst="1" differentOddEven="1" scaleWithDoc="0"><oddHeader>&amp;LRevenue &amp; Profit</oddHeader><oddFooter>&amp;P / &amp;N</oddFooter><evenHeader>Even</evenHeader><evenFooter>Even footer</evenFooter><firstHeader>First</firstHeader><firstFooter>First footer</firstFooter></headerFooter>
<rowBreaks count="2" manualBreakCount="1"><brk id="20" min="0" max="10" man="1"/><brk id="40" min="1" max="5" pt="1"/></rowBreaks>
<colBreaks count="1" manualBreakCount="1"><brk id="3" min="0" max="20" man="true"/></colBreaks>`;

describe('XLSX worksheet print settings', () => {
  it('parses complete print settings and repeating-title defined names', async () => {
    const document = await parsePrint(
      COMPLETE_PRINT,
      {},
      '<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$2</definedName></definedNames>',
    );
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    expect(sheet.kind === 'worksheet' ? sheet.print : undefined).toEqual({
      columnBreaks: [
        { end: 20, manual: true, pivot: false, position: 3, start: 0 },
      ],
      headerFooter: {
        alignWithMargins: false,
        differentFirst: true,
        differentOddEven: true,
        evenFooter: 'Even footer',
        evenHeader: 'Even',
        firstFooter: 'First footer',
        firstHeader: 'First',
        oddFooter: '&P / &N',
        oddHeader: '&LRevenue & Profit',
        scaleWithDocument: false,
      },
      margins: {
        bottom: 0.75,
        footer: 0.3,
        header: 0.3,
        left: 0.7,
        right: 0.7,
        top: 0.75,
      },
      options: {
        gridLines: true,
        gridLinesSet: true,
        headings: true,
        horizontalCentered: true,
        verticalCentered: true,
      },
      pageSetup: {
        blackAndWhite: true,
        cellComments: 'at-end',
        copies: 2,
        draft: true,
        errors: 'not-available',
        firstPageNumber: 3,
        fitToHeight: 4,
        fitToWidth: 2,
        horizontalDpi: 600,
        orientation: 'landscape',
        pageOrder: 'over-then-down',
        paperHeight: { unit: 'mm', value: 297 },
        paperSize: 9,
        paperWidth: { unit: 'cm', value: 21 },
        scale: 85,
        useFirstPageNumber: true,
        usePrinterDefaults: true,
        verticalDpi: 600,
      },
      properties: { autoPageBreaks: false, fitToPage: true },
      rowBreaks: [
        { end: 10, manual: true, pivot: false, position: 20, start: 0 },
        { end: 5, manual: false, pivot: true, position: 40, start: 1 },
      ],
    });
    expect(document.workbook.definedNames).toMatchObject([
      { expression: 'Sheet1!$1:$2', name: '_xlnm.Print_Titles', sheetIndex: 0 },
    ]);
    expect(JSON.parse(JSON.stringify(document))).toEqual(document);
  });

  it('emits explicit element defaults without inventing unauthored print state', async () => {
    const absent = await parsePrint('<!--sheetData-->');
    const absentSheet = absent.sheets[0]!;
    expect(
      absentSheet.kind === 'worksheet' ? absentSheet.print : undefined,
    ).toBeUndefined();

    const document = await parsePrint(
      '<sheetPr><pageSetUpPr/></sheetPr><!--sheetData--><printOptions/><pageSetup/><headerFooter/><rowBreaks/><colBreaks/>',
    );
    const sheet = document.sheets[0]!;
    expect(sheet.kind === 'worksheet' ? sheet.print : undefined).toStrictEqual({
      columnBreaks: [],
      headerFooter: {
        alignWithMargins: true,
        differentFirst: false,
        differentOddEven: false,
        scaleWithDocument: true,
      },
      options: {
        gridLines: false,
        gridLinesSet: false,
        headings: false,
        horizontalCentered: false,
        verticalCentered: false,
      },
      pageSetup: {
        blackAndWhite: false,
        cellComments: 'none',
        draft: false,
        errors: 'displayed',
        orientation: 'default',
        pageOrder: 'down-then-over',
        useFirstPageNumber: false,
        usePrinterDefaults: false,
      },
      properties: { autoPageBreaks: true, fitToPage: false },
      rowBreaks: [],
    });
  });

  it.each([
    ['<!--sheetData--><colBreaks/>', 'columnBreaks'],
    ['<!--sheetData--><headerFooter/>', 'headerFooter'],
    [
      '<!--sheetData--><pageMargins bottom="1" footer="1" header="1" left="1" right="1" top="1"/>',
      'margins',
    ],
    ['<!--sheetData--><pageSetup/>', 'pageSetup'],
    ['<sheetPr><pageSetUpPr/></sheetPr><!--sheetData-->', 'properties'],
    ['<!--sheetData--><printOptions/>', 'options'],
    ['<!--sheetData--><rowBreaks/>', 'rowBreaks'],
  ] as const)('emits isolated print component %s', async (xml, property) => {
    const document = await parsePrint(xml);
    const sheet = document.sheets[0]!;
    const print = sheet.kind === 'worksheet' ? sheet.print : undefined;
    expect(print).toBeDefined();
    expect(Object.keys(print!)).toEqual([property]);
  });

  it('round-trips print metadata through portable exact R0', async () => {
    const source = await bytes(COMPLETE_PRINT);
    const snapshot = await readXlsxRoundTrip(source);
    const output = await writeXlsxRoundTrip(
      await validateXlsxRoundTripJson(
        JSON.parse(JSON.stringify(snapshot)) as unknown,
      ),
    );
    expect(output.data).toEqual(source);
    expect(output.report.level).toBe('R0');
  });

  it('parses Strict prefixed print elements', async () => {
    const strictSheet = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictRelationship =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const source = await createIndependentXlsx({
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelationship}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${strictRelationship}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheet}" xmlns:r="${strictRelationship}"><s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets></s:workbook>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheet}"><s:sheetPr><s:pageSetUpPr fitToPage="1"/></s:sheetPr><s:sheetData/><s:pageMargins left="1" right="1" top="1" bottom="1" header="0.5" footer="0.5"/><s:headerFooter><s:oddHeader>Strict</s:oddHeader></s:headerFooter></s:worksheet>`,
    });
    const document = await parseXlsx(source);
    const sheet = document.sheets[0]!;
    expect(sheet.kind === 'worksheet' ? sheet.print : undefined).toMatchObject({
      headerFooter: { oddHeader: 'Strict' },
      properties: { fitToPage: true },
    });
  });

  it.each([
    [
      '<printOptions gridLines="bad"/>',
      'Worksheet print grid-lines flag is invalid',
    ],
    ['<pageMargins/>', 'Worksheet page bottom margin is missing'],
    [
      '<pageMargins bottom="-1" footer="0" header="0" left="0" right="0" top="0"/>',
      'Worksheet page bottom margin is invalid',
    ],
    [
      '<pageSetup orientation="wide"/>',
      'Worksheet page orientation is invalid',
    ],
    ['<pageSetup scale="9"/>', 'Worksheet page scale is invalid'],
    ['<pageSetup scale="401"/>', 'Worksheet page scale is invalid'],
    ['<pageSetup paperSize="0"/>', 'Worksheet paper size is invalid'],
    ['<pageSetup copies="0"/>', 'Worksheet print copies are invalid'],
    ['<pageSetup paperWidth="0mm"/>', 'Worksheet paper width is invalid'],
    ['<pageSetup fitToWidth="01"/>', 'Worksheet fit-to-width value is invalid'],
    [
      '<headerFooter differentFirst="bad"/>',
      'Worksheet first-header/footer flag is invalid',
    ],
    [
      '<headerFooter><oddHeader>One</oddHeader><oddHeader>Two</oddHeader></headerFooter>',
      'Worksheet header/footer contains duplicate fields',
    ],
    [
      '<headerFooter>bad<oddHeader>One</oddHeader></headerFooter>',
      'Worksheet header/footer text is invalid',
    ],
    [
      '<rowBreaks><brk/></rowBreaks>',
      'Worksheet page-break position is invalid',
    ],
    [
      '<rowBreaks count="2"><brk id="1"/></rowBreaks>',
      'Worksheet page-break count does not match',
    ],
    [
      '<rowBreaks manualBreakCount="1"><brk id="1"/></rowBreaks>',
      'Worksheet manual page-break count does not match',
    ],
    [
      '<rowBreaks><brk id="1"/><brk id="1"/></rowBreaks>',
      'Worksheet contains duplicate page-break positions',
    ],
    [
      '<rowBreaks><brk id="1" min="2" max="1"/></rowBreaks>',
      'Worksheet page-break extent is invalid',
    ],
    [
      '<rowBreaks><brk id="1" man="bad"/></rowBreaks>',
      'Worksheet manual page-break flag is invalid',
    ],
    [
      '<rowBreaks>bad<brk id="1"/></rowBreaks>',
      'Worksheet page-break text is invalid',
    ],
    [
      '<printOptions/><!--sheetData--><printOptions/>',
      'Worksheet contains duplicate printOptions elements',
    ],
    [
      '<!--sheetData--><pageMargins bottom="1" footer="1" header="1" left="1" right="1" top="1"/><pageMargins bottom="1" footer="1" header="1" left="1" right="1" top="1"/>',
      'Worksheet contains duplicate pageMargins elements',
    ],
    [
      '<!--sheetData--><pageSetup/><pageSetup/>',
      'Worksheet contains duplicate pageSetup elements',
    ],
    [
      '<!--sheetData--><headerFooter/><headerFooter/>',
      'Worksheet contains duplicate headerFooter elements',
    ],
    [
      '<!--sheetData--><rowBreaks/><rowBreaks/>',
      'Worksheet contains duplicate rowBreaks elements',
    ],
    [
      '<!--sheetData--><colBreaks/><colBreaks/>',
      'Worksheet contains duplicate colBreaks elements',
    ],
    [
      '<sheetPr><pageSetUpPr/><pageSetUpPr/></sheetPr><!--sheetData-->',
      'Worksheet contains duplicate page setup properties',
    ],
  ] as const)('rejects invalid print contract %#', async (xml, message) => {
    const source = xml.includes('<!--sheetData-->')
      ? xml
      : `<!--sheetData-->${xml}`;
    expect((await capture(source)).diagnostic.message).toBe(message);
  });

  it('enforces header/footer text, page-break, and grid limits exactly', async () => {
    const exact = await parsePrint(
      '<!--sheetData--><headerFooter><oddHeader>abcd</oddHeader></headerFooter><rowBreaks><brk id="2" min="0" max="1"/></rowBreaks>',
      {
        limits: {
          maxColumnsPerWorksheet: 2,
          maxRangeAreas: 1,
          maxRowsPerWorksheet: 2,
          maxTextCharacters: 13,
        },
      },
    );
    expect(exact.sheets).toHaveLength(1);
    expect(
      (
        await capture(
          '<!--sheetData--><headerFooter><oddHeader>abcd</oddHeader></headerFooter>',
          { limits: { maxTextCharacters: 12 } },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 13,
      code: 'resource-limit-exceeded',
      limit: 12,
      limitName: 'maxTextCharacters',
    });
    expect(
      (
        await capture(
          '<!--sheetData--><rowBreaks count="2"><brk id="1"/><brk id="2"/></rowBreaks>',
          { limits: { maxRangeAreas: 1 } },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      code: 'resource-limit-exceeded',
      limit: 1,
      limitName: 'maxRangeAreas',
    });
    expect(
      (
        await capture('<!--sheetData--><rowBreaks count="2"/>', {
          limits: { maxRangeAreas: 1 },
        })
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      code: 'resource-limit-exceeded',
      limit: 1,
      limitName: 'maxRangeAreas',
    });
    for (const [xml, limits, limitName] of [
      [
        '<!--sheetData--><rowBreaks><brk id="3"/></rowBreaks>',
        { maxRowsPerWorksheet: 2 },
        'maxRowsPerWorksheet',
      ],
      [
        '<!--sheetData--><colBreaks><brk id="3"/></colBreaks>',
        { maxColumnsPerWorksheet: 2 },
        'maxColumnsPerWorksheet',
      ],
      [
        '<!--sheetData--><rowBreaks><brk id="1" max="2"/></rowBreaks>',
        { maxColumnsPerWorksheet: 2 },
        'maxColumnsPerWorksheet',
      ],
    ] as const) {
      expect((await capture(xml, { limits })).diagnostic).toMatchObject({
        actual: 3,
        code: 'resource-limit-exceeded',
        limit: 2,
        limitName,
      });
    }
  });
});

describe('XLSX normalized print helpers and captures', () => {
  it.each([
    ['gridLines', 'Worksheet print grid-lines flag is invalid'],
    ['gridLinesSet', 'Worksheet print grid-lines-set flag is invalid'],
    ['headings', 'Worksheet print headings flag is invalid'],
    [
      'horizontalCentered',
      'Worksheet print horizontal-centering flag is invalid',
    ],
    ['verticalCentered', 'Worksheet print vertical-centering flag is invalid'],
  ] as const)('rejects invalid normalized print option %s', (name, message) => {
    expect(
      captureTree(() =>
        parseXlsxPrintOptions(element('printOptions', { [name]: 'bad' }), PART),
      ).diagnostic.message,
    ).toBe(message);
  });

  it.each([
    ['bottom', 'Worksheet page bottom margin is invalid'],
    ['footer', 'Worksheet page footer margin is invalid'],
    ['header', 'Worksheet page header margin is invalid'],
    ['left', 'Worksheet page left margin is invalid'],
    ['right', 'Worksheet page right margin is invalid'],
    ['top', 'Worksheet page top margin is invalid'],
  ] as const)('rejects invalid normalized page margin %s', (name, message) => {
    expect(
      captureTree(() =>
        parseXlsxPageMargins(
          element('pageMargins', {
            bottom: '1',
            footer: '1',
            header: '1',
            left: '1',
            right: '1',
            top: '1',
            [name]: 'bad',
          }),
          PART,
        ),
      ).diagnostic.message,
    ).toBe(message);
  });

  it('accepts every finite margin lexical family and rejects its boundaries', () => {
    const parsed = (bottom: string) =>
      parseXlsxPageMargins(
        element('pageMargins', {
          bottom,
          footer: '0',
          header: '0',
          left: '0',
          right: '0',
          top: '0',
        }),
        PART,
      ).bottom;
    for (const [source, expected] of [
      ['+1', 1],
      ['12', 12],
      ['1.', 1],
      ['.5', 0.5],
      ['.55', 0.55],
      ['1e3', 1000],
      ['1e30', 1e30],
      ['2.5E-2', 0.025],
    ] as const) {
      expect(parsed(source)).toBe(expected);
    }
    for (const source of [
      ' 1',
      '1 ',
      '1x',
      'x1',
      '.x',
      '1ex',
      '1e+',
      '1e309',
      '-1',
    ]) {
      expect(captureTree(() => parsed(source)).diagnostic.message).toBe(
        'Worksheet page bottom margin is invalid',
      );
    }
  });

  it.each([
    [
      { autoPageBreaks: 'bad' },
      'Worksheet automatic page-break flag is invalid',
    ],
    [{ fitToPage: 'bad' }, 'Worksheet fit-to-page flag is invalid'],
  ] as const)('rejects invalid page setup properties %#', (attrs, message) => {
    expect(
      captureTree(() =>
        parseXlsxPageSetupProperties(element('pageSetUpPr', attrs), PART),
      ).diagnostic.message,
    ).toBe(message);
  });

  it.each([
    [
      { blackAndWhite: 'bad' },
      'Worksheet black-and-white print flag is invalid',
    ],
    [{ cellComments: 'bad' }, 'Worksheet printed-comment mode is invalid'],
    [{ copies: '01' }, 'Worksheet print copies are invalid'],
    [{ draft: 'bad' }, 'Worksheet draft print flag is invalid'],
    [{ errors: 'bad' }, 'Worksheet printed-error mode is invalid'],
    [{ firstPageNumber: '01' }, 'Worksheet first page number is invalid'],
    [{ fitToHeight: '01' }, 'Worksheet fit-to-height value is invalid'],
    [{ fitToWidth: '01' }, 'Worksheet fit-to-width value is invalid'],
    [{ horizontalDpi: '01' }, 'Worksheet horizontal DPI is invalid'],
    [{ orientation: 'bad' }, 'Worksheet page orientation is invalid'],
    [{ pageOrder: 'bad' }, 'Worksheet page order is invalid'],
    [{ paperHeight: 'bad' }, 'Worksheet paper height is invalid'],
    [{ paperSize: '01' }, 'Worksheet paper size is invalid'],
    [{ paperWidth: 'bad' }, 'Worksheet paper width is invalid'],
    [{ scale: '01' }, 'Worksheet page scale is invalid'],
    [
      { useFirstPageNumber: 'bad' },
      'Worksheet use-first-page-number flag is invalid',
    ],
    [
      { usePrinterDefaults: 'bad' },
      'Worksheet use-printer-defaults flag is invalid',
    ],
    [{ verticalDpi: '01' }, 'Worksheet vertical DPI is invalid'],
  ] as const)('rejects invalid normalized page setup %#', (attrs, message) => {
    expect(
      captureTree(() => parseXlsxPageSetup(element('pageSetup', attrs), PART))
        .diagnostic.message,
    ).toBe(message);
  });

  it('accepts exact page setup numeric boundaries', () => {
    expect(
      parseXlsxPageSetup(
        element('pageSetup', {
          firstPageNumber: '4294967295',
          scale: '10',
        }),
        PART,
      ),
    ).toMatchObject({ firstPageNumber: 0xffff_ffff, scale: 10 });
    expect(
      parseXlsxPageSetup(element('pageSetup', { scale: '400' }), PART).scale,
    ).toBe(400);
    expect(
      captureTree(() =>
        parseXlsxPageSetup(
          element('pageSetup', { firstPageNumber: '4294967296' }),
          PART,
        ),
      ).diagnostic.message,
    ).toBe('Worksheet first page number is invalid');
  });

  it('parses every supported universal unit and normalizes negative zero', () => {
    for (const unit of ['cm', 'in', 'mm', 'pc', 'pi', 'pt'] as const) {
      expect(
        parseXlsxPageSetup(
          element('pageSetup', { paperWidth: `1${unit}` }),
          PART,
        ).paperWidth,
      ).toEqual({ unit, value: 1 });
    }
    expect(
      parseXlsxPageMargins(
        element('pageMargins', {
          bottom: '-0',
          footer: '0',
          header: '0',
          left: '0',
          right: '0',
          top: '0',
        }),
        PART,
      ).bottom,
    ).toBe(0);
    for (const source of [
      'x1mm',
      '1mmx',
      '1.2.3mm',
      '.xmm',
      `${'9'.repeat(400)}mm`,
    ]) {
      expect(
        captureTree(() =>
          parseXlsxPageSetup(
            element('pageSetup', { paperWidth: source }),
            PART,
          ),
        ).diagnostic.message,
      ).toBe('Worksheet paper width is invalid');
    }
    for (const source of ['12.5mm', '1.mm', '.55mm']) {
      expect(
        parseXlsxPageSetup(element('pageSetup', { paperWidth: source }), PART)
          .paperWidth?.unit,
      ).toBe('mm');
    }
  });

  it('parses lexical booleans and ignores foreign attributes', () => {
    for (const [source, expected] of [
      ['0', false],
      ['false', false],
      ['1', true],
      ['true', true],
    ] as const) {
      expect(
        parseXlsxPrintOptions(
          element('printOptions', { headings: source }),
          PART,
        ).headings,
      ).toBe(expected);
    }
    const properties = element('pageSetUpPr', { fitToPage: '1' });
    properties.attributes = new Map([
      ['{}fitToPage', '1'],
      ['xxfitToPage', 'bad'],
    ]);
    expect(parseXlsxPageSetupProperties(properties, PART).fitToPage).toBe(true);
  });

  it('captures header/footer chunks with exact text accounting', () => {
    const state = budget();
    const capture = new XlsxHeaderFooterCapture(
      state,
      { ...defaultXlsxResourceLimits(), maxTextCharacters: 4 },
      PART,
    );
    capture.openElement(element('headerFooter'));
    capture.text(' \n');
    capture.openElement(element('oddHeader'));
    capture.text('ab');
    capture.text('cd');
    capture.closeElement(element('oddHeader'));
    capture.closeElement(element('headerFooter'));
    expect(capture.result()).toMatchObject({ oddHeader: 'abcd' });
    expect(state.textCharacters).toBe(4);

    const over = new XlsxHeaderFooterCapture(
      budget(),
      { ...defaultXlsxResourceLimits(), maxTextCharacters: 3 },
      PART,
    );
    over.openElement(element('headerFooter'));
    over.openElement(element('oddHeader'));
    expect(captureLimit(() => over.text('abcd'))).toMatchObject({
      actual: 4,
      limit: 3,
      limitName: 'maxTextCharacters',
    });
  });

  it('rejects malformed header/footer capture state', () => {
    const create = () =>
      new XlsxHeaderFooterCapture(budget(), defaultXlsxResourceLimits(), PART);
    expect(
      captureTree(() => create().openElement(element('wrong'))).diagnostic
        .message,
    ).toBe('Worksheet header/footer capture root is invalid');
    const incomplete = create();
    incomplete.openElement(element('headerFooter'));
    expect(captureTree(() => incomplete.result()).diagnostic.message).toBe(
      'Worksheet header/footer capture is incomplete',
    );
    const nesting = create();
    nesting.openElement(element('headerFooter'));
    nesting.openElement(element('oddHeader'));
    expect(
      captureTree(() => nesting.openElement(element('oddFooter'))).diagnostic
        .message,
    ).toBe('Worksheet header/footer nesting is invalid');

    const noRoot = create();
    expect(captureTree(() => noRoot.result()).diagnostic.message).toBe(
      'Worksheet header/footer capture is incomplete',
    );
    const openChild = create();
    openChild.openElement(element('headerFooter'));
    openChild.openElement(element('oddHeader'));
    expect(captureTree(() => openChild.result()).diagnostic.message).toBe(
      'Worksheet header/footer capture is incomplete',
    );
    expect(
      captureTree(() => openChild.closeElement(element('headerFooter')))
        .diagnostic.message,
    ).toBe('Worksheet header/footer nesting is invalid');

    const wrongClose = create();
    wrongClose.openElement(element('headerFooter'));
    wrongClose.openElement(element('oddHeader'));
    expect(
      captureTree(() => wrongClose.closeElement(element('oddFooter')))
        .diagnostic.message,
    ).toBe('Worksheet header/footer nesting is invalid');

    const closed = create();
    closed.openElement(element('headerFooter'));
    closed.closeElement(element('headerFooter'));
    expect(
      captureTree(() => closed.closeElement(element('headerFooter'))).diagnostic
        .message,
    ).toBe('Worksheet header/footer nesting is invalid');
    expect(
      captureTree(() => closed.openElement(element('oddHeader'))).diagnostic
        .message,
    ).toBe('Worksheet header/footer nesting is invalid');
  });

  it.each([
    [
      'alignWithMargins',
      'Worksheet header/footer margin-alignment flag is invalid',
    ],
    ['differentFirst', 'Worksheet first-header/footer flag is invalid'],
    ['differentOddEven', 'Worksheet odd/even-header/footer flag is invalid'],
    [
      'scaleWithDoc',
      'Worksheet header/footer scale-with-document flag is invalid',
    ],
  ] as const)('rejects invalid header/footer flag %s', (name, message) => {
    const capture = new XlsxHeaderFooterCapture(
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    capture.openElement(element('headerFooter', { [name]: 'bad' }));
    capture.closeElement(element('headerFooter'));
    expect(captureTree(() => capture.result()).diagnostic.message).toBe(
      message,
    );
  });

  it('captures page breaks and rejects malformed capture state', () => {
    const state = budget();
    const capture = new XlsxPageBreaksCapture(
      'row',
      state,
      { ...defaultXlsxResourceLimits(), maxColumnsPerWorksheet: 2 },
      PART,
    );
    capture.openElement(element('rowBreaks', { count: '1' }));
    capture.text(' \n');
    capture.openElement(element('brk', { id: '2', max: '1' }));
    capture.closeElement(element('brk'));
    capture.closeElement(element('rowBreaks'));
    expect(capture.result()).toStrictEqual([
      { end: 1, manual: false, pivot: false, position: 2, start: 0 },
    ]);
    expect(state.rangeAreas).toBe(1);

    const wrong = new XlsxPageBreaksCapture(
      'column',
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(
      captureTree(() => wrong.openElement(element('rowBreaks'))).diagnostic
        .message,
    ).toBe('Worksheet page-break capture root is invalid');
    const incomplete = new XlsxPageBreaksCapture(
      'row',
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    incomplete.openElement(element('rowBreaks'));
    expect(captureTree(() => incomplete.result()).diagnostic.message).toBe(
      'Worksheet page-break capture is incomplete',
    );
  });

  it('validates every page-break field and exact range boundary', () => {
    const parse = (
      axis: 'column' | 'row',
      root: 'colBreaks' | 'rowBreaks',
      rootAttrs: Record<string, string>,
      breakAttrs?: Record<string, string>,
      limits = defaultXlsxResourceLimits(),
    ) => {
      const capture = new XlsxPageBreaksCapture(axis, budget(), limits, PART);
      capture.openElement(element(root, rootAttrs));
      if (breakAttrs) {
        capture.openElement(element('brk', breakAttrs));
        capture.closeElement(element('brk'));
      }
      capture.closeElement(element(root));
      return capture.result();
    };
    expect(
      parse(
        'row',
        'rowBreaks',
        { count: '1', manualBreakCount: '0' },
        { id: '1', max: '0', min: '0' },
        { ...defaultXlsxResourceLimits(), maxRangeAreas: 1 },
      ),
    ).toStrictEqual([
      { end: 0, manual: false, pivot: false, position: 1, start: 0 },
    ]);
    expect(parse('column', 'colBreaks', {}, { id: '1' })).toStrictEqual([
      {
        end: 1_048_575,
        manual: false,
        pivot: false,
        position: 1,
        start: 0,
      },
    ]);
    for (const [rootAttrs, breakAttrs, message] of [
      [{ count: '01' }, undefined, 'Worksheet page-break count is invalid'],
      [
        { manualBreakCount: '01' },
        undefined,
        'Worksheet manual page-break count is invalid',
      ],
      [{}, { id: '01' }, 'Worksheet page-break position is invalid'],
      [
        {},
        { id: '1', max: '1', min: '01' },
        'Worksheet page-break start is invalid',
      ],
      [{}, { id: '1', max: '01' }, 'Worksheet page-break end is invalid'],
      [
        {},
        { id: '1', max: '0', pt: 'bad' },
        'Worksheet pivot page-break flag is invalid',
      ],
    ] as const) {
      expect(
        captureTree(() => parse('row', 'rowBreaks', rootAttrs, breakAttrs))
          .diagnostic.message,
      ).toBe(message);
    }
    expect(
      captureLimit(() =>
        parse(
          'column',
          'colBreaks',
          {},
          { id: '1', max: '2' },
          { ...defaultXlsxResourceLimits(), maxRowsPerWorksheet: 2 },
        ),
      ),
    ).toMatchObject({
      actual: 3,
      limit: 2,
      limitName: 'maxRowsPerWorksheet',
    });
  });

  it('rejects every malformed page-break capture transition', () => {
    const create = () =>
      new XlsxPageBreaksCapture(
        'row',
        budget(),
        defaultXlsxResourceLimits(),
        PART,
      );
    const wrongChild = create();
    wrongChild.openElement(element('rowBreaks'));
    expect(
      captureTree(() => wrongChild.openElement(element('wrong'))).diagnostic
        .message,
    ).toBe('Worksheet page-break nesting is invalid');

    const nested = create();
    nested.openElement(element('rowBreaks'));
    nested.openElement(element('brk', { id: '1' }));
    expect(
      captureTree(() => nested.openElement(element('brk', { id: '2' })))
        .diagnostic.message,
    ).toBe('Worksheet page-break nesting is invalid');
    expect(captureTree(() => nested.text('bad')).diagnostic.message).toBe(
      'Worksheet page-break text is invalid',
    );
    expect(
      captureTree(() => nested.closeElement(element('rowBreaks'))).diagnostic
        .message,
    ).toBe('Worksheet page-break nesting is invalid');

    const noChild = create();
    noChild.openElement(element('rowBreaks'));
    expect(
      captureTree(() => noChild.closeElement(element('brk'))).diagnostic
        .message,
    ).toBe('Worksheet page-break nesting is invalid');

    const wrongRoot = create();
    wrongRoot.openElement(element('rowBreaks'));
    expect(
      captureTree(() => wrongRoot.closeElement(element('colBreaks'))).diagnostic
        .message,
    ).toBe('Worksheet page-break nesting is invalid');

    const closed = create();
    closed.openElement(element('rowBreaks'));
    closed.closeElement(element('rowBreaks'));
    expect(
      captureTree(() => closed.openElement(element('brk', { id: '1' })))
        .diagnostic.message,
    ).toBe('Worksheet page-break nesting is invalid');
    expect(
      captureTree(() => closed.closeElement(element('rowBreaks'))).diagnostic
        .message,
    ).toBe('Worksheet page-break nesting is invalid');
  });
});
