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
import type { XlsxResolvedSheetSelection } from '../../src/formats/xlsx/internal/selection';
import type { XlsxXmlElement } from '../../src/formats/xlsx/internal/streaming-xml';
import {
  parseXlsxProtectedRanges,
  parseXlsxWorksheetProtection,
  XlsxProtectedRangesCapture,
} from '../../src/formats/xlsx/internal/worksheet-protection';
import type { XlsxWorksheetBudget } from '../../src/formats/xlsx/internal/worksheet';
import { parseXlsxWorkbookProtection } from '../../src/formats/xlsx/internal/workbook-protection';
import {
  createIndependentXlsx,
  independentWorkbook,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const WORKBOOK_PART = 'xl/workbook.xml';
const WORKSHEET_PART = 'xl/worksheets/sheet1.xml';
const FULL: XlsxResolvedSheetSelection = { kind: 'full-sheet' };

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

function worksheet(protection: string): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/>${protection}</worksheet>`;
}

function workbook(protection: string): string {
  return independentWorkbook(
    '<sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/>',
  ).replace('<sheets>', `${protection}<sheets>`);
}

async function parseProtection(
  sheetXml: string,
  workbookProtection = '',
  options: Parameters<typeof parseXlsx>[1] = {},
) {
  return parseXlsx(
    await createIndependentXlsx({
      'xl/workbook.xml': workbook(workbookProtection),
      'xl/worksheets/sheet1.xml': worksheet(sheetXml),
    }),
    options,
  );
}

async function capture(
  sheetXml: string,
  workbookProtection = '',
  options: Parameters<typeof parseXlsx>[1] = {},
): Promise<XlsxParseError> {
  try {
    await parseProtection(sheetXml, workbookProtection, options);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected XLSX protection parsing to fail');
}

function captureTree(action: () => unknown): XlsxParseError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected normalized protection parsing to fail');
}

function captureLimit(action: () => unknown): XlsxResourceLimitError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxResourceLimitError);
    return error as XlsxResourceLimitError;
  }
  throw new Error('Expected normalized protection resource failure');
}

const STRONG =
  'algorithmName="SHA-512" hashValue="AQID" saltValue="BAUG" spinCount="1000"';
const COMPLETE_SHEET_PROTECTION = `<sheetProtection sheet="1" objects="1" scenarios="0" formatCells="1" formatColumns="0" formatRows="1" insertColumns="0" insertRows="1" insertHyperlinks="0" deleteColumns="1" deleteRows="0" selectLockedCells="1" selectUnlockedCells="0" sort="1" autoFilter="0" pivotTables="1" password="ab12" ${STRONG}/>
<protectedRanges><protectedRange name="Input" sqref="A1:A2 C1" password="cdef" securityDescriptor="S-1-5-21"/></protectedRanges>`;

const COMPLETE_WORKBOOK_PROTECTION = `<workbookProtection lockStructure="1" lockWindows="0" lockRevision="1" workbookPassword="1234" workbookAlgorithmName="SHA-512" workbookHashValue="AQID" workbookSaltValue="BAUG" workbookSpinCount="1000" revisionsPassword="abcd" revisionsAlgorithmName="SHA-256" revisionsHashValue="BwgJ" revisionsSaltValue="CgsM" revisionsSpinCount="2000"/>`;

describe('XLSX workbook and worksheet protection', () => {
  it('preserves locks and complete hash metadata without verifying passwords', async () => {
    const document = await parseProtection(
      COMPLETE_SHEET_PROTECTION,
      COMPLETE_WORKBOOK_PROTECTION,
    );
    expect(document.workbook.protection).toEqual({
      lockRevisions: true,
      lockStructure: true,
      lockWindows: false,
      revisionsCredential: {
        legacyHash: 'ABCD',
        strongHash: {
          algorithmName: 'SHA-256',
          hashValue: 'BwgJ',
          saltValue: 'CgsM',
          spinCount: 2000,
        },
      },
      workbookCredential: {
        legacyHash: '1234',
        strongHash: {
          algorithmName: 'SHA-512',
          hashValue: 'AQID',
          saltValue: 'BAUG',
          spinCount: 1000,
        },
      },
    });
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    expect(sheet.kind === 'worksheet' ? sheet.protection : undefined).toEqual({
      credential: {
        legacyHash: 'AB12',
        strongHash: {
          algorithmName: 'SHA-512',
          hashValue: 'AQID',
          saltValue: 'BAUG',
          spinCount: 1000,
        },
      },
      protectAutoFilter: false,
      protectDeleteColumns: true,
      protectDeleteRows: false,
      protectFormatCells: true,
      protectFormatColumns: false,
      protectFormatRows: true,
      protectInsertColumns: false,
      protectInsertHyperlinks: false,
      protectInsertRows: true,
      protectObjects: true,
      protectPivotTables: true,
      protectScenarios: false,
      protectSelectLockedCells: true,
      protectSelectUnlockedCells: false,
      protectSheet: true,
      protectSort: true,
    });
    expect(sheet.kind === 'worksheet' ? sheet.protectedRanges : []).toEqual([
      {
        credential: { legacyHash: 'CDEF' },
        name: 'Input',
        ranges: [
          {
            end: { column: 1, row: 2 },
            reference: 'A1:A2',
            start: { column: 1, row: 1 },
          },
          {
            end: { column: 3, row: 1 },
            reference: 'C1',
            start: { column: 3, row: 1 },
          },
        ],
        securityDescriptor: 'S-1-5-21',
        selectionRelation: 'full-sheet',
      },
    ]);
    expect(JSON.parse(JSON.stringify(document))).toEqual(document);
  });

  it.each([
    ['0', false],
    ['false', false],
    ['1', true],
    ['true', true],
  ] as const)(
    'parses every workbook and worksheet boolean %s',
    async (source, expected) => {
      const document = await parseProtection(
        `<sheetProtection sheet="${source}" objects="${source}" scenarios="${source}" formatCells="${source}" formatColumns="${source}" formatRows="${source}" insertColumns="${source}" insertRows="${source}" insertHyperlinks="${source}" deleteColumns="${source}" deleteRows="${source}" selectLockedCells="${source}" selectUnlockedCells="${source}" sort="${source}" autoFilter="${source}" pivotTables="${source}"/>`,
        `<workbookProtection lockStructure="${source}" lockWindows="${source}" lockRevision="${source}"/>`,
      );
      expect(document.workbook.protection).toMatchObject({
        lockRevisions: expected,
        lockStructure: expected,
        lockWindows: expected,
      });
      const sheet = document.sheets[0]!;
      expect(
        sheet.kind === 'worksheet' ? sheet.protection : undefined,
      ).toMatchObject(
        Object.fromEntries(
          [
            'protectAutoFilter',
            'protectDeleteColumns',
            'protectDeleteRows',
            'protectFormatCells',
            'protectFormatColumns',
            'protectFormatRows',
            'protectInsertColumns',
            'protectInsertHyperlinks',
            'protectInsertRows',
            'protectObjects',
            'protectPivotTables',
            'protectScenarios',
            'protectSelectLockedCells',
            'protectSelectUnlockedCells',
            'protectSheet',
            'protectSort',
          ].map((name) => [name, expected]),
        ),
      );
    },
  );

  it.each([
    ['autoFilter', 'Worksheet protection auto-filter flag is invalid'],
    ['deleteColumns', 'Worksheet protection delete-columns flag is invalid'],
    ['deleteRows', 'Worksheet protection delete-rows flag is invalid'],
    ['formatCells', 'Worksheet protection format-cells flag is invalid'],
    ['formatColumns', 'Worksheet protection format-columns flag is invalid'],
    ['formatRows', 'Worksheet protection format-rows flag is invalid'],
    ['insertColumns', 'Worksheet protection insert-columns flag is invalid'],
    [
      'insertHyperlinks',
      'Worksheet protection insert-hyperlinks flag is invalid',
    ],
    ['insertRows', 'Worksheet protection insert-rows flag is invalid'],
    ['objects', 'Worksheet protection objects flag is invalid'],
    ['pivotTables', 'Worksheet protection pivot-tables flag is invalid'],
    ['scenarios', 'Worksheet protection scenarios flag is invalid'],
    [
      'selectLockedCells',
      'Worksheet protection select-locked-cells flag is invalid',
    ],
    [
      'selectUnlockedCells',
      'Worksheet protection select-unlocked-cells flag is invalid',
    ],
    ['sheet', 'Worksheet protection sheet flag is invalid'],
    ['sort', 'Worksheet protection sort flag is invalid'],
  ] as const)(
    'rejects invalid worksheet protection flag %s',
    async (name, message) => {
      expect(
        (await capture(`<sheetProtection ${name}="bad"/>`)).diagnostic.message,
      ).toBe(message);
    },
  );

  it.each([
    ['lockRevision', 'Workbook revision-lock flag is invalid'],
    ['lockStructure', 'Workbook structure-lock flag is invalid'],
    ['lockWindows', 'Workbook window-lock flag is invalid'],
  ] as const)(
    'rejects invalid workbook protection flag %s',
    async (name, message) => {
      expect(
        (await capture('', `<workbookProtection ${name}="bad"/>`)).diagnostic
          .message,
      ).toBe(message);
    },
  );

  it('filters protected ranges by selection while preserving sheet locks', async () => {
    const included = await parseProtection(COMPLETE_SHEET_PROTECTION, '', {
      selection: { ranges: { Sheet1: ['C1'] } },
    });
    const includedSheet = included.sheets[0]!;
    expect(
      includedSheet.kind === 'worksheet' ? includedSheet.protectedRanges : [],
    ).toMatchObject([{ selectionRelation: 'intersects-selection' }]);
    const excluded = await parseProtection(COMPLETE_SHEET_PROTECTION, '', {
      selection: { ranges: { Sheet1: ['Z1'] } },
    });
    const excludedSheet = excluded.sheets[0]!;
    expect(
      excludedSheet.kind === 'worksheet' ? excludedSheet.protectedRanges : [],
    ).toEqual([]);
    expect(
      excludedSheet.kind === 'worksheet' ? excludedSheet.protection : undefined,
    ).toBeDefined();
  });

  it('round-trips protection metadata through portable exact R0', async () => {
    const bytes = await createIndependentXlsx({
      'xl/workbook.xml': workbook(COMPLETE_WORKBOOK_PROTECTION),
      'xl/worksheets/sheet1.xml': worksheet(COMPLETE_SHEET_PROTECTION),
    });
    const snapshot = await readXlsxRoundTrip(bytes);
    const result = await writeXlsxRoundTrip(
      await validateXlsxRoundTripJson(
        JSON.parse(JSON.stringify(snapshot)) as unknown,
      ),
    );
    expect(result.data).toEqual(bytes);
    expect(result.report.level).toBe('R0');
  });

  it('parses prefixed Strict protection elements', async () => {
    const strictSheet = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictRelationship =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const bytes = await createIndependentXlsx({
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelationship}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${strictRelationship}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheet}" xmlns:r="${strictRelationship}"><s:workbookProtection lockStructure="1"/><s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets></s:workbook>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheet}"><s:sheetData/><s:sheetProtection sheet="1"/><s:protectedRanges><s:protectedRange name="Input" sqref="A1"/></s:protectedRanges></s:worksheet>`,
    });
    const document = await parseXlsx(bytes);
    const sheet = document.sheets[0]!;
    expect(document.workbook.protection?.lockStructure).toBe(true);
    expect(
      sheet.kind === 'worksheet' ? sheet.protection?.protectSheet : false,
    ).toBe(true);
    expect(
      sheet.kind === 'worksheet' ? sheet.protectedRanges : [],
    ).toHaveLength(1);
  });

  it.each([
    [
      '<sheetProtection sheet="bad"/>',
      '',
      'Worksheet protection sheet flag is invalid',
    ],
    [
      '<sheetProtection password="123"/>',
      '',
      'Worksheet protection legacy password hash is invalid',
    ],
    [
      '<sheetProtection password="xAB12"/>',
      '',
      'Worksheet protection legacy password hash is invalid',
    ],
    [
      '<sheetProtection password="AB12x"/>',
      '',
      'Worksheet protection legacy password hash is invalid',
    ],
    [
      '<sheetProtection algorithmName="SHA-512"/>',
      '',
      'Worksheet protection strong hash metadata is incomplete',
    ],
    [
      '<sheetProtection algorithmName="bad name" hashValue="AQID" saltValue="BAUG" spinCount="1"/>',
      '',
      'Worksheet protection hash algorithm is invalid',
    ],
    [
      '<sheetProtection algorithmName="SHA-512" hashValue="bad" saltValue="BAUG" spinCount="1"/>',
      '',
      'Worksheet protection hash value is invalid',
    ],
    [
      '<sheetProtection algorithmName="SHA-512" hashValue="AB==" saltValue="BAUG" spinCount="1"/>',
      '',
      'Worksheet protection hash value is invalid',
    ],
    [
      '<sheetProtection algorithmName="SHA-512" hashValue="AQID" saltValue="" spinCount="1"/>',
      '',
      'Worksheet protection salt value is invalid',
    ],
    [
      '<sheetProtection algorithmName="SHA-512" hashValue="AQID" saltValue="BAUG" spinCount="0"/>',
      '',
      'Worksheet protection spin count is invalid',
    ],
    [
      '<sheetProtection algorithmName="SHA-512" hashValue="AQID" saltValue="BAUG" spinCount="x1"/>',
      '',
      'Worksheet protection spin count is invalid',
    ],
    [
      '<sheetProtection algorithmName="SHA-512" hashValue="AQID" saltValue="BAUG" spinCount="1x"/>',
      '',
      'Worksheet protection spin count is invalid',
    ],
    [
      '<protectedRanges/>',
      '',
      'Protected-range collection is empty or invalid',
    ],
    [
      '<protectedRanges><protectedRange sqref="A1"/></protectedRanges>',
      '',
      'Protected-range name is invalid',
    ],
    [
      '<protectedRanges><protectedRange name="Input" sqref="bad"/></protectedRanges>',
      '',
      'Protected-range reference is invalid',
    ],
    [
      '<protectedRanges><protectedRange name="Input" sqref="A1 A1"/></protectedRanges>',
      '',
      'Protected-range reference list contains duplicates',
    ],
    [
      '<protectedRanges>bad<protectedRange name="Input" sqref="A1"/></protectedRanges>',
      '',
      'Protected-ranges text content is invalid',
    ],
    [
      '<sheetProtection/><sheetProtection/>',
      '',
      'Worksheet contains duplicate sheetProtection elements',
    ],
    [
      '<protectedRanges><protectedRange name="One" sqref="A1"/></protectedRanges><protectedRanges><protectedRange name="Two" sqref="B1"/></protectedRanges>',
      '',
      'Worksheet contains duplicate protectedRanges elements',
    ],
    [
      '',
      '<workbookProtection lockStructure="bad"/>',
      'Workbook structure-lock flag is invalid',
    ],
    [
      '',
      '<workbookProtection workbookPassword="123"/>',
      'Workbook protection legacy password hash is invalid',
    ],
    [
      '',
      '<workbookProtection workbookAlgorithmName="SHA-512"/>',
      'Workbook protection strong hash metadata is incomplete',
    ],
    [
      '',
      '<workbookProtection revisionsAlgorithmName="SHA-512" revisionsHashValue="AQID" revisionsSaltValue="BAUG" revisionsSpinCount="4294967296"/>',
      'Revision protection spin count is invalid',
    ],
  ] as const)(
    'rejects invalid protection contract %#',
    async (sheetXml, bookXml, message) => {
      expect((await capture(sheetXml, bookXml)).diagnostic.message).toBe(
        message,
      );
    },
  );

  it('charges workbook protection text to the public aggregate limit', async () => {
    const error = await capture(
      '',
      '<workbookProtection workbookPassword="1234"/>',
      { limits: { maxTextCharacters: 12 } },
    );
    expect(error.diagnostic).toMatchObject({
      actual: 13,
      code: 'resource-limit-exceeded',
      limit: 12,
      limitName: 'maxTextCharacters',
      part: WORKBOOK_PART,
    });
  });
});

describe('XLSX normalized protection helpers and capture', () => {
  it('enforces protected-range text, range, grid, and selection-work limits', () => {
    const value = {
      attrs: {},
      protectedRange: {
        attrs: {
          name: 'Input',
          securityDescriptor: 'abc',
          sqref: 'A1 B2',
        },
      },
    };
    expect(
      parseXlsxProtectedRanges(
        value,
        FULL,
        budget(),
        {
          ...defaultXlsxResourceLimits(),
          maxRangeAreas: 2,
          maxTextCharacters: 8,
        },
        WORKSHEET_PART,
      ),
    ).toHaveLength(1);
    expect(
      captureLimit(() =>
        parseXlsxProtectedRanges(
          value,
          FULL,
          budget(),
          { ...defaultXlsxResourceLimits(), maxTextCharacters: 7 },
          WORKSHEET_PART,
        ),
      ),
    ).toMatchObject({ actual: 8, limit: 7, limitName: 'maxTextCharacters' });
    expect(
      captureLimit(() =>
        parseXlsxProtectedRanges(
          value,
          FULL,
          budget(),
          { ...defaultXlsxResourceLimits(), maxRangeAreas: 1 },
          WORKSHEET_PART,
        ),
      ),
    ).toMatchObject({ actual: 2, limit: 1, limitName: 'maxRangeAreas' });
    const selection: XlsxResolvedSheetSelection = {
      endRowPrefix: [3],
      kind: 'selected-ranges',
      ranges: [
        {
          end: { column: 3, row: 3 },
          reference: 'C3',
          start: { column: 3, row: 3 },
        },
      ],
    };
    expect(
      captureLimit(() =>
        parseXlsxProtectedRanges(
          value,
          selection,
          budget(),
          { ...defaultXlsxResourceLimits(), maxScannedCells: 1 },
          WORKSHEET_PART,
        ),
      ),
    ).toMatchObject({ actual: 2, limit: 1, limitName: 'maxScannedCells' });
    for (const [sqref, limits, limitName] of [
      ['A2', { maxRowsPerWorksheet: 1 }, 'maxRowsPerWorksheet'],
      ['B1', { maxColumnsPerWorksheet: 1 }, 'maxColumnsPerWorksheet'],
    ] as const) {
      expect(
        captureLimit(() =>
          parseXlsxProtectedRanges(
            {
              attrs: {},
              protectedRange: { attrs: { name: 'Input', sqref } },
            },
            FULL,
            budget(),
            { ...defaultXlsxResourceLimits(), ...limits },
            WORKSHEET_PART,
          ),
        ),
      ).toMatchObject({ actual: 2, limit: 1, limitName });
    }
    expect(
      parseXlsxProtectedRanges(
        {
          attrs: {},
          protectedRange: {
            attrs: { name: 'Boundary', sqref: 'B2' },
          },
        },
        FULL,
        budget(),
        {
          ...defaultXlsxResourceLimits(),
          maxColumnsPerWorksheet: 2,
          maxRowsPerWorksheet: 2,
        },
        WORKSHEET_PART,
      ),
    ).toHaveLength(1);
  });

  it('distinguishes every protected-range intersection boundary', () => {
    const selected = (reference: string): XlsxResolvedSheetSelection => {
      const [column, row] = [
        reference.charCodeAt(0) - 64,
        Number(reference[1]),
      ];
      return {
        endRowPrefix: [row],
        kind: 'selected-ranges',
        ranges: [
          {
            end: { column, row },
            reference,
            start: { column, row },
          },
        ],
      };
    };
    const parsed = (sqref: string, selection: XlsxResolvedSheetSelection) =>
      parseXlsxProtectedRanges(
        {
          attrs: {},
          protectedRange: { attrs: { name: 'Input', sqref } },
        },
        selection,
        budget(),
        defaultXlsxResourceLimits(),
        WORKSHEET_PART,
      );

    expect(parsed('A3', selected('A1'))).toEqual([]);
    expect(parsed('A1', selected('A3'))).toEqual([]);
    expect(parsed('C1', selected('A1'))).toEqual([]);
    expect(parsed('A1', selected('C1'))).toEqual([]);
    expect(parsed('A1:B2', selected('B2'))).toMatchObject([
      { selectionRelation: 'intersects-selection' },
    ]);
    expect(parsed('A1', { kind: 'not-selected' })).toEqual([]);
  });

  it('returns explicit defaults and exact text accounting', () => {
    expect(
      parseXlsxWorksheetProtection(
        element('sheetProtection'),
        budget(),
        defaultXlsxResourceLimits(),
        WORKSHEET_PART,
      ),
    ).toStrictEqual({
      protectAutoFilter: false,
      protectDeleteColumns: false,
      protectDeleteRows: false,
      protectFormatCells: false,
      protectFormatColumns: false,
      protectFormatRows: false,
      protectInsertColumns: false,
      protectInsertHyperlinks: false,
      protectInsertRows: false,
      protectObjects: false,
      protectPivotTables: false,
      protectScenarios: false,
      protectSelectLockedCells: false,
      protectSelectUnlockedCells: false,
      protectSheet: false,
      protectSort: false,
    });

    const namespaced = element('sheetProtection');
    namespaced.attributes = new Map([
      ['{}sheet', '1'],
      ['xxsheet', 'bad'],
    ]);
    expect(
      parseXlsxWorksheetProtection(
        namespaced,
        budget(),
        defaultXlsxResourceLimits(),
        WORKSHEET_PART,
      ).protectSheet,
    ).toBe(true);
    expect(parseXlsxWorkbookProtection(undefined, WORKBOOK_PART)).toEqual({
      textCharacters: 0,
    });
    expect(
      parseXlsxWorkbookProtection({ attrs: {} }, WORKBOOK_PART),
    ).toStrictEqual({
      protection: {
        lockRevisions: false,
        lockStructure: false,
        lockWindows: false,
      },
      textCharacters: 0,
    });
    expect(
      captureTree(() => parseXlsxWorkbookProtection([], WORKBOOK_PART))
        .diagnostic.message,
    ).toBe('Workbook protection is invalid');
    expect(
      parseXlsxWorkbookProtection(
        {
          attrs: {
            workbookAlgorithmName: 'SHA-512',
            workbookHashValue: 'AQID',
            workbookSaltValue: 'BAUG',
            workbookSpinCount: '4294967295',
          },
        },
        WORKBOOK_PART,
      ),
    ).toMatchObject({
      protection: {
        lockRevisions: false,
        lockStructure: false,
        lockWindows: false,
        workbookCredential: {
          strongHash: { spinCount: 0xffff_ffff },
        },
      },
      textCharacters: 'SHA-512AQIDBAUG'.length,
    });
    expect(
      parseXlsxWorkbookProtection(
        {
          attrs: {
            revisionsAlgorithmName: 'SHA-256',
            revisionsHashValue: 'BwgJ',
            revisionsPassword: 'ABCD',
            revisionsSaltValue: 'CgsM',
            revisionsSpinCount: '2',
            workbookAlgorithmName: 'SHA-512',
            workbookHashValue: 'AQID',
            workbookPassword: '1234',
            workbookSaltValue: 'BAUG',
            workbookSpinCount: '1',
          },
        },
        WORKBOOK_PART,
      ).textCharacters,
    ).toBe('SHA-256BwgJABCDCgsMSHA-512AQID1234BAUG'.length);
    for (const workbookSpinCount of [' 1', '1 ']) {
      expect(
        captureTree(() =>
          parseXlsxWorkbookProtection(
            {
              attrs: {
                workbookAlgorithmName: 'SHA-512',
                workbookHashValue: 'AQID',
                workbookSaltValue: 'BAUG',
                workbookSpinCount,
              },
            },
            WORKBOOK_PART,
          ),
        ).diagnostic.message,
      ).toBe('Workbook protection spin count is invalid');
    }
  });

  it('validates normalized protected-range shapes and reference lists', () => {
    for (const [value, message] of [
      [undefined, 'Protected-ranges collection is invalid'],
      ['bad', 'Protected-ranges collection is invalid'],
      [
        { attrs: {}, protectedRange: 'bad' },
        'Protected-range collection is empty or invalid',
      ],
      [
        { attrs: {}, protectedRange: [] },
        'Protected-range collection is empty or invalid',
      ],
      [
        { attrs: {}, protectedRange: { attrs: { name: '', sqref: 'A1' } } },
        'Protected-range name is invalid',
      ],
      [
        { attrs: {}, protectedRange: { attrs: { name: 'Input', sqref: 1 } } },
        'Protected-range reference list is invalid',
      ],
      [
        {
          attrs: {},
          protectedRange: { attrs: { name: 'Input', sqref: '   ' } },
        },
        'Protected-range reference list is invalid',
      ],
      [
        {
          attrs: {},
          protectedRange: {
            attrs: { name: 'Input', password: 'bad', sqref: 'A1' },
          },
        },
        'Protected range legacy password hash is invalid',
      ],
    ] as const) {
      expect(
        captureTree(() =>
          parseXlsxProtectedRanges(
            value,
            FULL,
            budget(),
            defaultXlsxResourceLimits(),
            WORKSHEET_PART,
          ),
        ).diagnostic.message,
      ).toBe(message);
    }

    expect(
      parseXlsxProtectedRanges(
        {
          attrs: {},
          protectedRange: {
            attrs: { name: 'Input', sqref: '\tA1  B2\r\n' },
          },
        },
        FULL,
        budget(),
        defaultXlsxResourceLimits(),
        WORKSHEET_PART,
      ),
    ).toHaveLength(1);
  });

  it('preserves protected-range strong credentials with exact accounting', () => {
    const value = {
      attrs: {},
      protectedRange: {
        attrs: {
          algorithmName: 'SHA-512',
          hashValue: 'AQID',
          name: 'Input',
          saltValue: 'BAUG',
          securityDescriptor: 'acl',
          spinCount: '10',
          sqref: 'A1',
        },
      },
    };
    const expectedText = 'InputaclSHA-512AQIDBAUG'.length;
    expect(
      parseXlsxProtectedRanges(
        value,
        FULL,
        budget(),
        {
          ...defaultXlsxResourceLimits(),
          maxTextCharacters: expectedText,
        },
        WORKSHEET_PART,
      ),
    ).toStrictEqual([
      {
        credential: {
          strongHash: {
            algorithmName: 'SHA-512',
            hashValue: 'AQID',
            saltValue: 'BAUG',
            spinCount: 10,
          },
        },
        name: 'Input',
        ranges: [
          {
            end: { column: 1, row: 1 },
            reference: 'A1',
            start: { column: 1, row: 1 },
          },
        ],
        securityDescriptor: 'acl',
        selectionRelation: 'full-sheet',
      },
    ]);
    expect(
      captureLimit(() =>
        parseXlsxProtectedRanges(
          value,
          FULL,
          budget(),
          {
            ...defaultXlsxResourceLimits(),
            maxTextCharacters: expectedText - 1,
          },
          WORKSHEET_PART,
        ),
      ),
    ).toMatchObject({
      actual: expectedText,
      limit: expectedText - 1,
      limitName: 'maxTextCharacters',
    });
  });

  it('captures protected ranges and rejects malformed capture', () => {
    const capture = new XlsxProtectedRangesCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      WORKSHEET_PART,
    );
    capture.openElement(element('protectedRanges'));
    capture.openElement(
      element('protectedRange', { name: 'Input', sqref: 'A1' }),
    );
    capture.closeElement(element('protectedRange'));
    capture.closeElement(element('protectedRanges'));
    expect(capture.result()).toStrictEqual([
      {
        name: 'Input',
        ranges: [
          {
            end: { column: 1, row: 1 },
            reference: 'A1',
            start: { column: 1, row: 1 },
          },
        ],
        selectionRelation: 'full-sheet',
      },
    ]);
    const invalid = new XlsxProtectedRangesCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      WORKSHEET_PART,
    );
    expect(
      captureTree(() => invalid.openElement(element('wrong'))).diagnostic
        .message,
    ).toBe('Protected-ranges capture root is invalid');
    const incomplete = new XlsxProtectedRangesCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      WORKSHEET_PART,
    );
    incomplete.openElement(element('protectedRanges'));
    expect(captureTree(() => incomplete.result()).diagnostic.message).toBe(
      'Protected-ranges capture is incomplete',
    );
    const nesting = new XlsxProtectedRangesCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      WORKSHEET_PART,
    );
    nesting.openElement(element('protectedRanges'));
    expect(
      captureTree(() => nesting.closeElement(element('wrong'))).diagnostic
        .message,
    ).toBe('Protected-ranges capture nesting is invalid');
    const text = new XlsxProtectedRangesCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      WORKSHEET_PART,
    );
    text.openElement(element('protectedRanges'));
    text.text(' \n\t');
    expect(captureTree(() => text.text('bad')).diagnostic.message).toBe(
      'Protected-ranges text content is invalid',
    );
    const duplicate = new XlsxProtectedRangesCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      WORKSHEET_PART,
    );
    duplicate.openElement(element('protectedRanges'));
    duplicate.closeElement(element('protectedRanges'));
    expect(
      captureTree(() => duplicate.openElement(element('protectedRanges')))
        .diagnostic.message,
    ).toBe('Protected-ranges capture root is invalid');
  });

  it('captures repeated protected-range siblings', () => {
    const capture = new XlsxProtectedRangesCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      WORKSHEET_PART,
    );
    capture.openElement(element('protectedRanges'));
    for (const [name, sqref] of [
      ['One', 'A1'],
      ['Two', 'B2'],
    ] as const) {
      capture.openElement(element('protectedRange', { name, sqref }));
      capture.closeElement(element('protectedRange'));
    }
    capture.closeElement(element('protectedRanges'));
    expect(capture.result().map((range) => range.name)).toEqual(['One', 'Two']);
  });

  it('preserves raw expanded protected-range attributes', () => {
    const capture = new XlsxProtectedRangesCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      WORKSHEET_PART,
    );
    capture.openElement(element('protectedRanges'));
    const range = element('protectedRange');
    range.attributes = new Map([
      ['{}name', 'Input'],
      ['sqref', 'A1'],
    ]);
    capture.openElement(range);
    capture.closeElement(element('protectedRange'));
    capture.closeElement(element('protectedRanges'));
    expect(capture.result()).toMatchObject([{ name: 'Input' }]);
  });
});
