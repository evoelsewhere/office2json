import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  parseXlsxWithDiagnostics,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
  XlsxParseError,
} from '../../src/formats/xlsx';
import {
  parseXlsxPivotDataDisplayMode,
  parseXlsxPivotDateTime,
  parseXlsxPivotFiniteNumber,
  parseXlsxPivotFilterType,
  parseXlsxPivotFieldIndexes,
  parseXlsxPivotRecordCount,
  parseXlsxPivotSignedInteger,
  parseXlsxPivotSubtotal,
  parseXlsxPivotUnsignedInteger,
  XlsxPivotCacheRecordsSink,
} from '../../src/formats/xlsx/internal/pivot';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import type { XlsxXmlElement } from '../../src/formats/xlsx/internal/streaming-xml';
import { createXlsxWorksheetBudget } from '../../src/formats/xlsx/internal/worksheet';
import {
  createIndependentXlsx,
  type XlsxBlackBoxOverrides,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const CACHE_DEFINITION_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}pivotCacheDefinition`;
const CACHE_RECORDS_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}pivotCacheRecords`;
const PIVOT_TABLE_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}pivotTable`;

function pivotElement(
  localName: string,
  namespace = XLSX_SPREADSHEET_NS,
  attributes: ReadonlyMap<string, string> = new Map(),
): XlsxXmlElement {
  return { attributes, localName, namespace };
}

function captureSync(action: () => unknown): XlsxParseError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected pivot record capture to fail');
}

function recordSink(
  fields: ConstructorParameters<typeof XlsxPivotCacheRecordsSink>[1] = [],
): XlsxPivotCacheRecordsSink {
  return new XlsxPivotCacheRecordsSink(
    XLSX_SPREADSHEET_NS,
    fields,
    { records: 0 },
    createXlsxWorksheetBudget({ part: null, values: [] }),
    defaultXlsxResourceLimits(),
    'xl/pivotCache/pivotCacheRecords1.xml',
  );
}

const CONTENT_TYPES = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
</Types>`;

const WORKBOOK = `<workbook xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="sheet"/></sheets><pivotCaches><pivotCache cacheId="7" r:id="cache"/></pivotCaches></workbook>`;
const WORKBOOK_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="cache" Type="${CACHE_DEFINITION_RELATIONSHIP}" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>`;
const CACHE_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="records" Type="${CACHE_RECORDS_RELATIONSHIP}" Target="pivotCacheRecords1.xml"/></Relationships>`;
const WORKSHEET = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row></sheetData><pivotTableParts count="1"><pivotTablePart r:id="pivot"/></pivotTableParts></worksheet>`;
const WORKSHEET_WITHOUT_PIVOT = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/></worksheet>`;
const WORKSHEET_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="pivot" Type="${PIVOT_TABLE_RELATIONSHIP}" Target="../pivotTables/pivotTable1.xml"/></Relationships>`;

function cacheDefinition(overrides = ''): string {
  return `<pivotCacheDefinition xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}" recordCount="2" refreshedBy="Producer" refreshedDate="45292.5" refreshOnLoad="1" saveData="true" enableRefresh="0" backgroundQuery="1" upgradeOnRefresh="true" tupleCache="1" supportAdvancedDrill="false" missingItemsLimit="10" ${overrides}><cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="Sheet1"/></cacheSource><cacheFields count="2"><cacheField name="Category" databaseField="1" serverField="0" uniqueList="true"><sharedItems count="2" containsString="1" containsBlank="0"><s v="A"/><s v="B"/></sharedItems></cacheField><cacheField name="Amount"><sharedItems count="2" containsNumber="true" containsInteger="1" minValue="1" maxValue="2"><n v="1"/><n v="2"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>`;
}

function cacheRecords(): string {
  return `<pivotCacheRecords xmlns="${XLSX_SPREADSHEET_NS}" count="2"><r><x v="0"/><n v="1"/></r><r><x v="1"/><n v="2"/></r></pivotCacheRecords>`;
}

function cacheDefinitionWithSource(source: string): string {
  return cacheDefinition().replace(
    '<cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="Sheet1"/></cacheSource>',
    source,
  );
}

function pivotTable(overrides = ''): string {
  return `<pivotTableDefinition xmlns="${XLSX_SPREADSHEET_NS}" name="SalesPivot" cacheId="7" dataCaption="Values" grandTotalCaption="Grand" compact="0" outline="1" showHeaders="true" rowGrandTotals="0" colGrandTotals="1" ${overrides}><location ref="D1:H10" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><pivotFields count="2"><pivotField axis="axisRow" name="Category" showAll="0" compact="1" outline="0" subtotalTop="false" sortType="ascending"><items count="2"><item x="0"/><item x="1" t="blank"/></items></pivotField><pivotField axis="axisValues" dataField="1"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><colFields count="1"><field x="-2"/></colFields><pageFields count="1"><pageField fld="0" item="1" hier="0" name="Page"/></pageFields><dataFields count="1"><dataField name="Sum Amount" fld="1" subtotal="sum" showDataAs="percentOfTotal" baseField="0" baseItem="0"/></dataFields><filters count="1"><filter fld="0" type="captionEqual" id="1" evalOrder="-1" name="Filter" description="Only A" stringValue1="A"/></filters><pivotTableStyleInfo name="PivotStyleMedium9" showRowHeaders="1" showColHeaders="true" showRowStripes="1" showColStripes="0" showLastColumn="false"/></pivotTableDefinition>`;
}

function parts(changes: XlsxBlackBoxOverrides = {}): XlsxBlackBoxOverrides {
  return {
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': CACHE_RELS,
    'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition(),
    'xl/pivotCache/pivotCacheRecords1.xml': cacheRecords(),
    'xl/pivotTables/pivotTable1.xml': pivotTable(),
    'xl/sharedStrings.xml': null,
    'xl/styles.xml': null,
    'xl/workbook.xml': WORKBOOK,
    'xl/worksheets/sheet1.xml': WORKSHEET,
    'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS,
    ...changes,
  };
}

async function bytes(changes: XlsxBlackBoxOverrides = {}): Promise<Uint8Array> {
  return createIndependentXlsx(parts(changes));
}

async function capture(
  changes: XlsxBlackBoxOverrides,
  options: Parameters<typeof parseXlsx>[1] = { errorMode: 'strict' },
): Promise<XlsxParseError> {
  try {
    await parseXlsx(await bytes(changes), options);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected XLSX pivot parsing to fail');
}

describe('XLSX pivot caches', () => {
  it.each([
    ['.55', 0.55],
    ['1.', 1],
    ['1e30', 1e30],
    ['1e+30', 1e30],
    ['-2.5E-2', -0.025],
    ['-0', 0],
  ] as const)('normalizes pivot finite number %s', (source, expected) => {
    expect(
      parseXlsxPivotFiniteNumber(source, 'Pivot test number', 'pivot.xml'),
    ).toBe(expected);
  });

  it.each([' 1', '1 ', '1e', '1e309', {}, null, undefined])(
    'rejects invalid pivot finite number %#',
    (source) => {
      expect(() =>
        parseXlsxPivotFiniteNumber(source, 'Pivot test number', 'pivot.xml'),
      ).toThrow('Pivot test number');
    },
  );

  it('rejects trailing whitespace in a signed pivot integer', () => {
    expect(() =>
      parseXlsxPivotSignedInteger('1 ', 'Pivot signed whitespace', 'pivot.xml'),
    ).toThrow('Pivot signed whitespace');
  });

  it('reports direct pivot record-count and field-index structure errors', () => {
    expect(() => parseXlsxPivotRecordCount('bad', 'pivot.xml')).toThrow(
      'Pivot cache record count is invalid',
    );
    expect(() =>
      parseXlsxPivotFieldIndexes('bad', 'column fields', 'pivot.xml'),
    ).toThrow('Pivot column fields are invalid');
  });

  it.each([0, 1, 9007199254740991])(
    'normalizes pivot unsigned integer %s',
    (source) => {
      expect(
        parseXlsxPivotUnsignedInteger(
          String(source),
          'Pivot test integer',
          'pivot.xml',
        ),
      ).toBe(source);
    },
  );

  it.each(['-1', '01', '1.0', 'x1', '1x', '9007199254740992', {}, null])(
    'rejects invalid pivot unsigned integer %#',
    (source) => {
      expect(() =>
        parseXlsxPivotUnsignedInteger(
          source,
          'Pivot test integer',
          'pivot.xml',
        ),
      ).toThrow('Pivot test integer');
    },
  );

  it.each([
    ['-2', -2],
    ['0', 0],
    ['2', 2],
  ] as const)('normalizes pivot signed integer %s', (source, expected) => {
    expect(
      parseXlsxPivotSignedInteger(source, 'Pivot test signed', 'pivot.xml'),
    ).toBe(expected);
  });

  it.each([
    '-0',
    '+1',
    '01',
    '1.0',
    'x1',
    '1x',
    '12a',
    '1 ',
    '-1 ',
    '9007199254740992',
    '-9007199254740992',
    {},
    null,
  ])('rejects invalid pivot signed integer %#', (source) => {
    expect(() =>
      parseXlsxPivotSignedInteger(source, 'Pivot test signed', 'pivot.xml'),
    ).toThrow('Pivot test signed');
  });

  it.each([
    '2024-02-29T23:59:59Z',
    '2020-02-29T00:00:00',
    '2000-02-29T00:00:00+09:30',
    '0001-01-01T00:00:00Z',
    `${'9'.repeat(400)}2000-02-29T00:00:00Z`,
    '2024-01-01T00:00:00.125+14:00',
    '12024-12-31T12:30:15-13:59',
  ])('normalizes pivot date-time %s', (source) => {
    expect(parseXlsxPivotDateTime(source, 'Pivot test date', 'pivot.xml')).toBe(
      source,
    );
  });

  it.each([
    '0000-01-01T00:00:00Z',
    '2023-02-29T00:00:00Z',
    '1900-02-29T00:00:00Z',
    '2024-04-31T00:00:00Z',
    '2024-06-31T00:00:00Z',
    '2024-09-31T00:00:00Z',
    '2024-11-31T00:00:00Z',
    '2024-00-01T00:00:00Z',
    '2024-13-01T00:00:00Z',
    '2024-01-00T00:00:00Z',
    '2024-01-32T00:00:00Z',
    '2024-01-01T24:00:00Z',
    '2024-01-01T00:60:00Z',
    '2024-01-01T00:00:60Z',
    '2024-01-01',
    'x2024-01-01T00:00:00Z',
    '2024-01-01T00:00:00Zx',
    {},
    null,
  ])('rejects invalid pivot date-time %#', (source) => {
    expect(() =>
      parseXlsxPivotDateTime(source, 'Pivot test date', 'pivot.xml'),
    ).toThrow('Pivot test date');
  });

  it.each([
    ['difference', 'difference'],
    ['index', 'index'],
    ['normal', 'normal'],
    ['percent', 'percent'],
    ['percentDiff', 'percentDifference'],
    ['percentOfCol', 'percentOfColumn'],
    ['percentOfRow', 'percentOfRow'],
    ['percentOfTotal', 'percentOfTotal'],
    ['runTotal', 'runningTotal'],
    [undefined, 'normal'],
  ] as const)('maps pivot display mode %#', (source, expected) => {
    expect(parseXlsxPivotDataDisplayMode(source, 'pivot.xml')).toBe(expected);
  });

  it.each([
    ['average', 'average'],
    ['count', 'count'],
    ['countNums', 'countNumbers'],
    ['max', 'maximum'],
    ['min', 'minimum'],
    ['product', 'product'],
    ['stdDev', 'standardDeviation'],
    ['stdDevp', 'standardDeviationPopulation'],
    ['sum', 'sum'],
    ['var', 'variance'],
    ['varp', 'variancePopulation'],
    [undefined, 'sum'],
  ] as const)('maps pivot subtotal %#', (source, expected) => {
    expect(parseXlsxPivotSubtotal(source, 'pivot.xml')).toBe(expected);
  });

  it('rejects unknown pivot display and subtotal values', () => {
    expect(() => parseXlsxPivotDataDisplayMode('bad', 'pivot.xml')).toThrow(
      'Pivot display mode is invalid',
    );
    expect(() => parseXlsxPivotSubtotal('bad', 'pivot.xml')).toThrow(
      'Pivot subtotal is invalid',
    );
  });

  it.each([
    'captionBeginsWith',
    'captionBetween',
    'captionContains',
    'captionEndsWith',
    'captionEqual',
    'captionGreaterThan',
    'captionGreaterThanOrEqual',
    'captionLessThan',
    'captionLessThanOrEqual',
    'captionNotBeginsWith',
    'captionNotBetween',
    'captionNotContains',
    'captionNotEndsWith',
    'captionNotEqual',
    'count',
    'dateBetween',
    'dateEqual',
    'dateNewerThan',
    'dateNewerThanOrEqual',
    'dateNotBetween',
    'dateNotEqual',
    'dateOlderThan',
    'dateOlderThanOrEqual',
    'lastMonth',
    'lastQuarter',
    'lastWeek',
    'lastYear',
    'month1',
    'month10',
    'month11',
    'month12',
    'month2',
    'month3',
    'month4',
    'month5',
    'month6',
    'month7',
    'month8',
    'month9',
    'nextMonth',
    'nextQuarter',
    'nextWeek',
    'nextYear',
    'percent',
    'quarter1',
    'quarter2',
    'quarter3',
    'quarter4',
    'sum',
    'thisMonth',
    'thisQuarter',
    'thisWeek',
    'thisYear',
    'today',
    'tomorrow',
    'unknown',
    'valueBetween',
    'valueEqual',
    'valueGreaterThan',
    'valueGreaterThanOrEqual',
    'valueLessThan',
    'valueLessThanOrEqual',
    'valueNotBetween',
    'valueNotEqual',
    'yearToDate',
    'yesterday',
  ] as const)('accepts pivot filter type %s', (source) => {
    expect(parseXlsxPivotFilterType(source, 'pivot.xml')).toBe(source);
  });

  it.each([undefined, '', 'bad'])('rejects pivot filter type %#', (source) => {
    expect(() => parseXlsxPivotFilterType(source, 'pivot.xml')).toThrow(
      'Pivot filter type is invalid',
    );
  });

  it('validates pivot record capture root, nesting, text, and completion', () => {
    expect(captureSync(() => recordSink().result()).diagnostic.message).toBe(
      'Pivot cache-record capture is incomplete',
    );
    const wrongNamespace = recordSink();
    expect(
      captureSync(() =>
        wrongNamespace.openElement(
          pivotElement('pivotCacheRecords', 'urn:wrong'),
        ),
      ).diagnostic.message,
    ).toBe('Pivot cache-record element has the wrong namespace');
    const wrongRoot = recordSink();
    expect(
      captureSync(() => wrongRoot.openElement(pivotElement('wrong'))).diagnostic
        .message,
    ).toBe('Pivot cache-record root is missing');
    const incomplete = recordSink();
    incomplete.openElement(pivotElement('pivotCacheRecords'));
    expect(captureSync(() => incomplete.result()).diagnostic.message).toBe(
      'Pivot cache-record capture is incomplete',
    );
    const nesting = recordSink();
    nesting.openElement(pivotElement('pivotCacheRecords'));
    expect(
      captureSync(() => nesting.closeElement(pivotElement('wrong'))).diagnostic
        .message,
    ).toBe('Pivot cache-record nesting is invalid');
    const text = recordSink();
    text.openElement(pivotElement('pivotCacheRecords'));
    expect(() => text.text(' \n\t ')).not.toThrow();
    expect(captureSync(() => text.text('bad')).diagnostic.message).toBe(
      'Pivot cache-record text is invalid',
    );
    const duplicate = recordSink();
    duplicate.openElement(pivotElement('pivotCacheRecords'));
    duplicate.closeElement(pivotElement('pivotCacheRecords'));
    expect(
      captureSync(() =>
        duplicate.openElement(pivotElement('pivotCacheRecords')),
      ).diagnostic.message,
    ).toBe('Pivot cache-record root is duplicated');
    const valid = recordSink();
    valid.openElement(
      pivotElement(
        'pivotCacheRecords',
        XLSX_SPREADSHEET_NS,
        new Map([['{}count', '0']]),
      ),
    );
    valid.closeElement(pivotElement('pivotCacheRecords'));
    expect(valid.result()).toStrictEqual([]);

    const countMismatch = recordSink();
    countMismatch.openElement(
      pivotElement(
        'pivotCacheRecords',
        XLSX_SPREADSHEET_NS,
        new Map([['{}count', '1']]),
      ),
    );
    countMismatch.closeElement(pivotElement('pivotCacheRecords'));
    expect(captureSync(() => countMismatch.result()).diagnostic.message).toBe(
      'Pivot cache-record count does not match',
    );

    const wrongChild = recordSink();
    wrongChild.openElement(pivotElement('pivotCacheRecords'));
    expect(
      captureSync(() => wrongChild.openElement(pivotElement('wrong')))
        .diagnostic.message,
    ).toBe('Pivot cache-record structure is invalid');

    const oneField = recordSink([
      {
        databaseField: true,
        name: 'Value',
        serverField: false,
        uniqueList: true,
      },
    ]);
    oneField.openElement(pivotElement('pivotCacheRecords'));
    oneField.openElement(pivotElement('r'));
    oneField.openElement(
      pivotElement('s', XLSX_SPREADSHEET_NS, new Map([['{}v', 'value']])),
    );
    oneField.closeElement(pivotElement('s'));
    oneField.closeElement(pivotElement('r'));
    oneField.closeElement(pivotElement('pivotCacheRecords'));
    expect(oneField.result()).toStrictEqual([
      [{ kind: 'text', value: 'value' }],
    ]);

    const unknownValue = recordSink([
      {
        databaseField: true,
        name: 'Value',
        serverField: false,
        uniqueList: true,
      },
    ]);
    unknownValue.openElement(pivotElement('pivotCacheRecords'));
    unknownValue.openElement(pivotElement('r'));
    expect(
      captureSync(() =>
        unknownValue.openElement(
          pivotElement('z', XLSX_SPREADSHEET_NS, new Map([['{}v', 'value']])),
        ),
      ).diagnostic.message,
    ).toBe('Pivot cache item type is invalid');

    const nestedValue = recordSink([
      {
        databaseField: true,
        name: 'Value',
        serverField: false,
        uniqueList: true,
      },
    ]);
    nestedValue.openElement(pivotElement('pivotCacheRecords'));
    nestedValue.openElement(pivotElement('r'));
    nestedValue.openElement(
      pivotElement('s', XLSX_SPREADSHEET_NS, new Map([['{}v', 'value']])),
    );
    expect(
      captureSync(() => nestedValue.openElement(pivotElement('wrong')))
        .diagnostic.message,
    ).toBe('Pivot cache-record structure is invalid');

    const nestedRecord = recordSink([
      {
        databaseField: true,
        name: 'Value',
        serverField: false,
        uniqueList: true,
      },
    ]);
    nestedRecord.openElement(pivotElement('pivotCacheRecords'));
    nestedRecord.openElement(pivotElement('r'));
    nestedRecord.openElement(
      pivotElement('s', XLSX_SPREADSHEET_NS, new Map([['{}v', 'value']])),
    );
    expect(
      captureSync(() => nestedRecord.openElement(pivotElement('r'))).diagnostic
        .message,
    ).toBe('Pivot cache-record structure is invalid');

    const missingSharedIndex = recordSink([
      {
        databaseField: true,
        items: [{ kind: 'text', value: 'Value' }],
        name: 'Value',
        serverField: false,
        uniqueList: true,
      },
    ]);
    missingSharedIndex.openElement(pivotElement('pivotCacheRecords'));
    missingSharedIndex.openElement(pivotElement('r'));
    expect(
      captureSync(() => missingSharedIndex.openElement(pivotElement('x')))
        .diagnostic.message,
    ).toBe('Pivot record shared-item index is invalid');
  });
  it('parses bounded cache metadata without reading records by default', async () => {
    const document = await parseXlsx(await bytes(), { errorMode: 'strict' });
    expect(document.pivotCaches).toStrictEqual([
      {
        backgroundQuery: true,
        enableRefresh: false,
        fields: [
          {
            databaseField: true,
            items: [
              { kind: 'text', value: 'A' },
              { kind: 'text', value: 'B' },
            ],
            name: 'Category',
            serverField: false,
            sharedItems: {
              containsBlank: false,
              containsDate: false,
              containsInteger: false,
              containsMixedTypes: false,
              containsNonDate: false,
              containsNumber: false,
              containsSemiMixedTypes: false,
              containsString: true,
              longText: false,
            },
            uniqueList: true,
          },
          {
            databaseField: true,
            items: [
              { kind: 'number', value: 1 },
              { kind: 'number', value: 2 },
            ],
            name: 'Amount',
            serverField: false,
            sharedItems: {
              containsBlank: false,
              containsDate: false,
              containsInteger: true,
              containsMixedTypes: false,
              containsNonDate: false,
              containsNumber: true,
              containsSemiMixedTypes: false,
              containsString: false,
              longText: false,
              maximumNumber: 2,
              minimumNumber: 1,
            },
            uniqueList: true,
          },
        ],
        index: 0,
        missingItemsLimit: 10,
        recordCount: 2,
        refreshedBy: 'Producer',
        refreshedDate: 45292.5,
        refreshOnLoad: true,
        saveData: true,
        source: {
          kind: 'worksheet',
          range: {
            end: { column: 2, row: 3 },
            reference: 'A1:B3',
            start: { column: 1, row: 1 },
          },
          sheet: 'Sheet1',
        },
        supportAdvancedDrill: false,
        tupleCache: true,
        upgradeOnRefresh: true,
      },
    ]);
  });

  it('parses pivot table fields, axes, filters, style, cache mapping, and selection', async () => {
    const document = await parseXlsx(await bytes(), { errorMode: 'strict' });
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    expect(
      sheet.kind === 'worksheet' ? sheet.pivotTables : undefined,
    ).toStrictEqual([
      {
        cacheIndex: 0,
        columnFields: [-2],
        compact: false,
        dataCaption: 'Values',
        dataFields: [
          {
            baseField: 0,
            baseItem: 0,
            field: 1,
            name: 'Sum Amount',
            showDataAs: 'percentOfTotal',
            subtotal: 'sum',
          },
        ],
        fields: [
          {
            axis: 'row',
            compact: true,
            dataField: false,
            items: [
              { sharedItemIndex: 0, type: 'data' },
              { sharedItemIndex: 1, type: 'blank' },
            ],
            name: 'Category',
            outline: false,
            showAll: false,
            sortType: 'ascending',
            subtotalTop: false,
          },
          {
            axis: 'values',
            compact: true,
            dataField: true,
            items: [],
            outline: true,
            showAll: true,
            sortType: 'manual',
            subtotalTop: true,
          },
        ],
        filters: [
          {
            description: 'Only A',
            evaluationOrder: -1,
            field: 0,
            id: 1,
            name: 'Filter',
            stringValue1: 'A',
            type: 'captionEqual',
          },
        ],
        grandTotalCaption: 'Grand',
        location: {
          end: { column: 8, row: 10 },
          reference: 'D1:H10',
          start: { column: 4, row: 1 },
        },
        name: 'SalesPivot',
        outline: true,
        pageFields: [{ field: 0, hierarchy: 0, item: 1, name: 'Page' }],
        rowFields: [0],
        selectionRelation: 'full-sheet',
        showColumnGrandTotals: true,
        showHeaders: true,
        showRowGrandTotals: false,
        style: {
          name: 'PivotStyleMedium9',
          showColumnHeaders: true,
          showColumnStripes: false,
          showLastColumn: false,
          showRowHeaders: true,
          showRowStripes: true,
        },
      },
    ]);

    const selected = await parseXlsx(await bytes(), {
      errorMode: 'strict',
      selection: { ranges: { Sheet1: ['H10'] } },
    });
    const selectedSheet = selected.sheets[0]!;
    expect(
      selectedSheet.kind === 'worksheet'
        ? selectedSheet.pivotTables?.[0]?.selectionRelation
        : undefined,
    ).toBe('intersects-selection');
    const excluded = await parseXlsx(await bytes(), {
      errorMode: 'strict',
      selection: { ranges: { Sheet1: ['A20'] } },
    });
    const excludedSheet = excluded.sheets[0]!;
    expect(excludedSheet).not.toHaveProperty('pivotTables');
  });

  it('preserves cache, field, table, and style defaults exactly', async () => {
    const cache = `<pivotCacheDefinition xmlns="${XLSX_SPREADSHEET_NS}"><cacheSource type="scenario"/><cacheFields count="1"><cacheField name="Field"/></cacheFields></pivotCacheDefinition>`;
    const table = `<pivotTableDefinition xmlns="${XLSX_SPREADSHEET_NS}" name="Defaults" cacheId="7" dataCaption="Values"><location ref="D1:E2"/><pivotFields count="1"><pivotField/></pivotFields></pivotTableDefinition>`;
    const document = await parseXlsx(
      await bytes({
        'xl/pivotCache/pivotCacheDefinition1.xml': cache,
        'xl/pivotTables/pivotTable1.xml': table,
      }),
      { errorMode: 'strict' },
    );
    expect(document.pivotCaches).toStrictEqual([
      {
        backgroundQuery: false,
        enableRefresh: true,
        fields: [
          {
            databaseField: true,
            name: 'Field',
            serverField: false,
            uniqueList: true,
          },
        ],
        index: 0,
        refreshOnLoad: false,
        saveData: true,
        source: { kind: 'scenario' },
        supportAdvancedDrill: true,
        tupleCache: false,
        upgradeOnRefresh: false,
      },
    ]);
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.pivotTables : undefined,
    ).toStrictEqual([
      {
        cacheIndex: 0,
        columnFields: [],
        compact: true,
        dataCaption: 'Values',
        dataFields: [],
        fields: [
          {
            compact: true,
            dataField: false,
            items: [],
            outline: true,
            showAll: true,
            sortType: 'manual',
            subtotalTop: true,
          },
        ],
        filters: [],
        location: {
          end: { column: 5, row: 2 },
          reference: 'D1:E2',
          start: { column: 4, row: 1 },
        },
        name: 'Defaults',
        outline: false,
        pageFields: [],
        rowFields: [],
        selectionRelation: 'full-sheet',
        showColumnGrandTotals: true,
        showHeaders: true,
        showRowGrandTotals: true,
        style: {
          showColumnHeaders: true,
          showColumnStripes: false,
          showLastColumn: false,
          showRowHeaders: true,
          showRowStripes: false,
        },
      },
    ]);
  });

  it('accepts empty optional pivot containers and explicit manual sorting', async () => {
    const table = `<pivotTableDefinition xmlns="${XLSX_SPREADSHEET_NS}" name="Empty" cacheId="7" dataCaption="Values"><location ref="D1:E2"/><pivotFields count="1"><pivotField sortType="manual"><items/></pivotField></pivotFields><rowFields/><colFields/><pageFields/><dataFields/><filters/></pivotTableDefinition>`;
    const document = await parseXlsx(
      await bytes({
        'xl/pivotCache/pivotCacheDefinition1.xml': `<pivotCacheDefinition xmlns="${XLSX_SPREADSHEET_NS}"><cacheSource type="scenario"/><cacheFields count="1"><cacheField name="Field"/></cacheFields></pivotCacheDefinition>`,
        'xl/pivotTables/pivotTable1.xml': table,
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const pivot =
      sheet.kind === 'worksheet' ? sheet.pivotTables?.[0] : undefined;
    expect(pivot).toMatchObject({
      columnFields: [],
      dataFields: [],
      filters: [],
      pageFields: [],
      rowFields: [],
    });
    expect(pivot?.fields[0]).toMatchObject({ items: [], sortType: 'manual' });
  });

  it('accepts omitted optional collection counts with authored children', async () => {
    const table = pivotTable()
      .replace('<items count="2">', '<items>')
      .replace('<rowFields count="1">', '<rowFields>')
      .replace('<colFields count="1">', '<colFields>')
      .replace('<pageFields count="1">', '<pageFields>')
      .replace('<dataFields count="1">', '<dataFields>')
      .replace('<filters count="1">', '<filters>');
    const document = await parseXlsx(
      await bytes({ 'xl/pivotTables/pivotTable1.xml': table }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.pivotTables?.[0] : undefined,
    ).toMatchObject({
      columnFields: [-2],
      dataFields: [{ field: 1 }],
      filters: [{ field: 0 }],
      pageFields: [{ field: 0 }],
      rowFields: [0],
    });
  });

  it('maps every pivot field axis and item type in authored order', async () => {
    const cache = `<pivotCacheDefinition xmlns="${XLSX_SPREADSHEET_NS}"><cacheSource type="scenario"/><cacheFields count="4"><cacheField name="A"/><cacheField name="B"/><cacheField name="C"/><cacheField name="D"/></cacheFields></pivotCacheDefinition>`;
    const itemTypes = [
      undefined,
      'avg',
      'blank',
      'count',
      'default',
      'grand',
      'max',
      'min',
      'product',
      'stdDev',
      'stdDevP',
      'sum',
      'var',
      'varP',
    ];
    const items = itemTypes
      .map((type, index) => `<item x="${index}"${type ? ` t="${type}"` : ''}/>`)
      .join('');
    const table = `<pivotTableDefinition xmlns="${XLSX_SPREADSHEET_NS}" name="Axes" cacheId="7" dataCaption="Values"><location ref="D1:H10"/><pivotFields count="4"><pivotField axis="axisRow" sortType="descending"><items count="14">${items}</items></pivotField><pivotField axis="axisCol"/><pivotField axis="axisPage"/><pivotField axis="axisValues"/></pivotFields></pivotTableDefinition>`;
    const document = await parseXlsx(
      await bytes({
        'xl/pivotCache/pivotCacheDefinition1.xml': cache,
        'xl/pivotTables/pivotTable1.xml': table,
      }),
      { errorMode: 'strict', pivotCacheMode: 'none' },
    );
    const sheet = document.sheets[0]!;
    const fields =
      sheet.kind === 'worksheet' ? sheet.pivotTables?.[0]?.fields : undefined;
    expect(fields?.map((field) => field.axis)).toStrictEqual([
      'row',
      'column',
      'page',
      'values',
    ]);
    expect(fields?.[0]?.sortType).toBe('descending');
    expect(fields?.[0]?.items.map((item) => item.type)).toStrictEqual([
      'data',
      'average',
      'blank',
      'count',
      'default',
      'grand-total',
      'maximum',
      'minimum',
      'product',
      'standard-deviation',
      'standard-deviation-population',
      'sum',
      'variance',
      'variance-population',
    ]);
  });

  it('preserves optional pivot data/filter fields and their defaults', async () => {
    const table = pivotTable()
      .replace(
        '<dataField name="Sum Amount" fld="1" subtotal="sum" showDataAs="percentOfTotal" baseField="0" baseItem="0"/>',
        '<dataField fld="1"/>',
      )
      .replace(
        'stringValue1="A"',
        'iMeasureFld="1" iMeasureHier="2" stringValue1="A" stringValue2="B"',
      );
    const document = await parseXlsx(
      await bytes({ 'xl/pivotTables/pivotTable1.xml': table }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const pivot =
      sheet.kind === 'worksheet' ? sheet.pivotTables?.[0] : undefined;
    expect(pivot?.dataFields).toStrictEqual([
      { field: 1, showDataAs: 'normal', subtotal: 'sum' },
    ]);
    expect(pivot?.filters[0]).toMatchObject({
      measureField: 1,
      measureHierarchy: 2,
      stringValue1: 'A',
      stringValue2: 'B',
    });
  });

  it('rejects case-insensitive duplicate pivot table names', async () => {
    const contentTypes = CONTENT_TYPES.replace(
      '</Types>',
      '<Override PartName="/xl/pivotTables/pivotTable2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/></Types>',
    );
    const worksheet = WORKSHEET.replace(
      '<pivotTableParts count="1"><pivotTablePart r:id="pivot"/></pivotTableParts>',
      '<pivotTableParts count="2"><pivotTablePart r:id="pivot"/><pivotTablePart r:id="pivot2"/></pivotTableParts>',
    );
    const relationships = WORKSHEET_RELS.replace(
      '</Relationships>',
      `<Relationship Id="pivot2" Type="${PIVOT_TABLE_RELATIONSHIP}" Target="../pivotTables/pivotTable2.xml"/></Relationships>`,
    );
    expect(
      (
        await capture({
          '[Content_Types].xml': contentTypes,
          'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
            'name="SalesPivot"',
            'name="Straße"',
          ),
          'xl/pivotTables/pivotTable2.xml': pivotTable().replace(
            'name="SalesPivot"',
            'name="STRASSE"',
          ),
          'xl/worksheets/_rels/sheet1.xml.rels': relationships,
          'xl/worksheets/sheet1.xml': worksheet,
        })
      ).diagnostic.message,
    ).toBe('Worksheet contains duplicate pivot table names');
  });

  it.each([
    [
      '<cacheSource type="external" connectionId="3"/>',
      { connectionId: 3, kind: 'external' },
    ],
    ['<cacheSource type="consolidation"/>', { kind: 'consolidation' }],
    ['<cacheSource type="scenario"/>', { kind: 'scenario' }],
    [
      '<cacheSource type="worksheet"><worksheetSource name="DynamicSource"/></cacheSource>',
      { kind: 'worksheet', name: 'DynamicSource' },
    ],
  ] as const)('parses pivot cache source %#', async (source, expected) => {
    const document = await parseXlsx(
      await bytes({
        'xl/pivotCache/pivotCacheDefinition1.xml':
          cacheDefinitionWithSource(source),
      }),
      { errorMode: 'strict' },
    );
    expect(document.pivotCaches?.[0]?.source).toStrictEqual(expected);
  });

  it('charges pivot worksheet ranges but not named sources', async () => {
    const contentTypes = CONTENT_TYPES.replace(
      '</Types>',
      '<Override PartName="/xl/pivotCache/pivotCacheDefinition2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/></Types>',
    );
    const workbook = WORKBOOK.replace(
      '</pivotCaches>',
      '<pivotCache cacheId="8" r:id="cache2"/></pivotCaches>',
    );
    const relationships = WORKBOOK_RELS.replace(
      '</Relationships>',
      `<Relationship Id="cache2" Type="${CACHE_DEFINITION_RELATIONSHIP}" Target="pivotCache/pivotCacheDefinition2.xml"/></Relationships>`,
    );
    const sourceParts = {
      '[Content_Types].xml': contentTypes,
      'xl/_rels/workbook.xml.rels': relationships,
      'xl/workbook.xml': workbook,
      'xl/worksheets/sheet1.xml': WORKSHEET_WITHOUT_PIVOT,
      'xl/worksheets/_rels/sheet1.xml.rels': null,
    } as const;
    const namedDefinition = cacheDefinitionWithSource(
      '<cacheSource type="worksheet"><worksheetSource name="NamedSource"/></cacheSource>',
    );
    await expect(
      parseXlsx(
        await bytes({
          ...sourceParts,
          'xl/pivotCache/pivotCacheDefinition1.xml': namedDefinition,
          'xl/pivotCache/pivotCacheDefinition2.xml': namedDefinition,
        }),
        { errorMode: 'strict', limits: { maxRangeAreas: 1 } },
      ),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(
          {
            ...sourceParts,
            'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition(),
            'xl/pivotCache/pivotCacheDefinition2.xml': cacheDefinition(),
          },
          { errorMode: 'strict', limits: { maxRangeAreas: 1 } },
        )
      ).diagnostic,
    ).toMatchObject({ actual: 2, limitName: 'maxRangeAreas' });
  });

  it('parses a complete Strict pivot cache, records, and table graph', async () => {
    const strictSheet = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictRelationship =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const source = await createIndependentXlsx(
      parts({
        '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelationship}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELS.replaceAll(
          XLSX_OFFICE_REL_TYPE,
          `${strictRelationship}/`,
        ),
        'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels':
          CACHE_RELS.replaceAll(XLSX_OFFICE_REL_TYPE, `${strictRelationship}/`),
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replaceAll(
          XLSX_SPREADSHEET_NS,
          strictSheet,
        ),
        'xl/pivotCache/pivotCacheRecords1.xml': cacheRecords().replaceAll(
          XLSX_SPREADSHEET_NS,
          strictSheet,
        ),
        'xl/pivotTables/pivotTable1.xml': pivotTable().replaceAll(
          XLSX_SPREADSHEET_NS,
          strictSheet,
        ),
        'xl/workbook.xml': WORKBOOK.replaceAll(
          XLSX_SPREADSHEET_NS,
          strictSheet,
        ).replaceAll(XLSX_OFFICE_REL_NS, strictRelationship),
        'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replaceAll(
          XLSX_OFFICE_REL_TYPE,
          `${strictRelationship}/`,
        ),
        'xl/worksheets/sheet1.xml': WORKSHEET.replaceAll(
          XLSX_SPREADSHEET_NS,
          strictSheet,
        ).replaceAll(XLSX_OFFICE_REL_NS, strictRelationship),
      }),
    );
    const document = await parseXlsx(source, {
      errorMode: 'strict',
      pivotCacheMode: 'records',
    });
    expect(document.pivotCaches?.[0]?.records).toHaveLength(2);
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.pivotTables?.[0]?.name : undefined,
    ).toBe('SalesPivot');
  });

  it('parses namespace-prefixed pivot definitions and children', async () => {
    const prefixElements = (xml: string, prefix: 'p' | 'x'): string =>
      xml
        .replace(
          `xmlns="${XLSX_SPREADSHEET_NS}"`,
          `xmlns:${prefix}="${XLSX_SPREADSHEET_NS}"`,
        )
        .replaceAll(
          /<(\/?)((?:pivotCacheDefinition|cacheSource|worksheetSource|cacheFields|cacheField|sharedItems|pivotTableDefinition|location|pivotFields|pivotField|items|item|rowFields|colFields|field|pageFields|pageField|dataFields|dataField|filters|filter|pivotTableStyleInfo|s|n))(?=[\s/>])/gu,
          `<$1${prefix}:$2`,
        );
    for (const prefix of ['p', 'x'] as const) {
      const document = await parseXlsx(
        await bytes({
          'xl/pivotCache/pivotCacheDefinition1.xml': prefixElements(
            cacheDefinition(),
            prefix,
          ),
          'xl/pivotTables/pivotTable1.xml': prefixElements(
            pivotTable(),
            prefix,
          ),
        }),
        { errorMode: 'strict' },
      );
      expect(
        document.pivotCaches?.[0]?.fields.map((field) => field.name),
      ).toEqual(['Category', 'Amount']);
      const sheet = document.sheets[0]!;
      expect(
        sheet.kind === 'worksheet' ? sheet.pivotTables?.[0]?.name : undefined,
      ).toBe('SalesPivot');
    }
  });

  it.each([
    ['D1', true],
    ['H10', true],
    ['C1', false],
    ['I1', false],
    ['D11', false],
    ['A10', false],
  ] as const)(
    'classifies pivot selection boundary %s',
    async (range, included) => {
      const document = await parseXlsx(await bytes(), {
        errorMode: 'strict',
        selection: { ranges: { Sheet1: [range] } },
      });
      const sheet = document.sheets[0]!;
      expect(
        sheet.kind === 'worksheet'
          ? (sheet.pivotTables?.length ?? 0) > 0
          : false,
      ).toBe(included);
    },
  );

  it('includes a pivot when any one of several selected ranges intersects it', async () => {
    const document = await parseXlsx(await bytes(), {
      errorMode: 'strict',
      selection: { ranges: { Sheet1: ['A20', 'E2'] } },
    });
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.pivotTables?.[0]?.name : undefined,
    ).toBe('SalesPivot');
  });

  it('distinguishes vertical range separation from boundary overlap', async () => {
    const table = pivotTable().replace('ref="D1:H10"', 'ref="D5:H10"');
    const separated = await parseXlsx(
      await bytes({ 'xl/pivotTables/pivotTable1.xml': table }),
      {
        errorMode: 'strict',
        selection: { ranges: { Sheet1: ['D1:H4'] } },
      },
    );
    const separatedSheet = separated.sheets[0]!;
    expect(
      separatedSheet.kind === 'worksheet'
        ? separatedSheet.pivotTables
        : undefined,
    ).toBeUndefined();

    const touching = await parseXlsx(
      await bytes({ 'xl/pivotTables/pivotTable1.xml': table }),
      {
        errorMode: 'strict',
        selection: { ranges: { Sheet1: ['D4:H5'] } },
      },
    );
    const touchingSheet = touching.sheets[0]!;
    expect(
      touchingSheet.kind === 'worksheet'
        ? touchingSheet.pivotTables?.[0]?.name
        : undefined,
    ).toBe('SalesPivot');
  });

  it('accepts unrelated cache relationships and rejects duplicate record owners', async () => {
    const withUnrelated = CACHE_RELS.replace(
      '</Relationships>',
      `<Relationship Id="image" Type="${XLSX_OFFICE_REL_TYPE}image" Target="../media/image1.png"/></Relationships>`,
    );
    await expect(
      parseXlsx(
        await bytes({
          'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': withUnrelated,
        }),
        { errorMode: 'strict', pivotCacheMode: 'records' },
      ),
    ).resolves.toBeDefined();

    const duplicate = CACHE_RELS.replace(
      '</Relationships>',
      `<Relationship Id="records2" Type="${CACHE_RECORDS_RELATIONSHIP}" Target="pivotCacheRecords1.xml"/></Relationships>`,
    );
    expect(
      (
        await capture(
          {
            'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': duplicate,
          },
          { errorMode: 'strict', pivotCacheMode: 'records' },
        )
      ).diagnostic.message,
    ).toBe('Pivot cache-record relationship is invalid');
  });

  it('validates workbook pivot-cache declarations and ownership exactly', async () => {
    expect(
      (
        await capture({
          'xl/workbook.xml': WORKBOOK.replace(
            /<pivotCaches>[\s\S]*?<\/pivotCaches>/u,
            '<pivotCaches>bad</pivotCaches>',
          ),
        })
      ).diagnostic.message,
    ).toBe('Workbook pivot-cache declarations are invalid');

    expect(
      (
        await capture({
          'xl/workbook.xml': WORKBOOK.replace(
            '</pivotCaches>',
            '<pivotCache cacheId="7" r:id="cache"/></pivotCaches>',
          ),
        })
      ).diagnostic.message,
    ).toBe('Workbook contains duplicate pivot cache IDs');

    expect(
      (
        await capture({
          'xl/workbook.xml': WORKBOOK.replace('r:id="cache"', 'r:id=""'),
        })
      ).diagnostic.message,
    ).toBe('Workbook pivot cache relationship reference is invalid');

    expect(
      (
        await capture({
          'xl/_rels/workbook.xml.rels': WORKBOOK_RELS.replace(
            'Target="pivotCache/pivotCacheDefinition1.xml"',
            'Target="https://example.invalid/cache.xml" TargetMode="External"',
          ),
        })
      ).diagnostic.message,
    ).toBe('Workbook pivot cache relationship is invalid');
  });

  it('validates worksheet pivot-table part declarations exactly', async () => {
    const cases = [
      [
        WORKSHEET.replace('pivotTableParts count="1"', 'pivotTableParts'),
        'Worksheet pivot-table parts must not be empty',
      ],
      [
        WORKSHEET.replace(
          'pivotTableParts count="1"',
          'pivotTableParts count="0"',
        ),
        'Worksheet pivot-table parts must not be empty',
      ],
      [
        WORKSHEET.replace(
          'pivotTableParts count="1"',
          'pivotTableParts count="bad"',
        ),
        'Worksheet pivot-table-part count is invalid',
      ],
      [
        WORKSHEET.replace(
          'pivotTableParts count="1"',
          'pivotTableParts count="2"',
        ),
        'Worksheet pivot-table-part count does not match',
      ],
      [
        WORKSHEET.replace(
          '</pivotTableParts>',
          '</pivotTableParts><pivotTableParts count="1"><pivotTablePart r:id="pivot"/></pivotTableParts>',
        ),
        'Worksheet contains duplicate pivotTableParts elements',
      ],
      [
        WORKSHEET.replace('r:id="pivot"', 'r:id=""'),
        'Worksheet pivot-table relationship reference is invalid',
      ],
      [
        WORKSHEET.replace(
          '<pivotTableParts count="1"><pivotTablePart r:id="pivot"/></pivotTableParts>',
          '<pivotTableParts count="2"><pivotTablePart r:id="pivot"/><pivotTablePart r:id="pivot"/></pivotTableParts>',
        ),
        'Worksheet contains duplicate pivot-table relationships',
      ],
    ] as const;
    for (const [worksheet, message] of cases) {
      expect(
        (await capture({ 'xl/worksheets/sheet1.xml': worksheet })).diagnostic
          .message,
      ).toBe(message);
    }

    const withForeignChild = WORKSHEET.replace(
      '<pivotTablePart r:id="pivot"/>',
      '<foreign r:id="ignored"/><pivotTablePart r:id="pivot"/>',
    );
    expect(
      (await capture({ 'xl/worksheets/sheet1.xml': withForeignChild }))
        .diagnostic.message,
    ).toBe('Worksheet element nesting is invalid');

    const wrongOwner = WORKSHEET.replace(
      '<sheetData>',
      '<sheetViews><sheetView workbookViewId="0"><pivotTablePart r:id="pivot"/></sheetView></sheetViews><sheetData>',
    ).replace(/<pivotTableParts[\s\S]*?<\/pivotTableParts>/u, '');
    expect(
      (await capture({ 'xl/worksheets/sheet1.xml': wrongOwner })).diagnostic
        .message,
    ).toBe('Worksheet element nesting is invalid');
  });

  it('does not require a records part for zero or omitted cache record counts', async () => {
    for (const definition of [
      cacheDefinition().replace('recordCount="2"', 'recordCount="0"'),
      cacheDefinition().replace('recordCount="2"', ''),
    ]) {
      const document = await parseXlsx(
        await bytes({
          'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': null,
          'xl/pivotCache/pivotCacheDefinition1.xml': definition,
          'xl/pivotCache/pivotCacheRecords1.xml': null,
        }),
        { errorMode: 'strict', pivotCacheMode: 'records' },
      );
      expect(document.pivotCaches?.[0]).not.toHaveProperty('records');
    }
  });

  it('does not require cache-record relationships in metadata mode', async () => {
    const document = await parseXlsx(
      await bytes({
        'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': null,
        'xl/pivotCache/pivotCacheRecords1.xml': null,
      }),
      { errorMode: 'strict', pivotCacheMode: 'metadata' },
    );
    expect(document.pivotCaches?.[0]).not.toHaveProperty('records');
  });

  it('loads records when the definition omits its optional record count', async () => {
    const document = await parseXlsx(
      await bytes({
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'recordCount="2"',
          '',
        ),
      }),
      { errorMode: 'strict', pivotCacheMode: 'records' },
    );
    expect(document.pivotCaches?.[0]?.records).toHaveLength(2);
  });

  it('loads records only in explicit records mode and omits caches in none mode', async () => {
    const records = await parseXlsx(await bytes(), {
      errorMode: 'strict',
      pivotCacheMode: 'records',
    });
    expect(records.pivotCaches?.[0]?.records).toStrictEqual([
      [
        { index: 0, kind: 'shared-item' },
        { kind: 'number', value: 1 },
      ],
      [
        { index: 1, kind: 'shared-item' },
        { kind: 'number', value: 2 },
      ],
    ]);
    const none = await parseXlsx(
      await bytes({
        'xl/pivotCache/pivotCacheDefinition1.xml': '<broken',
        'xl/pivotCache/pivotCacheRecords1.xml': '<broken',
      }),
      { errorMode: 'strict', pivotCacheMode: 'none' },
    );
    expect(none).not.toHaveProperty('pivotCaches');
  });

  it('streams every pivot record value kind and skips malformed records in metadata mode', async () => {
    const definition = `<pivotCacheDefinition xmlns="${XLSX_SPREADSHEET_NS}" recordCount="1"><cacheSource type="scenario"/><cacheFields count="6"><cacheField name="Text"/><cacheField name="Bool"/><cacheField name="Date"/><cacheField name="Error"/><cacheField name="Blank"/><cacheField name="Number"/></cacheFields></pivotCacheDefinition>`;
    const records = `<pivotCacheRecords xmlns="${XLSX_SPREADSHEET_NS}" count="1"><r><s v="Text"/><b v="true"/><d v="2024-02-29T00:00:00Z"/><e v="#N/A"/><m/><n v="-0"/></r></pivotCacheRecords>`;
    const document = await parseXlsx(
      await bytes({
        'xl/pivotCache/pivotCacheDefinition1.xml': definition,
        'xl/pivotCache/pivotCacheRecords1.xml': records,
        'xl/worksheets/sheet1.xml': WORKSHEET_WITHOUT_PIVOT,
        'xl/worksheets/_rels/sheet1.xml.rels': null,
      }),
      { errorMode: 'strict', pivotCacheMode: 'records' },
    );
    expect(document.pivotCaches?.[0]?.records).toStrictEqual([
      [
        { kind: 'text', value: 'Text' },
        { kind: 'boolean', value: true },
        { kind: 'date', value: '2024-02-29T00:00:00Z' },
        { kind: 'error', value: '#N/A' },
        { kind: 'blank' },
        { kind: 'number', value: 0 },
      ],
    ]);

    await expect(
      parseXlsx(
        await bytes({
          'xl/pivotCache/pivotCacheRecords1.xml': '<broken',
        }),
        { errorMode: 'strict', pivotCacheMode: 'metadata' },
      ),
    ).resolves.toBeDefined();
  });

  it('preserves complete shared-item type metadata and values', async () => {
    const definition = `<pivotCacheDefinition xmlns="${XLSX_SPREADSHEET_NS}"><cacheSource type="scenario"/><cacheFields count="1"><cacheField name="Mixed" databaseField="0" serverField="1" uniqueList="0"><sharedItems count="6" containsBlank="1" containsDate="true" containsInteger="1" containsMixedTypes="true" containsNonDate="1" containsNumber="true" containsSemiMixedTypes="1" containsString="true" longText="1" minDate="2024-01-01T00:00:00Z" maxDate="2024-12-31T00:00:00Z" minValue="-1" maxValue="10"><s v="Text"/><n v="10"/><b v="1"/><d v="2024-01-01T00:00:00Z"/><e v="#N/A"/><m/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>`;
    const document = await parseXlsx(
      await bytes({
        'xl/pivotCache/pivotCacheDefinition1.xml': definition,
        'xl/worksheets/sheet1.xml': WORKSHEET_WITHOUT_PIVOT,
        'xl/worksheets/_rels/sheet1.xml.rels': null,
      }),
      { errorMode: 'strict' },
    );
    expect(document.pivotCaches?.[0]?.fields[0]).toStrictEqual({
      databaseField: false,
      items: [
        { kind: 'text', value: 'Text' },
        { kind: 'number', value: 10 },
        { kind: 'boolean', value: true },
        { kind: 'date', value: '2024-01-01T00:00:00Z' },
        { kind: 'error', value: '#N/A' },
        { kind: 'blank' },
      ],
      name: 'Mixed',
      serverField: true,
      sharedItems: {
        containsBlank: true,
        containsDate: true,
        containsInteger: true,
        containsMixedTypes: true,
        containsNonDate: true,
        containsNumber: true,
        containsSemiMixedTypes: true,
        containsString: true,
        longText: true,
        maximumDate: '2024-12-31T00:00:00Z',
        maximumNumber: 10,
        minimumDate: '2024-01-01T00:00:00Z',
        minimumNumber: -1,
      },
      uniqueList: false,
    });
  });

  it('ignores the defined shared-items extension container', async () => {
    const definition = cacheDefinition().replace(
      '<s v="B"/></sharedItems>',
      '<s v="B"/><extLst><ext uri="urn:test"/></extLst></sharedItems>',
    );
    const document = await parseXlsx(
      await bytes({
        'xl/pivotCache/pivotCacheDefinition1.xml': definition,
        'xl/worksheets/sheet1.xml': WORKSHEET_WITHOUT_PIVOT,
        'xl/worksheets/_rels/sheet1.xml.rels': null,
      }),
      { errorMode: 'strict' },
    );
    expect(document.pivotCaches?.[0]?.fields[0]?.items).toHaveLength(2);
  });

  it('preserves false cache booleans from records', async () => {
    const definition = `<pivotCacheDefinition xmlns="${XLSX_SPREADSHEET_NS}" recordCount="2"><cacheSource type="scenario"/><cacheFields count="1"><cacheField name="Flag"/></cacheFields></pivotCacheDefinition>`;
    const records = `<pivotCacheRecords xmlns="${XLSX_SPREADSHEET_NS}" count="2"><r><b v="false"/></r><r><b v="0"/></r></pivotCacheRecords>`;
    const document = await parseXlsx(
      await bytes({
        'xl/pivotCache/pivotCacheDefinition1.xml': definition,
        'xl/pivotCache/pivotCacheRecords1.xml': records,
        'xl/worksheets/sheet1.xml': WORKSHEET_WITHOUT_PIVOT,
        'xl/worksheets/_rels/sheet1.xml.rels': null,
      }),
      { errorMode: 'strict', pivotCacheMode: 'records' },
    );
    expect(document.pivotCaches?.[0]?.records).toStrictEqual([
      [{ kind: 'boolean', value: false }],
      [{ kind: 'boolean', value: false }],
    ]);
  });

  it('round-trips cache metadata through portable exact R0', async () => {
    const source = await bytes();
    const snapshot = await readXlsxRoundTrip(source);
    const output = await writeXlsxRoundTrip(
      await validateXlsxRoundTripJson(
        JSON.parse(JSON.stringify(snapshot)) as unknown,
      ),
    );
    expect(output.data).toEqual(source);
    expect(output.report.level).toBe('R0');
  });

  it('enforces pivot-record limits at exact boundaries', async () => {
    await expect(
      parseXlsx(await bytes(), {
        errorMode: 'strict',
        limits: { maxPivotRecords: 2 },
        pivotCacheMode: 'records',
      }),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(
          {},
          {
            errorMode: 'strict',
            limits: { maxPivotRecords: 1 },
            pivotCacheMode: 'records',
          },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxPivotRecords',
    });
  });

  it('recovers malformed optional cache metadata in tolerant mode', async () => {
    const result = await parseXlsxWithDiagnostics(
      await bytes({ 'xl/pivotCache/pivotCacheDefinition1.xml': '<broken' }),
    );
    expect(result.document).not.toHaveProperty('pivotCaches');
    expect(result.diagnostics).toMatchObject([
      { part: 'xl/pivotCache/pivotCacheDefinition1.xml', severity: 'warning' },
      {
        message: 'Pivot table cache reference is invalid',
        part: 'xl/pivotTables/pivotTable1.xml',
        severity: 'warning',
      },
    ]);
  });

  it.each([
    ['backgroundQuery', 'Pivot cache background-query flag is invalid'],
    ['enableRefresh', 'Pivot cache enable-refresh flag is invalid'],
    ['refreshOnLoad', 'Pivot cache refresh-on-load flag is invalid'],
    ['saveData', 'Pivot cache save-data flag is invalid'],
    ['supportAdvancedDrill', 'Pivot cache advanced-drill flag is invalid'],
    ['tupleCache', 'Pivot cache tuple flag is invalid'],
    ['upgradeOnRefresh', 'Pivot cache upgrade-on-refresh flag is invalid'],
  ] as const)('rejects invalid pivot cache flag %s', async (name, message) => {
    const xml = cacheDefinition().replace(
      new RegExp(`${name}="[^"]*"`, 'u'),
      `${name}="bad"`,
    );
    expect(
      (await capture({ 'xl/pivotCache/pivotCacheDefinition1.xml': xml }))
        .diagnostic.message,
    ).toBe(message);
  });

  it('reports exact cache record-count and axis-field lexical errors', async () => {
    expect(
      (
        await capture({
          'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
            'recordCount="2"',
            'recordCount="bad"',
          ),
        })
      ).diagnostic.message,
    ).toBe('Pivot cache record count is invalid');

    expect(
      (
        await capture({
          'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
            '<field x="-2"/>',
            '<field x="bad"/>',
          ),
        })
      ).diagnostic.message,
    ).toBe('Pivot column fields field index is invalid');

    expect(
      (
        await capture({
          'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
            /<colFields[\s\S]*?<\/colFields>/u,
            '<colFields>bad</colFields>',
          ),
        })
      ).diagnostic.message,
    ).toBe('Pivot column fields are invalid');
  });

  it.each([
    ['containsBlank', 'Pivot shared-items blank flag is invalid'],
    ['containsDate', 'Pivot shared-items date flag is invalid'],
    ['containsInteger', 'Pivot shared-items integer flag is invalid'],
    ['containsMixedTypes', 'Pivot shared-items mixed flag is invalid'],
    ['containsNonDate', 'Pivot shared-items non-date flag is invalid'],
    ['containsNumber', 'Pivot shared-items number flag is invalid'],
    ['containsSemiMixedTypes', 'Pivot shared-items semi-mixed flag is invalid'],
    ['containsString', 'Pivot shared-items string flag is invalid'],
    ['longText', 'Pivot shared-items long-text flag is invalid'],
  ] as const)(
    'rejects invalid pivot shared-item flag %s',
    async (name, message) => {
      const definition = `<pivotCacheDefinition xmlns="${XLSX_SPREADSHEET_NS}"><cacheSource type="scenario"/><cacheFields count="1"><cacheField name="F"><sharedItems ${name}="bad"/></cacheField></cacheFields></pivotCacheDefinition>`;
      expect(
        (
          await capture({
            'xl/pivotCache/pivotCacheDefinition1.xml': definition,
            'xl/worksheets/sheet1.xml': WORKSHEET_WITHOUT_PIVOT,
            'xl/worksheets/_rels/sheet1.xml.rels': null,
          })
        ).diagnostic.message,
      ).toBe(message);
    },
  );

  it.each([
    ['compact', 'Pivot table compact flag is invalid'],
    ['outline', 'Pivot table outline flag is invalid'],
    ['colGrandTotals', 'Pivot table column-grand-total flag is invalid'],
    ['showHeaders', 'Pivot table header flag is invalid'],
    ['rowGrandTotals', 'Pivot table row-grand-total flag is invalid'],
  ] as const)('rejects invalid pivot table flag %s', async (name, message) => {
    const xml = pivotTable().replace(
      new RegExp(`${name}="[^"]*"`, 'u'),
      `${name}="bad"`,
    );
    expect(
      (await capture({ 'xl/pivotTables/pivotTable1.xml': xml })).diagnostic
        .message,
    ).toBe(message);
  });

  it.each([
    ['compact', 'Pivot field compact flag is invalid'],
    ['dataField', 'Pivot field data flag is invalid'],
    ['outline', 'Pivot field outline flag is invalid'],
    ['showAll', 'Pivot field show-all flag is invalid'],
    ['subtotalTop', 'Pivot field subtotal position is invalid'],
  ] as const)('rejects invalid pivot field flag %s', async (name, message) => {
    const source = pivotTable();
    const xml = source.replace(
      name === 'dataField'
        ? 'dataField="1"'
        : `${name}="${
            name === 'compact'
              ? '1'
              : name === 'outline' || name === 'showAll'
                ? '0'
                : 'false'
          }"`,
      `${name}="bad"`,
    );
    expect(
      (await capture({ 'xl/pivotTables/pivotTable1.xml': xml })).diagnostic
        .message,
    ).toBe(message);
  });

  it.each([
    ['showColHeaders', 'Pivot style column-header flag is invalid'],
    ['showColStripes', 'Pivot style column-stripe flag is invalid'],
    ['showLastColumn', 'Pivot style last-column flag is invalid'],
    ['showRowHeaders', 'Pivot style row-header flag is invalid'],
    ['showRowStripes', 'Pivot style row-stripe flag is invalid'],
  ] as const)('rejects invalid pivot style flag %s', async (name, message) => {
    const xml = pivotTable().replace(
      new RegExp(`${name}="[^"]*"`, 'u'),
      `${name}="bad"`,
    );
    expect(
      (await capture({ 'xl/pivotTables/pivotTable1.xml': xml })).diagnostic
        .message,
    ).toBe(message);
  });

  it.each([
    [
      { 'xl/workbook.xml': WORKBOOK.replace('cacheId="7"', 'cacheId="bad"') },
      'Workbook pivot cache ID is invalid',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELS.replace(
          CACHE_DEFINITION_RELATIONSHIP,
          `${XLSX_OFFICE_REL_TYPE}image`,
        ),
      },
      'Workbook pivot cache relationship is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'count="2"',
          'count="1"',
        ),
      },
      'Pivot cache field count does not match',
    ],
    [
      {
        'xl/pivotCache/pivotCacheRecords1.xml': cacheRecords().replace(
          'count="2"',
          'count="1"',
        ),
      },
      'Pivot cache-record count does not match',
    ],
    [
      {
        'xl/pivotCache/pivotCacheRecords1.xml': cacheRecords().replace(
          '<x v="1"/>',
          '<x v="2"/>',
        ),
      },
      'Pivot record shared-item reference is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': `<wrong xmlns="${XLSX_SPREADSHEET_NS}"/>`,
      },
      'Pivot pivotCacheDefinition root is missing',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinitionWithSource(
          '<cacheSource type="bad"/>',
        ),
      },
      'Pivot cache source type is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinitionWithSource(
          '<cacheSource type="worksheet"/>',
        ),
      },
      'Pivot worksheet source is missing',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinitionWithSource(
          '<cacheSource type="worksheet"><worksheetSource ref="bad"/></cacheSource>',
        ),
      },
      'Pivot worksheet source range is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinitionWithSource(
          '<cacheSource type="worksheet"><worksheetSource/></cacheSource>',
        ),
      },
      'Pivot worksheet source requires a range or name',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinitionWithSource(
          '<cacheSource type="external" connectionId="bad"/>',
        ),
      },
      'Pivot cache connection ID is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          '<cacheFields count="2">',
          '<cacheFields count="bad">',
        ),
      },
      'Pivot cache field count is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'name="Category"',
          'name=""',
        ),
      },
      'Pivot cache field name is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          '<sharedItems count="2" containsString="1" containsBlank="0">',
          '<sharedItems count="1" containsString="1" containsBlank="0">',
        ),
      },
      'Pivot shared-item count does not match',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          '<s v="A"/>',
          '<s/>',
        ),
      },
      'Pivot cache item value is missing',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          '<n v="1"/>',
          '<n v="bad"/>',
        ),
      },
      'Pivot cache number is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'containsString="1"',
          'containsString="bad"',
        ),
      },
      'Pivot shared-items string flag is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'databaseField="1"',
          'databaseField="bad"',
        ),
      },
      'Pivot cache database-field flag is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'serverField="0"',
          'serverField="bad"',
        ),
      },
      'Pivot cache server-field flag is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'uniqueList="true"',
          'uniqueList="bad"',
        ),
      },
      'Pivot cache unique-list flag is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheRecords1.xml': `<wrong xmlns="${XLSX_SPREADSHEET_NS}"/>`,
      },
      'Pivot cache-record root is missing',
    ],
    [
      {
        'xl/pivotCache/pivotCacheRecords1.xml': cacheRecords().replace(
          XLSX_SPREADSHEET_NS,
          'urn:wrong',
        ),
      },
      'Pivot cache-record element has the wrong namespace',
    ],
    [
      {
        'xl/pivotCache/pivotCacheRecords1.xml': cacheRecords().replace(
          '<r><x v="0"/>',
          '<wrong/><r><x v="0"/>',
        ),
      },
      'Pivot cache-record structure is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheRecords1.xml': cacheRecords().replace(
          '<r><x v="0"/>',
          '<r>bad<x v="0"/>',
        ),
      },
      'Pivot cache-record text is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          /<cacheSource[\s\S]*?<\/cacheSource>/u,
          '',
        ),
      },
      'Pivot cache source is missing',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          /<cacheFields[\s\S]*<\/cacheFields>/u,
          '',
        ),
      },
      'Pivot cache fields are missing',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          /<cacheFields[\s\S]*<\/cacheFields>/u,
          '<cacheFields count="1"><cacheField>bad</cacheField></cacheFields>',
        ),
      },
      'Pivot cache fields are invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          '<sharedItems count="2" containsString="1" containsBlank="0">',
          '<sharedItems count="bad" containsString="1" containsBlank="0">',
        ),
      },
      'Pivot shared-item count is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          '<sharedItems count="2" containsString="1" containsBlank="0">',
          '<sharedItems>bad</sharedItems><sharedItems count="2" containsString="1" containsBlank="0">',
        ),
      },
      'Pivot shared items are invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          '<s v="A"/>',
          '<s>bad</s>',
        ),
      },
      'Pivot shared-item collection is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          '<s v="A"/>',
          '<b v="bad"/>',
        ),
      },
      'Pivot cache boolean is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'minValue="1"',
          'minValue="bad"',
        ),
      },
      'Pivot shared-item minimum is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'maxValue="2"',
          'maxValue="bad"',
        ),
      },
      'Pivot shared-item maximum is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          '<s v="A"/>',
          '<d v="2023-02-29T00:00:00Z"/>',
        ),
      },
      'Pivot cache date is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          '<s v="A"/>',
          '<z v="A"/>',
        ),
      },
      'Pivot cache item type is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'minValue="1"',
          'minValue="1" minDate="bad"',
        ),
      },
      'Pivot shared-item minimum date is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'maxValue="2"',
          'maxValue="2" maxDate="bad"',
        ),
      },
      'Pivot shared-item maximum date is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          XLSX_SPREADSHEET_NS,
          'urn:wrong',
        ),
      },
      'Pivot pivotCacheDefinition root has the wrong namespace',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
          'application/xml',
        ),
      },
      'Workbook pivot cache target has the wrong content type',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
          'application/xml',
        ),
      },
      'Pivot cache-record target has the wrong content type',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
          'application/xml',
        ),
      },
      'Pivot table has the wrong content type',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'missingItemsLimit="10"',
          'missingItemsLimit="bad"',
        ),
      },
      'Pivot cache missing-items limit is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'refreshedDate="45292.5"',
          'refreshedDate="bad"',
        ),
      },
      'Pivot cache refreshed date is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheDefinition1.xml': cacheDefinition().replace(
          'recordCount="2"',
          'recordCount="3"',
        ),
      },
      'Pivot cache definition record count does not match',
    ],
    [
      {
        'xl/pivotCache/pivotCacheRecords1.xml': cacheRecords().replace(
          'count="2"',
          'count="bad"',
        ),
      },
      'Pivot cache-record count is invalid',
    ],
    [
      {
        'xl/pivotCache/pivotCacheRecords1.xml': cacheRecords().replace(
          '<r><x v="0"/><n v="1"/></r>',
          '<r><x v="0"/></r>',
        ),
      },
      'Pivot cache-record field count does not match',
    ],
    [
      {
        'xl/pivotCache/pivotCacheRecords1.xml': cacheRecords().replace(
          '<x v="0"/>',
          '<x v="bad"/>',
        ),
      },
      'Pivot record shared-item index is invalid',
    ],
    [
      {
        'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"/>`,
      },
      'Pivot cache records are missing',
    ],
    [
      {
        'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels':
          CACHE_RELS.replace(
            'Target="pivotCacheRecords1.xml"',
            'Target="https://example.invalid/records.xml" TargetMode="External"',
          ),
      },
      'Pivot cache-record relationship is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': `<wrong xmlns="${XLSX_SPREADSHEET_NS}"/>`,
      },
      'Pivot pivotTableDefinition root is missing',
    ],
    [
      {
        'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replace(
          PIVOT_TABLE_RELATIONSHIP,
          `${XLSX_OFFICE_REL_TYPE}image`,
        ),
      },
      'Worksheet pivot-table relationship is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'cacheId="7"',
          'cacheId="8"',
        ),
      },
      'Pivot table cache reference is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'ref="D1:H10"',
          'ref="bad"',
        ),
      },
      'Pivot table location is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'axis="axisRow"',
          'axis="bad"',
        ),
      },
      'Pivot field axis is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'sortType="ascending"',
          'sortType="bad"',
        ),
      },
      'Pivot field sort type is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          't="blank"',
          't="bad"',
        ),
      },
      'Pivot field item type is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<item x="1" t="blank"/>',
          '<item x="2" t="blank"/>',
        ),
      },
      'Pivot field shared-item reference is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<field x="0"/>',
          '<field x="2"/>',
        ),
      },
      'Pivot axis field reference is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'fld="1" subtotal',
          'fld="2" subtotal',
        ),
      },
      'Pivot data-field reference is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<pageField fld="0"',
          '<pageField fld="2"',
        ),
      },
      'Pivot page-field reference is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<filter fld="0"',
          '<filter fld="2"',
        ),
      },
      'Pivot filter field reference is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'name="SalesPivot"',
          'name=""',
        ),
      },
      'Pivot table name or data caption is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'cacheId="7"',
          'cacheId="bad"',
        ),
      },
      'Pivot table cache ID is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<pivotFields count="2">',
          '<pivotFields count="bad">',
        ),
      },
      'Pivot field count is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<pivotFields count="2">',
          '<pivotFields count="1">',
        ),
      },
      'Pivot field count does not match',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable()
          .replace('<pivotFields count="2">', '<pivotFields count="1">')
          .replace('<pivotField axis="axisValues" dataField="1"/>', ''),
      },
      'Pivot table field count does not match its cache',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<items count="2">',
          '<items count="1">',
        ),
      },
      'Pivot field item count does not match',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<items count="2">',
          '<items count="bad">',
        ),
      },
      'Pivot field item count is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<dataField name="Sum Amount" fld="1" subtotal="sum"',
          '<dataField name="Sum Amount" fld="1" subtotal="bad"',
        ),
      },
      'Pivot subtotal is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'showDataAs="percentOfTotal"',
          'showDataAs="bad"',
        ),
      },
      'Pivot display mode is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<field x="-2"/>',
          '<field x="-3"/>',
        ),
      },
      'Pivot column fields field index is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<rowFields count="1"><field x="0"/></rowFields>',
          '<rowFields count="1"><field x="-3"/></rowFields>',
        ),
      },
      'Pivot row fields field index is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<rowFields count="1">',
          '<rowFields count="2">',
        ),
      },
      'Pivot row fields count does not match',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<dataFields count="1">',
          '<dataFields count="2">',
        ),
      },
      'Pivot data-field count does not match',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<dataFields count="1">',
          '<dataFields count="bad">',
        ),
      },
      'Pivot data-field count is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          /<dataFields[\s\S]*?<\/dataFields>/u,
          '<dataFields>bad</dataFields>',
        ),
      },
      'Pivot data fields are invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'baseField="0"',
          'baseField="bad"',
        ),
      },
      'Pivot data-field base field is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'baseItem="0"',
          'baseItem="bad"',
        ),
      },
      'Pivot data-field base item is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<pageField fld="0"',
          '<pageField fld="bad"',
        ),
      },
      'Pivot page-field field is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'hier="0"',
          'hier="bad"',
        ),
      },
      'Pivot page-field hierarchy is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          /<pageFields[\s\S]*?<\/pageFields>/u,
          '<pageFields>bad</pageFields>',
        ),
      },
      'Pivot page fields are invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<pageFields count="1">',
          '<pageFields count="bad">',
        ),
      },
      'Pivot page-field count is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<pageFields count="1">',
          '<pageFields count="2">',
        ),
      },
      'Pivot page-field count does not match',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<filter fld="0" type="captionEqual"',
          '<filter fld="0" type=""',
        ),
      },
      'Pivot filter type is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'evalOrder="-1"',
          'evalOrder="bad"',
        ),
      },
      'Pivot filter evaluation order is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<filter fld="0"',
          '<filter fld="bad"',
        ),
      },
      'Pivot filter field is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'id="1"',
          'id="bad"',
        ),
      },
      'Pivot filter ID is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          /<filters[\s\S]*?<\/filters>/u,
          '<filters>bad</filters>',
        ),
      },
      'Pivot filters are invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<filters count="1">',
          '<filters count="bad">',
        ),
      },
      'Pivot filter count is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<filters count="1">',
          '<filters count="2">',
        ),
      },
      'Pivot filter count does not match',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          XLSX_SPREADSHEET_NS,
          'urn:wrong',
        ),
      },
      'Pivot pivotTableDefinition root has the wrong namespace',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<item x="0"/>',
          '<item x="bad"/>',
        ),
      },
      'Pivot field shared-item index is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          /<items[\s\S]*?<\/items>/u,
          '<items>bad</items>',
        ),
      },
      'Pivot field items are invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          /<pivotFields[\s\S]*?<\/pivotFields>/u,
          '',
        ),
      },
      'Pivot fields are missing',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          /<pivotFields[\s\S]*?<\/pivotFields>/u,
          '<pivotFields>bad</pivotFields>',
        ),
      },
      'Pivot fields are invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<colFields count="1">',
          '<colFields count="bad">',
        ),
      },
      'Pivot column fields count is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<colFields count="1">',
          '<colFields count="2">',
        ),
      },
      'Pivot column fields count does not match',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          '<dataField name="Sum Amount" fld="1"',
          '<dataField name="Sum Amount" fld="bad"',
        ),
      },
      'Pivot data-field field is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'item="1"',
          'item="bad"',
        ),
      },
      'Pivot page-field item is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'stringValue1="A"',
          'iMeasureFld="bad" stringValue1="A"',
        ),
      },
      'Pivot filter measure field is invalid',
    ],
    [
      {
        'xl/pivotTables/pivotTable1.xml': pivotTable().replace(
          'stringValue1="A"',
          'iMeasureHier="bad" stringValue1="A"',
        ),
      },
      'Pivot filter measure hierarchy is invalid',
    ],
    [
      { 'xl/pivotCache/pivotCacheRecords1.xml': null },
      'Required XLSX part is missing: xl/pivotCache/pivotCacheRecords1.xml',
    ],
    [
      {
        'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replace(
          'Target="../pivotTables/pivotTable1.xml"',
          'Target="https://example.invalid/pivot.xml" TargetMode="External"',
        ),
      },
      'Worksheet pivot-table relationship is invalid',
    ],
  ] as const)(
    'rejects invalid pivot cache contract %#',
    async (changes, message) => {
      try {
        expect(
          (
            await capture(changes, {
              errorMode: 'strict',
              pivotCacheMode: 'records',
            })
          ).diagnostic.message,
          JSON.stringify(changes),
        ).toBe(message);
      } catch (error) {
        throw new Error(
          `Pivot invalid case expected ${message}: ${JSON.stringify(changes)}`,
          { cause: error },
        );
      }
    },
  );
});
