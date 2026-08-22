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
  parseXlsxSlicerBoolean,
  parseXlsxSlicerCrossFilter,
  parseXlsxSlicerSortOrder,
  parseXlsxSlicerUnsignedInteger,
} from '../../src/formats/xlsx/internal/slicer';
import {
  createIndependentXlsx,
  type XlsxBlackBoxOverrides,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const X14_NAMESPACE =
  'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const X15_NAMESPACE =
  'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
const SLICER_RELATIONSHIP =
  'http://schemas.microsoft.com/office/2007/relationships/slicer';
const SLICER_CACHE_RELATIONSHIP =
  'http://schemas.microsoft.com/office/2007/relationships/slicerCache';
const TIMELINE_RELATIONSHIP =
  'http://schemas.microsoft.com/office/2011/relationships/timeline';
const TIMELINE_CACHE_RELATIONSHIP =
  'http://schemas.microsoft.com/office/2011/relationships/timelineCache';

const CONTENT_TYPES = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.bin" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
  <Override PartName="/xl/timelineCaches/timelineCache1.bin" ContentType="application/vnd.ms-excel.timelineCache+xml"/>
  <Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>
  <Override PartName="/xl/timelines/timeline1.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>
</Types>`;

const WORKBOOK = `<workbook xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="sheet"/></sheets></workbook>`;
const WORKBOOK_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="slicerCache" Type="${SLICER_CACHE_RELATIONSHIP}" Target="slicerCaches/slicerCache1.bin"/><Relationship Id="timelineCache" Type="${TIMELINE_CACHE_RELATIONSHIP}" Target="timelineCaches/timelineCache1.bin"/></Relationships>`;
const WORKSHEET = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/></worksheet>`;
const WORKSHEET_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="slicer" Type="${SLICER_RELATIONSHIP}" Target="../slicers/slicer1.xml"/><Relationship Id="timeline" Type="${TIMELINE_RELATIONSHIP}" Target="../timelines/timeline1.xml"/></Relationships>`;

const SLICER_CACHE = `<x14:slicerCacheDefinition xmlns:x14="${X14_NAMESPACE}" name="Cache_Country" sourceName="Country"><x14:pivotTables><x14:pivotTable tabId="1" name="SalesPivot"/></x14:pivotTables><x14:data><x14:tabular pivotCacheId="7" sortOrder="ascending" crossFilter="showItemsWithDataAtTop"/></x14:data></x14:slicerCacheDefinition>`;
const TIMELINE_CACHE = `<x15:timelineCacheDefinition xmlns:x15="${X15_NAMESPACE}" name="Cache_Date" sourceName="OrderDate"><x15:pivotTables><x15:pivotTable tabId="1" name="SalesPivot"/></x15:pivotTables><x15:timelineState pivotCacheId="7"/></x15:timelineCacheDefinition>`;
const TABLE_SLICER_CACHE = SLICER_CACHE.replace(
  /<x14:pivotTables>[\s\S]*?<\/x14:data>/u,
  `<x14:extLst><x14:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}"><x15:tableSlicerCache xmlns:x15="${X15_NAMESPACE}" tableId="2" column="3" sortOrder="descending" customListSort="0" crossFilter="none"/></x14:ext></x14:extLst>`,
);
const SLICER = `<x14:slicers xmlns:x14="${X14_NAMESPACE}"><x14:slicer name="Country Slicer" cache="Cache_Country" caption="Country" startItem="2" columnCount="3" showCaption="0" level="1" style="SlicerStyleLight2" lockedPosition="1" rowHeight="240000"/></x14:slicers>`;
const TIMELINE = `<x15:timelines xmlns:x15="${X15_NAMESPACE}"><x15:timeline name="Date Timeline" cache="Cache_Date" caption="Order Date" showHeader="0" showSelectionLabel="1" showTimeLevel="0" showHorizontalScrollbar="1" level="2" selectionLevel="1" scrollPosition="2024-01-01T00:00:00Z" style="TimeSlicerStyleLight2"/></x15:timelines>`;

function parts(changes: XlsxBlackBoxOverrides = {}): XlsxBlackBoxOverrides {
  return {
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
    'xl/sharedStrings.xml': null,
    'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE,
    'xl/slicers/slicer1.xml': SLICER,
    'xl/styles.xml': null,
    'xl/timelineCaches/timelineCache1.bin': TIMELINE_CACHE,
    'xl/timelines/timeline1.xml': TIMELINE,
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
): Promise<XlsxParseError> {
  try {
    await parseXlsx(await bytes(changes), { errorMode: 'strict' });
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected XLSX slicer parsing to fail');
}

describe('XLSX slicers and timelines', () => {
  it('normalizes slicer lexical helpers exactly', () => {
    expect(
      parseXlsxSlicerUnsignedInteger('0', undefined, 'integer', 'part.xml'),
    ).toBe(0);
    expect(
      parseXlsxSlicerUnsignedInteger(
        '4294967295',
        undefined,
        'integer',
        'part.xml',
      ),
    ).toBe(0xffff_ffff);
    expect(
      parseXlsxSlicerUnsignedInteger(undefined, 3, 'integer', 'part.xml'),
    ).toBe(3);
    for (const value of ['-1', '01', '1.0', '1 ', '1x', '4294967296']) {
      expect(() =>
        parseXlsxSlicerUnsignedInteger(value, undefined, 'integer', 'part.xml'),
      ).toThrow('integer');
    }

    expect(parseXlsxSlicerBoolean('0', true, 'boolean', 'part.xml')).toBe(
      false,
    );
    expect(parseXlsxSlicerBoolean('false', true, 'boolean', 'part.xml')).toBe(
      false,
    );
    expect(parseXlsxSlicerBoolean('1', false, 'boolean', 'part.xml')).toBe(
      true,
    );
    expect(parseXlsxSlicerBoolean('true', false, 'boolean', 'part.xml')).toBe(
      true,
    );
    expect(parseXlsxSlicerBoolean(undefined, true, 'boolean', 'part.xml')).toBe(
      true,
    );
    expect(() =>
      parseXlsxSlicerBoolean('bad', false, 'boolean', 'part.xml'),
    ).toThrow('boolean');
  });

  it('normalizes every slicer sort and cross-filter mode', () => {
    for (const value of [undefined, 'ascending', 'descending'] as const) {
      expect(parseXlsxSlicerSortOrder(value, 'part.xml')).toBe(value);
    }
    for (const value of [
      undefined,
      'none',
      'showItemsWithDataAtTop',
      'showItemsWithNoData',
    ] as const) {
      expect(parseXlsxSlicerCrossFilter(value, 'part.xml')).toBe(value);
    }
    expect(() => parseXlsxSlicerSortOrder('bad', 'part.xml')).toThrow(
      'Slicer sort order is invalid',
    );
    expect(() => parseXlsxSlicerCrossFilter('bad', 'part.xml')).toThrow(
      'Slicer cross-filter is invalid',
    );
  });

  it('parses cache ownership and safe display metadata', async () => {
    const document = await parseXlsx(await bytes(), { errorMode: 'strict' });
    expect(document.slicerCaches).toStrictEqual([
      {
        crossFilter: 'showItemsWithDataAtTop',
        customListSort: true,
        index: 0,
        name: 'Cache_Country',
        pivotCacheId: 7,
        pivotTables: [{ name: 'SalesPivot', sheetId: 1 }],
        sortOrder: 'ascending',
        showMissing: true,
        sourceKind: 'tabular',
        sourceName: 'Country',
      },
    ]);
    expect(document.timelineCaches).toStrictEqual([
      {
        index: 0,
        name: 'Cache_Date',
        pivotCacheId: 7,
        pivotTables: [{ name: 'SalesPivot', sheetId: 1 }],
        sourceName: 'OrderDate',
      },
    ]);
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    expect(
      sheet.kind === 'worksheet' ? sheet.slicers : undefined,
    ).toStrictEqual([
      {
        cacheIndex: 0,
        caption: 'Country',
        columnCount: 3,
        level: 1,
        lockedPosition: true,
        name: 'Country Slicer',
        rowHeight: 240000,
        selectionRelation: 'full-sheet',
        showCaption: false,
        startItem: 2,
        style: 'SlicerStyleLight2',
      },
    ]);
    expect(
      sheet.kind === 'worksheet' ? sheet.timelines : undefined,
    ).toStrictEqual([
      {
        cacheIndex: 0,
        caption: 'Order Date',
        level: 2,
        name: 'Date Timeline',
        scrollPosition: '2024-01-01T00:00:00Z',
        selectionLevel: 1,
        selectionRelation: 'full-sheet',
        showHeader: false,
        showHorizontalScrollbar: true,
        showSelectionLabel: true,
        showTimeLevel: false,
        style: 'TimeSlicerStyleLight2',
      },
    ]);
  });

  it('validates metadata but omits displays for range selections', async () => {
    const document = await parseXlsx(await bytes(), {
      errorMode: 'strict',
      selection: { ranges: { Sheet1: ['A1'] } },
    });
    expect(document.slicerCaches).toHaveLength(1);
    expect(document.timelineCaches).toHaveLength(1);
    const sheet = document.sheets[0]!;
    expect(sheet).not.toHaveProperty('slicers');
    expect(sheet).not.toHaveProperty('timelines');
  });

  it('preserves slicer metadata through portable exact R0', async () => {
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

  it('preserves defaults and optional omissions', async () => {
    const document = await parseXlsx(
      await bytes({
        'xl/slicers/slicer1.xml': `<x14:slicers xmlns:x14="${X14_NAMESPACE}"><x14:slicer name="Default Slicer" cache="Cache_Country" rowHeight="200000"/></x14:slicers>`,
        'xl/timelines/timeline1.xml': `<x15:timelines xmlns:x15="${X15_NAMESPACE}"><x15:timeline name="Default Timeline" cache="Cache_Date" level="0" selectionLevel="0"/></x15:timelines>`,
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.slicers?.[0] : undefined,
    ).toStrictEqual({
      cacheIndex: 0,
      columnCount: 1,
      lockedPosition: false,
      name: 'Default Slicer',
      rowHeight: 200000,
      selectionRelation: 'full-sheet',
      showCaption: true,
      startItem: 0,
    });
    expect(
      sheet.kind === 'worksheet' ? sheet.timelines?.[0] : undefined,
    ).toStrictEqual({
      cacheIndex: 0,
      level: 0,
      name: 'Default Timeline',
      selectionLevel: 0,
      selectionRelation: 'full-sheet',
      showHeader: true,
      showHorizontalScrollbar: true,
      showSelectionLabel: true,
      showTimeLevel: true,
    });
  });

  it('preserves cache enum alternatives and true optional omissions', async () => {
    const alternatives = await parseXlsx(
      await bytes({
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          'sortOrder="ascending" crossFilter="showItemsWithDataAtTop"',
          'sortOrder="descending" crossFilter="showItemsWithNoData"',
        ),
      }),
      { errorMode: 'strict' },
    );
    expect(alternatives.slicerCaches?.[0]).toMatchObject({
      crossFilter: 'showItemsWithNoData',
      sortOrder: 'descending',
    });

    const omitted = await parseXlsx(
      await bytes({
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          ' pivotCacheId="7" sortOrder="ascending" crossFilter="showItemsWithDataAtTop"',
          '',
        ),
      }),
      { errorMode: 'strict' },
    );
    expect(omitted.slicerCaches?.[0]).not.toHaveProperty('crossFilter');
    expect(omitted.slicerCaches?.[0]).not.toHaveProperty('sortOrder');
    expect(omitted.slicerCaches?.[0]).not.toHaveProperty('pivotCacheId');
  });

  it('parses unprefixed and reserved-prefix extension roots', async () => {
    for (const prefix of ['unprefixed', 'reserved'] as const) {
      const transform = (xml: string, sourcePrefix: 'x14' | 'x15'): string =>
        prefix === 'unprefixed'
          ? xml
              .replaceAll(`${sourcePrefix}:`, '')
              .replace(`xmlns:${sourcePrefix}=`, 'xmlns=')
          : xml
              .replaceAll(`${sourcePrefix}:`, 'p:')
              .replace(`xmlns:${sourcePrefix}=`, 'xmlns:p=');
      await expect(
        parseXlsx(
          await bytes({
            'xl/slicerCaches/slicerCache1.bin': transform(SLICER_CACHE, 'x14'),
            'xl/slicers/slicer1.xml': transform(SLICER, 'x14'),
            'xl/timelineCaches/timelineCache1.bin': transform(
              TIMELINE_CACHE,
              'x15',
            ),
            'xl/timelines/timeline1.xml': transform(TIMELINE, 'x15'),
          }),
          { errorMode: 'strict' },
        ),
      ).resolves.toBeDefined();
    }
  });

  it('accepts the maximum unsigned timeline level', async () => {
    const document = await parseXlsx(
      await bytes({
        'xl/timelines/timeline1.xml': TIMELINE.replace(
          'level="2"',
          'level="4294967295"',
        ),
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.timelines?.[0]?.level : undefined,
    ).toBe(0xffff_ffff);
  });

  it('finds a pivot cache ID after an earlier tabular extension node', async () => {
    const definition = SLICER_CACHE.replace(
      '<x14:tabular pivotCacheId="7" sortOrder="ascending" crossFilter="showItemsWithDataAtTop"/>',
      '<x14:tabular/><x14:tabular pivotCacheId="7"/>',
    );
    const document = await parseXlsx(
      await bytes({ 'xl/slicerCaches/slicerCache1.bin': definition }),
      { errorMode: 'strict' },
    );
    expect(document.slicerCaches?.[0]?.pivotCacheId).toBe(7);
  });

  it('parses table slicer cache ownership from the x15 extension', async () => {
    const document = await parseXlsx(
      await bytes({ 'xl/slicerCaches/slicerCache1.bin': TABLE_SLICER_CACHE }),
      { errorMode: 'strict' },
    );
    expect(document.slicerCaches?.[0]).toStrictEqual({
      crossFilter: 'none',
      customListSort: false,
      index: 0,
      name: 'Cache_Country',
      pivotTables: [],
      sortOrder: 'descending',
      sourceKind: 'table',
      sourceName: 'Country',
      table: { column: 3, id: 2 },
    });
  });

  it('parses OLAP slicer cache ownership without tabular defaults', async () => {
    const definition = SLICER_CACHE.replace(
      /<x14:data>[\s\S]*?<\/x14:data>/u,
      '<x14:data><x14:olap pivotCacheId="9"/></x14:data>',
    );
    const document = await parseXlsx(
      await bytes({ 'xl/slicerCaches/slicerCache1.bin': definition }),
      { errorMode: 'strict' },
    );
    expect(document.slicerCaches?.[0]).toStrictEqual({
      index: 0,
      name: 'Cache_Country',
      pivotCacheId: 9,
      pivotTables: [{ name: 'SalesPivot', sheetId: 1 }],
      sourceKind: 'olap',
      sourceName: 'Country',
    });
  });

  it('reads timeline pivot-cache ownership from a pivot-cache definition', async () => {
    const definition = TIMELINE_CACHE.replace(
      '<x15:timelineState pivotCacheId="7"/>',
      '<x15:timelinePivotCacheDefinition pivotCacheId="11"/>',
    );
    const document = await parseXlsx(
      await bytes({ 'xl/timelineCaches/timelineCache1.bin': definition }),
      { errorMode: 'strict' },
    );
    expect(document.timelineCaches?.[0]?.pivotCacheId).toBe(11);
  });

  it('enforces the shared visual-object limit exactly', async () => {
    await expect(
      parseXlsx(await bytes(), {
        errorMode: 'strict',
        limits: { maxDrawings: 2 },
      }),
    ).resolves.toBeDefined();
    expect(
      (
        await (async () => {
          try {
            await parseXlsx(await bytes(), {
              errorMode: 'strict',
              limits: { maxDrawings: 1 },
            });
          } catch (error) {
            return error as XlsxParseError;
          }
          throw new Error('Expected slicer resource limit failure');
        })()
      ).diagnostic,
    ).toMatchObject({ actual: 2, limit: 1, limitName: 'maxDrawings' });
  });

  it('maps slicer-cache text limits to a structured fatal error', async () => {
    try {
      await parseXlsx(await bytes(), {
        errorMode: 'strict',
        limits: { maxTextCharacters: 1 },
      });
    } catch (error) {
      expect(error).toBeInstanceOf(XlsxParseError);
      expect((error as XlsxParseError).diagnostic).toMatchObject({
        actual: 'Cache_Country'.length,
        limit: 1,
        limitName: 'maxTextCharacters',
      });
      return;
    }
    throw new Error('Expected slicer cache text limit failure');
  });

  it('recovers malformed optional slicer metadata in tolerant mode', async () => {
    const result = await parseXlsxWithDiagnostics(
      await bytes({ 'xl/slicerCaches/slicerCache1.bin': '<broken' }),
    );
    expect(result.document).not.toHaveProperty('slicerCaches');
    expect(result.diagnostics.some((item) => item.severity === 'warning')).toBe(
      true,
    );
  });

  it.each([
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'application/vnd.ms-excel.slicerCache+xml',
          'application/xml',
        ),
      },
      'XLSX slicer or timeline target has the wrong content type',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELS.replace(
          'Target="slicerCaches/slicerCache1.bin"',
          'Target="https://example.invalid/cache.bin" TargetMode="External"',
        ),
      },
      'XLSX slicer or timeline relationship must be internal',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': `<wrong xmlns="${X14_NAMESPACE}"/>`,
      },
      'XLSX slicerCacheDefinition root is missing',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          X14_NAMESPACE,
          'urn:wrong',
        ),
      },
      'XLSX slicerCacheDefinition root has the wrong namespace',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          'name="Cache_Country"',
          'name=""',
        ),
      },
      'Slicer or timeline cache name is invalid',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          'sourceName="Country"',
          'sourceName=""',
        ),
      },
      'Slicer or timeline cache source name is invalid',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          'tabId="1"',
          'tabId="bad"',
        ),
      },
      'Slicer pivot-table sheet ID is invalid',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          'name="SalesPivot"',
          'name=""',
        ),
      },
      'Slicer pivot-table owner is invalid',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          'pivotCacheId="7"',
          'pivotCacheId="bad"',
        ),
      },
      'Slicer pivot cache ID is invalid',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          /<x14:data>[\s\S]*?<\/x14:data>/u,
          '',
        ),
      },
      'Slicer cache source metadata is missing',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': TABLE_SLICER_CACHE.replace(
          'tableId="2"',
          'tableId="bad"',
        ),
      },
      'Table slicer table ID is invalid',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': TABLE_SLICER_CACHE.replace(
          'column="3"',
          'column="bad"',
        ),
      },
      'Table slicer column is invalid',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': TABLE_SLICER_CACHE.replace(
          ' tableId="2"',
          '',
        ),
      },
      'Table slicer owner is invalid',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': TABLE_SLICER_CACHE.replace(
          ' column="3"',
          '',
        ),
      },
      'Table slicer owner is invalid',
    ],
    [
      {
        'xl/timelineCaches/timelineCache1.bin': TIMELINE_CACHE.replace(
          'name="Cache_Date"',
          'name="Cache_Country"',
        ),
      },
      'Workbook contains duplicate slicer or timeline cache names',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          'sortOrder="ascending"',
          'sortOrder="bad"',
        ),
      },
      'Slicer sort order is invalid',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          'crossFilter="showItemsWithDataAtTop"',
          'crossFilter="bad"',
        ),
      },
      'Slicer cross-filter is invalid',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          'sortOrder="ascending"',
          'customListSort="bad" sortOrder="ascending"',
        ),
      },
      'Slicer custom-list-sort flag is invalid',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          'sortOrder="ascending"',
          'showMissing="bad" sortOrder="ascending"',
        ),
      },
      'Slicer show-missing flag is invalid',
    ],
    [
      {
        'xl/slicers/slicer1.xml': SLICER.replace(
          'cache="Cache_Country"',
          'cache="Missing"',
        ),
      },
      'Slicer or timeline cache reference is invalid',
    ],
    [
      {
        'xl/timelines/timeline1.xml': TIMELINE.replace(
          'cache="Cache_Date"',
          'cache="Cache_Country"',
        ),
      },
      'Slicer or timeline cache reference is invalid',
    ],
    [
      {
        'xl/slicers/slicer1.xml': SLICER.replace(
          'name="Country Slicer"',
          'name=""',
        ),
      },
      'Slicer or timeline name is invalid',
    ],
    [
      {
        'xl/slicers/slicer1.xml': SLICER.replace(
          'cache="Cache_Country"',
          'cache=""',
        ),
      },
      'Slicer or timeline cache name is invalid',
    ],
    [
      {
        'xl/slicers/slicer1.xml': `<x14:slicers xmlns:x14="${X14_NAMESPACE}"><x14:slicer>bad</x14:slicer></x14:slicers>`,
      },
      'XLSX slicer collection is invalid',
    ],
    [
      {
        'xl/slicers/slicer1.xml': `<x14:slicers xmlns:x14="${X14_NAMESPACE}"/>`,
      },
      'XLSX slicer or timeline display collection is empty',
    ],
    [
      {
        'xl/slicers/slicer1.xml': SLICER.replace(
          '<x14:slicer name=',
          '<evil:slicer xmlns:evil="urn:evil" name=',
        ),
      },
      'XLSX slicer element has the wrong namespace',
    ],
    [
      {
        'xl/slicerCaches/slicerCache1.bin': SLICER_CACHE.replace(
          '<x14:tabular ',
          '<evil:tabular xmlns:evil="urn:evil" ',
        ),
      },
      'Slicer cache source metadata is missing',
    ],
    [
      {
        'xl/slicers/slicer1.xml': SLICER.replace(
          'columnCount="3"',
          'columnCount="0"',
        ),
      },
      'Slicer column count is invalid',
    ],
    [
      {
        'xl/slicers/slicer1.xml': SLICER.replace(
          'columnCount="3"',
          'columnCount="bad"',
        ),
      },
      'Slicer column count is invalid',
    ],
    [
      {
        'xl/slicers/slicer1.xml': SLICER.replace(
          'showCaption="0"',
          'showCaption="bad"',
        ),
      },
      'Slicer show-caption flag is invalid',
    ],
    [
      {
        'xl/slicers/slicer1.xml': SLICER.replace(
          'lockedPosition="1"',
          'lockedPosition="bad"',
        ),
      },
      'Slicer locked-position flag is invalid',
    ],
    [
      {
        'xl/slicers/slicer1.xml': SLICER.replace(
          'startItem="2"',
          'startItem="bad"',
        ),
      },
      'Slicer start item is invalid',
    ],
    [
      {
        'xl/slicers/slicer1.xml': SLICER.replace(
          'rowHeight="240000"',
          'rowHeight="bad"',
        ),
      },
      'Slicer row height is invalid',
    ],
    [
      {
        'xl/slicers/slicer1.xml': SLICER.replace(' rowHeight="240000"', ''),
      },
      'Slicer row height is invalid',
    ],
    [
      {
        'xl/slicers/slicer1.xml': SLICER.replace('level="1"', 'level="bad"'),
      },
      'Slicer level is invalid',
    ],
    [
      {
        'xl/timelines/timeline1.xml': TIMELINE.replace(
          'level="2"',
          'level="bad"',
        ),
      },
      'Timeline level is invalid',
    ],
    [
      {
        'xl/timelines/timeline1.xml': TIMELINE.replace(
          'level="2"',
          'level="4294967296"',
        ),
      },
      'Timeline level is invalid',
    ],
    [
      {
        'xl/timelines/timeline1.xml': TIMELINE.replace(
          'selectionLevel="1"',
          'selectionLevel="bad"',
        ),
      },
      'Timeline selection level is invalid',
    ],
    [
      {
        'xl/timelines/timeline1.xml': TIMELINE.replace(
          ' selectionLevel="1"',
          '',
        ),
      },
      'Timeline level or selection level is invalid',
    ],
    [
      {
        'xl/timelines/timeline1.xml': TIMELINE.replace(
          'scrollPosition="2024-01-01T00:00:00Z"',
          'scrollPosition="bad"',
        ),
      },
      'Timeline scroll position is invalid',
    ],
    [
      {
        'xl/timelines/timeline1.xml': TIMELINE.replace(
          'showHeader="0"',
          'showHeader="bad"',
        ),
      },
      'Timeline show-header flag is invalid',
    ],
    [
      {
        'xl/timelines/timeline1.xml': TIMELINE.replace(
          'showHorizontalScrollbar="1"',
          'showHorizontalScrollbar="bad"',
        ),
      },
      'Timeline horizontal-scrollbar flag is invalid',
    ],
    [
      {
        'xl/timelines/timeline1.xml': TIMELINE.replace(
          'showSelectionLabel="1"',
          'showSelectionLabel="bad"',
        ),
      },
      'Timeline selection-label flag is invalid',
    ],
    [
      {
        'xl/timelines/timeline1.xml': TIMELINE.replace(
          'showTimeLevel="0"',
          'showTimeLevel="bad"',
        ),
      },
      'Timeline time-level flag is invalid',
    ],
    [
      {
        'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replace(
          '</Relationships>',
          `<Relationship Id="slicer2" Type="${SLICER_RELATIONSHIP}" Target="../slicers/slicer1.xml"/></Relationships>`,
        ),
      },
      'XLSX slicer or timeline relationship target is duplicated',
    ],
    [
      {
        'xl/slicers/slicer1.xml': `<x14:slicers xmlns:x14="${X14_NAMESPACE}"><x14:slicer name="Straße" cache="Cache_Country" rowHeight="200000"/><x14:slicer name="STRASSE" cache="Cache_Country" rowHeight="200000"/></x14:slicers>`,
      },
      'Worksheet contains duplicate slicer or timeline names',
    ],
  ] as const)(
    'rejects invalid slicer or timeline metadata %#',
    async (changes, message) => {
      expect((await capture(changes)).diagnostic.message).toBe(message);
    },
  );
});
