import { describe, expect, it } from 'vitest';

import { parseXlsx, XlsxParseError } from '../../src/formats/xlsx';
import {
  createIndependentXlsx,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

function worksheet(filter: string): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/>${filter}</worksheet>`;
}

const COMPLETE_FILTER = `<autoFilter ref="A1:G10">
  <filterColumn colId="0" hiddenButton="1" showButton="0">
    <filters blank="1" calendarType="gregorian">
      <filter val="Open"/><filter val="Closed"/>
      <dateGroupItem dateTimeGrouping="day" year="2024" month="2" day="29"/>
    </filters>
  </filterColumn>
  <filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThan" val="10"/><customFilter operator="lessThanOrEqual" val="20"/></customFilters></filterColumn>
  <filterColumn colId="2"><dynamicFilter type="aboveAverage" val="12.5" maxVal="20"/></filterColumn>
  <filterColumn colId="3"><top10 top="0" percent="1" val="5" filterVal="42"/></filterColumn>
  <filterColumn colId="4"><colorFilter cellColor="0"/></filterColumn>
  <filterColumn colId="5"><iconFilter iconSet="3Arrows" iconId="2"/></filterColumn>
  <filterColumn colId="6"/>
  <sortState ref="A2:G10" caseSensitive="1" columnSort="0" sortMethod="pinYin">
    <sortCondition ref="A2:A10" descending="1"/>
    <sortCondition ref="B2:B10" sortBy="fontColor"/>
    <sortCondition ref="C2:C10" sortBy="icon" iconSet="3Flags" iconId="1"/>
  </sortState>
</autoFilter>`;

async function parseFilter(
  filter: string,
  options: Parameters<typeof parseXlsx>[1] = {},
) {
  return parseXlsx(
    await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': worksheet(filter),
    }),
    options,
  );
}

async function capture(
  filter: string,
  options: Parameters<typeof parseXlsx>[1] = {},
): Promise<XlsxParseError> {
  try {
    await parseFilter(filter, options);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected XLSX auto-filter parsing to fail');
}

describe('XLSX auto filters and sorts', () => {
  it('preserves authored value, date, custom, dynamic, top, color, icon, and sort rules', async () => {
    const document = await parseFilter(COMPLETE_FILTER);
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    expect(sheet.kind === 'worksheet' ? sheet.autoFilter : undefined).toEqual({
      columns: [
        {
          columnId: 0,
          hiddenButton: true,
          rule: {
            blank: true,
            calendarType: 'gregorian',
            dates: [
              {
                day: 29,
                grouping: 'day',
                month: 2,
                year: 2024,
              },
            ],
            kind: 'values',
            values: ['Open', 'Closed'],
          },
          showButton: false,
        },
        {
          columnId: 1,
          hiddenButton: false,
          rule: {
            and: true,
            conditions: [
              { operator: 'greater-than', value: '10' },
              { operator: 'less-than-or-equal', value: '20' },
            ],
            kind: 'custom',
          },
          showButton: true,
        },
        {
          columnId: 2,
          hiddenButton: false,
          rule: {
            kind: 'dynamic',
            maxValue: 20,
            type: 'aboveAverage',
            value: 12.5,
          },
          showButton: true,
        },
        {
          columnId: 3,
          hiddenButton: false,
          rule: {
            filterValue: 42,
            kind: 'top',
            percent: true,
            top: false,
            value: 5,
          },
          showButton: true,
        },
        {
          columnId: 4,
          hiddenButton: false,
          rule: { cellColor: false, kind: 'color' },
          showButton: true,
        },
        {
          columnId: 5,
          hiddenButton: false,
          rule: { iconId: 2, iconSet: '3Arrows', kind: 'icon' },
          showButton: true,
        },
        {
          columnId: 6,
          hiddenButton: false,
          rule: { kind: 'none' },
          showButton: true,
        },
      ],
      range: {
        end: { column: 7, row: 10 },
        reference: 'A1:G10',
        start: { column: 1, row: 1 },
      },
      selectionRelation: 'full-sheet',
      sort: {
        caseSensitive: true,
        columnSort: false,
        conditions: [
          {
            descending: true,
            range: {
              end: { column: 1, row: 10 },
              reference: 'A2:A10',
              start: { column: 1, row: 2 },
            },
            sortBy: 'value',
          },
          {
            descending: false,
            range: {
              end: { column: 2, row: 10 },
              reference: 'B2:B10',
              start: { column: 2, row: 2 },
            },
            sortBy: 'font-color',
          },
          {
            descending: false,
            iconId: 1,
            iconSet: '3Flags',
            range: {
              end: { column: 3, row: 10 },
              reference: 'C2:C10',
              start: { column: 3, row: 2 },
            },
            sortBy: 'icon',
          },
        ],
        range: {
          end: { column: 7, row: 10 },
          reference: 'A2:G10',
          start: { column: 1, row: 2 },
        },
        sortMethod: 'pin-yin',
      },
    });
  });

  it('applies selection intersection without filtering worksheet rows', async () => {
    const included = await parseFilter(COMPLETE_FILTER, {
      selection: { ranges: { Sheet1: ['G10'] } },
    });
    const includedSheet = included.sheets[0]!;
    expect(
      includedSheet.kind === 'worksheet'
        ? includedSheet.autoFilter?.selectionRelation
        : undefined,
    ).toBe('intersects-selection');
    const excluded = await parseFilter(COMPLETE_FILTER, {
      selection: { ranges: { Sheet1: ['H1'] } },
    });
    const excludedSheet = excluded.sheets[0]!;
    expect(
      excludedSheet.kind === 'worksheet' ? excludedSheet.autoFilter : undefined,
    ).toBeUndefined();
  });

  it.each([
    'gregorian',
    'gregorianArabic',
    'gregorianMeFrench',
    'gregorianUs',
    'gregorianXlitEnglish',
    'gregorianXlitFrench',
    'hebrew',
    'hijri',
    'japan',
    'korea',
    'none',
    'saka',
    'taiwan',
    'thai',
  ])('accepts calendar type %s', async (calendarType) => {
    const document = await parseFilter(
      `<autoFilter ref="A1"><filterColumn colId="0"><filters blank="1" calendarType="${calendarType}"/></filterColumn></autoFilter>`,
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.autoFilter?.columns[0]?.rule
        : undefined,
    ).toMatchObject({ calendarType, kind: 'values' });
  });

  it.each([
    'aboveAverage',
    'belowAverage',
    'lastMonth',
    'lastQuarter',
    'lastWeek',
    'lastYear',
    'M1',
    'M10',
    'M11',
    'M12',
    'M2',
    'M3',
    'M4',
    'M5',
    'M6',
    'M7',
    'M8',
    'M9',
    'nextMonth',
    'nextQuarter',
    'nextWeek',
    'nextYear',
    'null',
    'Q1',
    'Q2',
    'Q3',
    'Q4',
    'thisMonth',
    'thisQuarter',
    'thisWeek',
    'thisYear',
    'today',
    'tomorrow',
    'yearToDate',
    'yesterday',
  ])('accepts dynamic filter type %s', async (type) => {
    const document = await parseFilter(
      `<autoFilter ref="A1"><filterColumn colId="0"><dynamicFilter type="${type}"/></filterColumn></autoFilter>`,
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.autoFilter?.columns[0]?.rule
        : undefined,
    ).toEqual({ kind: 'dynamic', type });
  });

  it.each([
    '3Arrows',
    '3ArrowsGray',
    '3Flags',
    '3Signs',
    '3Symbols',
    '3Symbols2',
    '3TrafficLights1',
    '3TrafficLights2',
    '4Arrows',
    '4ArrowsGray',
    '4Rating',
    '4RedToBlack',
    '4TrafficLights',
    '5Arrows',
    '5ArrowsGray',
    '5Quarters',
    '5Rating',
  ])('accepts icon filter set %s', async (iconSet) => {
    const document = await parseFilter(
      `<autoFilter ref="A1"><filterColumn colId="0"><iconFilter iconSet="${iconSet}" iconId="0"/></filterColumn></autoFilter>`,
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.autoFilter?.columns[0]?.rule
        : undefined,
    ).toEqual({ iconId: 0, iconSet, kind: 'icon' });
  });

  it.each([
    ['equal', 'equal'],
    ['greaterThan', 'greater-than'],
    ['greaterThanOrEqual', 'greater-than-or-equal'],
    ['lessThan', 'less-than'],
    ['lessThanOrEqual', 'less-than-or-equal'],
    ['notEqual', 'not-equal'],
  ] as const)('normalizes custom operator %s', async (source, operator) => {
    const document = await parseFilter(
      `<autoFilter ref="A1"><filterColumn colId="0"><customFilters><customFilter operator="${source}" val="x"/></customFilters></filterColumn></autoFilter>`,
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.autoFilter?.columns[0]?.rule
        : undefined,
    ).toMatchObject({ conditions: [{ operator, value: 'x' }] });
  });

  it.each([
    ['year', 'year="2024"', { grouping: 'year', year: 2024 }],
    [
      'month',
      'year="2024" month="2"',
      { grouping: 'month', month: 2, year: 2024 },
    ],
    [
      'day',
      'year="2024" month="2" day="29"',
      { day: 29, grouping: 'day', month: 2, year: 2024 },
    ],
    [
      'hour',
      'year="2024" month="2" day="29" hour="23"',
      { day: 29, grouping: 'hour', hour: 23, month: 2, year: 2024 },
    ],
    [
      'minute',
      'year="2024" month="2" day="29" hour="23" minute="59"',
      {
        day: 29,
        grouping: 'minute',
        hour: 23,
        minute: 59,
        month: 2,
        year: 2024,
      },
    ],
    [
      'second',
      'year="2024" month="2" day="29" hour="23" minute="59" second="58"',
      {
        day: 29,
        grouping: 'second',
        hour: 23,
        minute: 59,
        month: 2,
        second: 58,
        year: 2024,
      },
    ],
  ] as const)(
    'accepts complete date grouping %s',
    async (grouping, fields, expected) => {
      const document = await parseFilter(
        `<autoFilter ref="A1"><filterColumn colId="0"><filters><dateGroupItem dateTimeGrouping="${grouping}" ${fields}/></filters></filterColumn></autoFilter>`,
      );
      const sheet = document.sheets[0]!;
      const rule =
        sheet.kind === 'worksheet'
          ? sheet.autoFilter?.columns[0]?.rule
          : undefined;
      expect(rule?.kind === 'values' ? rule.dates[0] : undefined).toEqual(
        expected,
      );
    },
  );

  it.each([
    ['none', 'none'],
    ['pinYin', 'pin-yin'],
    ['stroke', 'stroke'],
  ] as const)('normalizes sort method %s', async (source, sortMethod) => {
    const document = await parseFilter(
      `<autoFilter ref="A1"><sortState ref="A1" sortMethod="${source}"><sortCondition ref="A1"/></sortState></autoFilter>`,
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.autoFilter?.sort?.sortMethod
        : undefined,
    ).toBe(sortMethod);
  });

  it.each([
    ['<autoFilter ref="bad"/>', 'Auto-filter range is invalid'],
    [
      '<autoFilter ref="A1">bad</autoFilter>',
      'Auto-filter text content is invalid',
    ],
    [
      '<autoFilter ref="A1"/><autoFilter ref="A1"/>',
      'Worksheet contains duplicate autoFilter elements',
    ],
    [
      '<autoFilter ref="A1:B2"><filterColumn colId="2"/></autoFilter>',
      'Filter column ID is invalid',
    ],
    [
      '<autoFilter ref="A1:B2"><filterColumn colId="0"/><filterColumn colId="0"/></autoFilter>',
      'Auto-filter contains duplicate column IDs',
    ],
    [
      '<autoFilter ref="A1"><filterColumn colId="0"><filters/><top10 val="1"/></filterColumn></autoFilter>',
      'Filter column contains multiple rule kinds',
    ],
    [
      '<autoFilter ref="A1"><filterColumn colId="0"><filters/></filterColumn></autoFilter>',
      'Value filter is empty',
    ],
    [
      '<autoFilter ref="A1"><filterColumn colId="0"><filters><dateGroupItem dateTimeGrouping="day" year="2024" month="2"/></filters></filterColumn></autoFilter>',
      'Date-group filter fields are incomplete',
    ],
    [
      '<autoFilter ref="A1"><filterColumn colId="0"><customFilters/></filterColumn></autoFilter>',
      'Custom-filter condition collection is invalid',
    ],
    [
      '<autoFilter ref="A1"><filterColumn colId="0"><customFilters><customFilter operator="bad" val="1"/></customFilters></filterColumn></autoFilter>',
      'Custom-filter operator is invalid',
    ],
    [
      '<autoFilter ref="A1"><filterColumn colId="0"><dynamicFilter type="bad"/></filterColumn></autoFilter>',
      'Dynamic-filter type is invalid',
    ],
    [
      '<autoFilter ref="A1"><filterColumn colId="0"><top10 val="-1"/></filterColumn></autoFilter>',
      'Top-filter value is invalid',
    ],
    [
      '<autoFilter ref="A1"><filterColumn colId="0"><iconFilter iconSet="bad" iconId="0"/></filterColumn></autoFilter>',
      'Icon-filter set is invalid',
    ],
    [
      '<autoFilter ref="A1"><sortState ref="A1"/></autoFilter>',
      'Sort condition collection is invalid',
    ],
    [
      '<autoFilter ref="A1"><sortState ref="A1"><sortCondition ref="A1" sortBy="icon"/></sortState></autoFilter>',
      'Icon sort condition metadata is missing',
    ],
  ] as const)(
    'rejects invalid filter/sort contract %#',
    async (filter, message) => {
      expect((await capture(filter)).diagnostic.message).toBe(message);
    },
  );

  it('enforces filter text and selection-work limits exactly', async () => {
    const filter = `<autoFilter ref="A1"><filterColumn colId="0"><filters><filter val="abc"/></filters></filterColumn></autoFilter>`;
    await expect(
      parseFilter(filter, { limits: { maxTextCharacters: 12 } }),
    ).resolves.toBeDefined();
    expect(
      (await capture(filter, { limits: { maxTextCharacters: 11 } })).diagnostic,
    ).toMatchObject({
      actual: 12,
      limit: 11,
      limitName: 'maxTextCharacters',
    });
    await expect(
      parseFilter('<autoFilter ref="A1"/>', {
        limits: { maxReturnedCells: 1, maxScannedCells: 1 },
        selection: { ranges: { Sheet1: ['A1'] } },
      }),
    ).resolves.toBeDefined();
  });
});
