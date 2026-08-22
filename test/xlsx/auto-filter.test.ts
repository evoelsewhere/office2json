import { describe, expect, it } from 'vitest';

import {
  parseXlsxAutoFilter,
  XlsxAutoFilterCapture,
} from '../../src/formats/xlsx/internal/auto-filter';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from '../../src/formats/xlsx/internal/resource-limits';
import type { XlsxResolvedSheetSelection } from '../../src/formats/xlsx/internal/selection';
import type { XlsxXmlElement } from '../../src/formats/xlsx/internal/streaming-xml';
import type { XlsxWorksheetBudget } from '../../src/formats/xlsx/internal/worksheet';
import { XlsxParseError } from '../../src/formats/xlsx';

const PART = 'xl/worksheets/sheet1.xml';
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

function parse(
  value: unknown,
  options: {
    differentialStyles?: number;
    limits?: Partial<ResolvedXlsxResourceLimits>;
    selection?: XlsxResolvedSheetSelection;
  } = {},
) {
  return parseXlsxAutoFilter(
    value,
    '',
    options.differentialStyles ?? 0,
    options.selection ?? FULL,
    budget(),
    { ...defaultXlsxResourceLimits(), ...options.limits },
    PART,
  );
}

function capture(action: () => unknown): XlsxParseError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected normalized auto-filter parsing to fail');
}

function captureLimit(action: () => unknown): XlsxResourceLimitError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxResourceLimitError);
    return error as XlsxResourceLimitError;
  }
  throw new Error(
    'Expected normalized auto-filter resource accounting to fail',
  );
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

describe('XLSX normalized auto-filter parser', () => {
  it('preserves decoded XML entities and numeric references in filter values', () => {
    const value = parse({
      attrs: { ref: 'A1' },
      filterColumn: {
        attrs: { colId: '0' },
        filters: {
          filter: [
            { attrs: { val: '&amp;&apos;&gt;&lt;&quot;' } },
            { attrs: { val: '&#65;&#x1F600;' } },
          ],
        },
      },
    });
    expect(value?.columns[0]?.rule).toMatchObject({
      values: ['&\'><"', 'A😀'],
    });
  });

  it.each([
    ['0', false],
    ['false', false],
    ['1', true],
    ['true', true],
  ] as const)('parses lexical boolean %s', (source, expected) => {
    const value = parse({
      attrs: { ref: 'A1' },
      filterColumn: {
        attrs: { colId: '0', hiddenButton: source, showButton: source },
      },
    });
    expect(value?.columns[0]).toMatchObject({
      hiddenButton: expected,
      showButton: expected,
    });
  });

  it.each([
    ['-1', 'Filter column ID is invalid'],
    ['01', 'Filter column ID is invalid'],
    ['1.0', 'Filter column ID is invalid'],
    ['9007199254740992', 'Filter column ID is invalid'],
    ['0x1', 'Filter column ID is invalid'],
    ['1x', 'Filter column ID is invalid'],
  ])('rejects invalid unsigned integer %s', (source, message) => {
    expect(
      capture(() =>
        parse({
          attrs: { ref: 'A1' },
          filterColumn: { attrs: { colId: source } },
        }),
      ).diagnostic.message,
    ).toBe(message);
  });

  it('rejects trailing whitespace after an unsigned integer', () => {
    expect(
      capture(() =>
        parse({
          attrs: { ref: 'A1:B1' },
          filterColumn: { attrs: { colId: '1 ' } },
        }),
      ).diagnostic.message,
    ).toBe('Filter column ID is invalid');
  });

  it.each(['+1', '-1.5', '.5', '.55', '1.', '1e3', '1e30', '-2.5E-2'])(
    'accepts finite numeric lexical form %s',
    (source) => {
      const value = parse({
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          dynamicFilter: { attrs: { type: 'aboveAverage', val: source } },
        },
      });
      expect(value?.columns[0]?.rule).toMatchObject({
        kind: 'dynamic',
        value: Number(source),
      });
    },
  );

  it('normalizes negative zero and rejects lexically valid non-finite numbers', () => {
    const value = parse({
      attrs: { ref: 'A1' },
      filterColumn: {
        attrs: { colId: '0' },
        dynamicFilter: { attrs: { type: 'aboveAverage', val: '-0' } },
      },
    });
    const rule = value?.columns[0]?.rule;
    const normalized = rule?.kind === 'dynamic' ? rule.value : undefined;
    expect(normalized).toBe(0);
    expect(Object.is(normalized, -0)).toBe(false);
    expect(
      capture(() =>
        parse({
          attrs: { ref: 'A1' },
          filterColumn: {
            attrs: { colId: '0' },
            dynamicFilter: {
              attrs: { type: 'aboveAverage', val: '1e309' },
            },
          },
        }),
      ).diagnostic.message,
    ).toBe('Dynamic-filter value is invalid');
  });

  it.each(['', ' 1', '1 ', '+', '.', '1e', '1e+', '1_0', 'NaN', 'Infinity'])(
    'rejects invalid numeric lexical form %#',
    (source) => {
      expect(
        capture(() =>
          parse({
            attrs: { ref: 'A1' },
            filterColumn: {
              attrs: { colId: '0' },
              dynamicFilter: {
                attrs: { type: 'aboveAverage', val: source },
              },
            },
          }),
        ).diagnostic.message,
      ).toBe('Dynamic-filter value is invalid');
    },
  );

  it('validates differential styles for filters and sort conditions', () => {
    const tree = {
      attrs: { ref: 'A1' },
      filterColumn: {
        attrs: { colId: '0' },
        colorFilter: { attrs: { dxfId: '0' } },
      },
      sortState: {
        attrs: { ref: 'A1' },
        sortCondition: {
          attrs: { dxfId: '0', ref: 'A1', sortBy: 'cellColor' },
        },
      },
    };
    expect(parse(tree, { differentialStyles: 1 })).toMatchObject({
      columns: [{ rule: { differentialStyle: 0 } }],
      sort: { conditions: [{ differentialStyle: 0 }] },
    });
    expect(capture(() => parse(tree)).diagnostic.message).toBe(
      'Color-filter differential-style reference is invalid',
    );
  });

  it('parses prefixed normalized child keys', () => {
    expect(
      parseXlsxAutoFilter(
        {
          attrs: { ref: 'A1' },
          's:filterColumn': {
            attrs: { colId: '0' },
            's:filters': { attrs: { blank: '1' } },
          },
        },
        's',
        0,
        FULL,
        budget(),
        defaultXlsxResourceLimits(),
        PART,
      )?.columns[0]?.rule,
    ).toMatchObject({ blank: true, kind: 'values' });
  });

  it('preserves defaults and optional numeric fields exactly', () => {
    expect(
      parse({
        attrs: { ref: 'A1:C1' },
        filterColumn: [
          { attrs: { colId: '0' }, top10: { attrs: { val: '0' } } },
          { attrs: { colId: '1' }, colorFilter: { attrs: {} } },
          {
            attrs: { colId: '2' },
            dynamicFilter: { attrs: { type: 'today' } },
          },
        ],
      })?.columns.map((column) => column.rule),
    ).toEqual([
      { kind: 'top', percent: false, top: true, value: 0 },
      { cellColor: true, kind: 'color' },
      { kind: 'dynamic', type: 'today' },
    ]);
  });

  it('preserves an explicitly authored value sort kind', () => {
    expect(
      parse({
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { ref: 'A1' },
          sortCondition: { attrs: { ref: 'A1', sortBy: 'value' } },
        },
      })?.sort?.conditions[0]?.sortBy,
    ).toBe('value');
  });

  it('preserves a bounded sort custom list and charges its public text', () => {
    const tree = {
      attrs: { ref: 'A1' },
      sortState: {
        attrs: { ref: 'A1' },
        sortCondition: {
          attrs: { customList: 'High,Low', ref: 'A1' },
        },
      },
    };
    expect(parse(tree)?.sort?.conditions[0]).toMatchObject({
      customList: 'High,Low',
    });
    expect(
      captureLimit(() => parse(tree, { limits: { maxTextCharacters: 7 } })),
    ).toMatchObject({
      actual: 8,
      limit: 7,
      limitName: 'maxTextCharacters',
    });
    expect(
      capture(() =>
        parse({
          attrs: { ref: 'A1' },
          sortState: {
            attrs: { ref: 'A1' },
            sortCondition: { attrs: { customList: 1, ref: 'A1' } },
          },
        }),
      ).diagnostic.message,
    ).toBe('Sort condition custom list is invalid');
  });

  it('enforces icon IDs against each icon-set cardinality', () => {
    expect(
      capture(() =>
        parse({
          attrs: { ref: 'A1' },
          filterColumn: {
            attrs: { colId: '0' },
            iconFilter: { attrs: { iconId: '3', iconSet: '3Arrows' } },
          },
        }),
      ).diagnostic.message,
    ).toBe('Icon-filter ID is invalid');
    expect(
      capture(() =>
        parse({
          attrs: { ref: 'A1' },
          sortState: {
            attrs: { ref: 'A1' },
            sortCondition: {
              attrs: {
                iconId: '4',
                iconSet: '4Arrows',
                ref: 'A1',
                sortBy: 'icon',
              },
            },
          },
        }),
      ).diagnostic.message,
    ).toBe('Sort condition icon ID is invalid');
    expect(
      parse({
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          iconFilter: { attrs: { iconId: '4', iconSet: '5Arrows' } },
        },
      })?.columns[0]?.rule,
    ).toMatchObject({ iconId: 4, iconSet: '5Arrows' });
  });

  it('accounts for filter, sort-state, and sort-condition ranges exactly', () => {
    const tree = {
      attrs: { ref: 'A1:B2' },
      sortState: {
        attrs: { ref: 'A1:B2' },
        sortCondition: { attrs: { ref: 'A1:A2' } },
      },
    };
    expect(parse(tree, { limits: { maxRangeAreas: 3 } })).toBeDefined();
    expect(
      captureLimit(() => parse(tree, { limits: { maxRangeAreas: 2 } })),
    ).toMatchObject({
      actual: 3,
      limit: 2,
      limitName: 'maxRangeAreas',
    });
  });

  it.each([
    ['A1:B2', 'A1:C2', 'A1:B2', 'Sort-state range is outside the auto-filter'],
    [
      'A1:B2',
      'A1:B2',
      'A1:C2',
      'Sort condition range is outside the sort state',
    ],
    [
      'A1:B2',
      'A1:B2',
      'A2:B3',
      'Sort condition range is outside the sort state',
    ],
    [
      'A1:C2',
      'B1:C2',
      'A1:B2',
      'Sort condition range is outside the sort state',
    ],
    [
      'A1:B3',
      'A2:B3',
      'A1:B3',
      'Sort condition range is outside the sort state',
    ],
  ] as const)(
    'rejects sort containment %#',
    (filterReference, stateReference, conditionReference, message) => {
      const tree = {
        attrs: { ref: filterReference },
        sortState: {
          attrs: { ref: stateReference },
          sortCondition: { attrs: { ref: conditionReference } },
        },
      };
      expect(capture(() => parse(tree)).diagnostic.message).toBe(message);
    },
  );

  it('enforces the fixed SpreadsheetML sort-condition cardinality', () => {
    const conditions = Array.from({ length: 64 }, () => ({
      attrs: { ref: 'A1' },
    }));
    expect(
      parse({
        attrs: { ref: 'A1' },
        sortState: { attrs: { ref: 'A1' }, sortCondition: conditions },
      })?.sort?.conditions,
    ).toHaveLength(64);
    expect(
      capture(() =>
        parse({
          attrs: { ref: 'A1' },
          sortState: {
            attrs: { ref: 'A1' },
            sortCondition: [...conditions, { attrs: { ref: 'A1' } }],
          },
        }),
      ).diagnostic.message,
    ).toBe('Sort condition count exceeds the SpreadsheetML bound');
  });

  it.each([
    ['A2', { maxRowsPerWorksheet: 1 }, 'maxRowsPerWorksheet'],
    ['B1', { maxColumnsPerWorksheet: 1 }, 'maxColumnsPerWorksheet'],
  ] as const)(
    'enforces configured filter range bound %s',
    (ref, limits, name) => {
      expect(
        captureLimit(() => parse({ attrs: { ref } }, { limits })),
      ).toMatchObject({ actual: 2, limit: 1, limitName: name });
    },
  );

  it('accepts the exact configured row bound', () => {
    expect(
      parse({ attrs: { ref: 'A1' } }, { limits: { maxRowsPerWorksheet: 1 } })
        ?.range.reference,
    ).toBe('A1');
  });

  it('rejects filter-column counts above the configured bound before IDs', () => {
    expect(
      capture(() =>
        parse(
          {
            attrs: { ref: 'A1' },
            filterColumn: [
              { attrs: { colId: '0' } },
              { attrs: { colId: '0' } },
            ],
          },
          { limits: { maxColumnsPerWorksheet: 1 } },
        ),
      ).diagnostic.message,
    ).toBe('Filter column count exceeds the worksheet bound');
  });

  it.each([
    [{ attrs: { ref: '$A$1' } }, 'Auto-filter range is invalid'],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: { attrs: { colId: '0', hiddenButton: 'bad' } },
      },
      'Filter hidden-button flag is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: { attrs: { colId: '0', showButton: 'bad' } },
      },
      'Filter show-button flag is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          filters: { filter: 'bad' },
        },
      },
      'Value-filter collection is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          filters: { filter: { attrs: {} } },
        },
      },
      'Filter value is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          filters: { attrs: { blank: 'bad' } },
        },
      },
      'Filter blank flag is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          filters: { attrs: { blank: '1', calendarType: 'bad' } },
        },
      },
      'Filter calendar type is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          customFilters: {
            attrs: { and: 'bad' },
            customFilter: { attrs: { val: '1' } },
          },
        },
      },
      'Custom-filter conjunction is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          customFilters: { customFilter: { attrs: {} } },
        },
      },
      'Custom-filter value is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          dynamicFilter: { attrs: { maxVal: 'bad', type: 'today' } },
        },
      },
      'Dynamic-filter maximum is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          top10: { attrs: { val: 'bad' } },
        },
      },
      'Top-filter value is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          top10: { attrs: { filterVal: 'bad', val: '1' } },
        },
      },
      'Top-filter threshold is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          top10: { attrs: { percent: 'bad', val: '1' } },
        },
      },
      'Top-filter percent flag is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          top10: { attrs: { top: 'bad', val: '1' } },
        },
      },
      'Top-filter direction is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          colorFilter: { attrs: { cellColor: 'bad' } },
        },
      },
      'Color-filter cell-color flag is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          iconFilter: { attrs: { iconId: '5', iconSet: '3Arrows' } },
        },
      },
      'Icon-filter ID is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { ref: 'A1' },
          sortCondition: { attrs: { ref: 'A1', sortBy: 'bad' } },
        },
      },
      'Sort condition kind is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { ref: 'A1' },
          sortCondition: {
            attrs: { iconSet: 'bad', ref: 'A1', sortBy: 'icon' },
          },
        },
      },
      'Sort condition icon set is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { ref: 'A1' },
          sortCondition: {
            attrs: {
              iconId: '5',
              iconSet: '3Flags',
              ref: 'A1',
              sortBy: 'icon',
            },
          },
        },
      },
      'Sort condition icon ID is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { ref: 'A1' },
          sortCondition: {
            attrs: { iconId: '0', ref: 'A1', sortBy: 'icon' },
          },
        },
      },
      'Icon sort condition metadata is missing',
    ],
    [
      {
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { ref: 'A1' },
          sortCondition: {
            attrs: { iconSet: '3Flags', ref: 'A1', sortBy: 'icon' },
          },
        },
      },
      'Icon sort condition metadata is missing',
    ],
    [
      {
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { ref: 'A1' },
          sortCondition: {
            attrs: { dxfId: '0', ref: 'A1', sortBy: 'cellColor' },
          },
        },
      },
      'Sort condition differential-style reference is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { ref: 'A1' },
          sortCondition: { attrs: { descending: 'bad', ref: 'A1' } },
        },
      },
      'Sort condition direction is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { ref: 'A1' },
          sortCondition: { attrs: { ref: 'bad' } },
        },
      },
      'Sort condition range is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { caseSensitive: 'bad', ref: 'A1' },
          sortCondition: { attrs: { ref: 'A1' } },
        },
      },
      'Sort case-sensitive flag is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { columnSort: 'bad', ref: 'A1' },
          sortCondition: { attrs: { ref: 'A1' } },
        },
      },
      'Sort column direction flag is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { ref: 'bad' },
          sortCondition: { attrs: { ref: 'A1' } },
        },
      },
      'Sort-state range is invalid',
    ],
  ] as const)('reports exact invalid rule %#', (tree, message) => {
    expect(capture(() => parse(tree)).diagnostic.message).toBe(message);
  });

  it.each([
    [
      { dateTimeGrouping: 'bad', year: '2024' },
      'Date-group filter grouping is invalid',
    ],
    [{ dateTimeGrouping: 'year' }, 'Date-group filter year is invalid'],
    [
      { dateTimeGrouping: 'year', year: '10000' },
      'Date-group filter year is invalid',
    ],
    [
      { dateTimeGrouping: 'month', month: '0', year: '2024' },
      'Date-group filter fields are incomplete',
    ],
    [
      { dateTimeGrouping: 'month', month: '13', year: '2024' },
      'Date-group filter month is invalid',
    ],
    [
      { dateTimeGrouping: 'day', day: '0', month: '1', year: '2024' },
      'Date-group filter fields are incomplete',
    ],
    [
      { dateTimeGrouping: 'day', day: '32', month: '1', year: '2024' },
      'Date-group filter day is invalid',
    ],
    [
      {
        dateTimeGrouping: 'hour',
        day: '1',
        hour: '24',
        month: '1',
        year: '2024',
      },
      'Date-group filter hour is invalid',
    ],
    [
      {
        dateTimeGrouping: 'minute',
        day: '1',
        hour: '1',
        minute: '60',
        month: '1',
        year: '2024',
      },
      'Date-group filter minute is invalid',
    ],
    [
      {
        dateTimeGrouping: 'second',
        day: '1',
        hour: '1',
        minute: '1',
        month: '1',
        second: '60',
        year: '2024',
      },
      'Date-group filter second is invalid',
    ],
    [
      {
        dateTimeGrouping: 'second',
        day: '1',
        hour: '1',
        minute: '1',
        month: '1',
        year: '2024',
      },
      'Date-group filter fields are incomplete',
    ],
    [
      { dateTimeGrouping: 'month', year: '2024' },
      'Date-group filter fields are incomplete',
    ],
    [
      { dateTimeGrouping: 'hour', day: '1', month: '1', year: '2024' },
      'Date-group filter fields are incomplete',
    ],
    [
      {
        dateTimeGrouping: 'minute',
        day: '1',
        hour: '1',
        month: '1',
        year: '2024',
      },
      'Date-group filter fields are incomplete',
    ],
    [
      {
        dateTimeGrouping: 'second',
        day: '1',
        month: '1',
        year: '2024',
      },
      'Date-group filter fields are incomplete',
    ],
    [
      {
        dateTimeGrouping: 'second',
        day: '1',
        hour: '1',
        month: '1',
        second: '1',
        year: '2024',
      },
      'Date-group filter fields are incomplete',
    ],
  ] as const)('rejects invalid date group %#', (attrs, message) => {
    const tree = {
      attrs: { ref: 'A1' },
      filterColumn: {
        attrs: { colId: '0' },
        filters: { dateGroupItem: { attrs } },
      },
    };
    expect(capture(() => parse(tree)).diagnostic.message).toBe(message);
  });

  it('requires an hour for every grouping at hour precision or finer', () => {
    for (const dateTimeGrouping of ['hour', 'minute', 'second'] as const) {
      const tree = {
        attrs: { ref: 'A1' },
        filterColumn: {
          attrs: { colId: '0' },
          filters: {
            dateGroupItem: {
              attrs: {
                dateTimeGrouping,
                day: '1',
                minute: '1',
                month: '1',
                second: '1',
                year: '2024',
              },
            },
          },
        },
      };
      expect(capture(() => parse(tree)).diagnostic.message).toBe(
        'Date-group filter fields are incomplete',
      );
    }
  });

  it('accepts the exact configured range bound and rejects one over', () => {
    expect(
      parse(
        {
          attrs: { ref: 'A1' },
          filterColumn: { attrs: { colId: '0' } },
        },
        { limits: { maxColumnsPerWorksheet: 1 } },
      )?.columns,
    ).toHaveLength(1);
    try {
      parse(
        { attrs: { ref: 'A1:B1' } },
        { limits: { maxColumnsPerWorksheet: 1 } },
      );
    } catch (error) {
      expect(error).toBeInstanceOf(XlsxResourceLimitError);
      expect(error).toMatchObject({
        actual: 2,
        limit: 1,
        limitName: 'maxColumnsPerWorksheet',
      });
      return;
    }
    throw new Error('Expected the auto-filter column bound to fail');
  });

  it('distinguishes every rectangle boundary during selection', () => {
    const selected = (reference: string): XlsxResolvedSheetSelection => {
      const [column, row] = [
        reference.codePointAt(0)! - 0x40,
        Number(reference.slice(1)),
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
    for (const reference of ['A3', 'F3', 'C1', 'C6']) {
      expect(
        parse({ attrs: { ref: 'B2:E5' } }, { selection: selected(reference) }),
      ).toBeUndefined();
    }
    for (const reference of ['B2', 'E5', 'C3']) {
      expect(
        parse({ attrs: { ref: 'B2:E5' } }, { selection: selected(reference) })
          ?.selectionRelation,
      ).toBe('intersects-selection');
    }
  });

  it('omits filters for an explicitly non-selected worksheet', () => {
    expect(
      parse({ attrs: { ref: 'A1' } }, { selection: { kind: 'not-selected' } }),
    ).toBeUndefined();
  });

  it('rejects more than two custom-filter conditions', () => {
    expect(
      capture(() =>
        parse({
          attrs: { ref: 'A1' },
          filterColumn: {
            attrs: { colId: '0' },
            customFilters: {
              customFilter: [
                { attrs: { val: '1' } },
                { attrs: { val: '2' } },
                { attrs: { val: '3' } },
              ],
            },
          },
        }),
      ).diagnostic.message,
    ).toBe('Custom-filter condition collection is invalid');
  });

  it.each([
    [null, 'Auto-filter structure is invalid'],
    [[], 'Auto-filter structure is invalid'],
    [
      { attrs: { ref: 'A1' }, filterColumn: 'bad' },
      'Filter column collection is invalid',
    ],
    [
      {
        attrs: { ref: 'A1' },
        filterColumn: { attrs: { colId: '0' }, filters: [] },
      },
      'Filter rule is invalid',
    ],
    [{ attrs: { ref: 'A1' }, sortState: [] }, 'Sort state is invalid'],
    [
      {
        attrs: { ref: 'A1' },
        sortState: {
          attrs: { ref: 'A1', sortMethod: 'bad' },
          sortCondition: { attrs: { ref: 'A1' } },
        },
      },
      'Sort method is invalid',
    ],
  ] as const)('rejects invalid normalized structure %#', (value, message) => {
    expect(capture(() => parse(value)).diagnostic.message).toBe(message);
  });
});

describe('XLSX auto-filter streaming capture', () => {
  function captureBuilder(): XlsxAutoFilterCapture {
    return new XlsxAutoFilterCapture(
      0,
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
  }

  it('captures a complete normalized subtree', () => {
    const builder = captureBuilder();
    builder.openElement(element('autoFilter', { ref: 'A1' }));
    builder.openElement(element('filterColumn', { colId: '0' }));
    builder.closeElement(element('filterColumn'));
    builder.closeElement(element('autoFilter'));
    expect(builder.result()).toMatchObject({
      columns: [{ columnId: 0 }],
      range: { reference: 'A1' },
    });
  });

  it('captures repeated children and both raw and expanded attribute names', () => {
    const builder = captureBuilder();
    builder.openElement({
      attributes: new Map([['ref', 'A1:C1']]),
      localName: 'autoFilter',
      namespace: 'worksheet',
    });
    for (const colId of ['0', '1', '2']) {
      builder.openElement(element('filterColumn', { colId }));
      builder.closeElement(element('filterColumn'));
    }
    builder.text(' \n ');
    builder.closeElement(element('autoFilter'));
    expect(builder.result()?.columns.map((column) => column.columnId)).toEqual([
      0, 1, 2,
    ]);
  });

  it('rejects invalid roots, nesting, text, and incomplete captures', () => {
    expect(() => captureBuilder().openElement(element('other'))).toThrow(
      'Auto-filter capture root is invalid',
    );
    const duplicate = captureBuilder();
    duplicate.openElement(element('autoFilter', { ref: 'A1' }));
    duplicate.closeElement(element('autoFilter'));
    expect(() => duplicate.openElement(element('autoFilter'))).toThrow(
      'Auto-filter capture root is invalid',
    );
    const nesting = captureBuilder();
    nesting.openElement(element('autoFilter', { ref: 'A1' }));
    expect(() => nesting.closeElement(element('filterColumn'))).toThrow(
      'Auto-filter capture nesting is invalid',
    );
    const text = captureBuilder();
    text.openElement(element('autoFilter', { ref: 'A1' }));
    expect(() => text.text('not whitespace')).toThrow(
      'Auto-filter text content is invalid',
    );
    expect(() => captureBuilder().result()).toThrow(
      'Auto-filter capture is incomplete',
    );
    const incomplete = captureBuilder();
    incomplete.openElement(element('autoFilter', { ref: 'A1' }));
    expect(() => incomplete.result()).toThrow(
      'Auto-filter capture is incomplete',
    );
  });
});
