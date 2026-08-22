import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { parseXlsxWorkbookViews } from '../../src/formats/xlsx/internal/workbook-views';
import type { XlsxSheet, XlsxWorksheet } from '../../src/formats/xlsx/types';

const PART = 'xl/workbook.xml';

function worksheet(
  index: number,
  state: XlsxWorksheet['state'] = 'visible',
): XlsxWorksheet {
  return {
    columns: [],
    comments: [],
    conditionalFormattings: [],
    dataValidations: [],
    drawings: [],
    hyperlinks: [],
    index,
    kind: 'worksheet',
    mergedRanges: [],
    name: `Sheet${index + 1}`,
    payload: 'full-sheet',
    protectedRanges: [],
    rows: [],
    state,
    tables: [],
    views: [],
  };
}

function view(attrs: Record<string, unknown> = {}): Record<string, unknown> {
  return { attrs };
}

function collection(...views: unknown[]): Record<string, unknown> {
  return { workbookView: views.length === 1 ? views[0] : views };
}

function parse(
  value: unknown,
  sheets: readonly XlsxSheet[] = [worksheet(0), worksheet(1), worksheet(2)],
) {
  return parseXlsxWorkbookViews(value, '', PART, sheets);
}

function capture(
  value: unknown,
  sheets?: readonly XlsxSheet[],
): XlsxParseError {
  try {
    parse(value, sheets);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected workbook-view parsing to fail');
}

describe('XLSX workbook views', () => {
  it('returns deterministic defaults using the first visible sheet', () => {
    expect(parse(undefined, [worksheet(0, 'hidden'), worksheet(1)])).toEqual([
      {
        activeSheetIndex: 1,
        autoFilterDateGrouping: true,
        firstVisibleSheetIndex: 1,
        minimized: false,
        showHorizontalScroll: true,
        showSheetTabs: true,
        showVerticalScroll: true,
        tabRatio: 600,
        visibility: 'visible',
      },
    ]);
  });

  it('normalizes authored views and numeric boundaries in order', () => {
    expect(
      parse(
        collection(
          view({
            activeTab: '1',
            autoFilterDateGrouping: '0',
            firstSheet: '2',
            minimized: 'true',
            showHorizontalScroll: 'false',
            showSheetTabs: '0',
            showVerticalScroll: '1',
            tabRatio: '0',
            visibility: 'veryHidden',
            windowHeight: '4294967295',
            windowWidth: '0',
            xWindow: '-2147483648',
            yWindow: '2147483647',
          }),
          view({ activeTab: '2', firstSheet: '1', tabRatio: '1000' }),
        ),
      ),
    ).toEqual([
      {
        activeSheetIndex: 1,
        autoFilterDateGrouping: false,
        firstVisibleSheetIndex: 2,
        minimized: true,
        showHorizontalScroll: false,
        showSheetTabs: false,
        showVerticalScroll: true,
        tabRatio: 0,
        visibility: 'very-hidden',
        windowHeight: 4_294_967_295,
        windowWidth: 0,
        xWindow: -2_147_483_648,
        yWindow: 2_147_483_647,
      },
      {
        activeSheetIndex: 2,
        autoFilterDateGrouping: true,
        firstVisibleSheetIndex: 1,
        minimized: false,
        showHorizontalScroll: true,
        showSheetTabs: true,
        showVerticalScroll: true,
        tabRatio: 1_000,
        visibility: 'visible',
      },
    ]);
    expect(
      parse(collection(view({ visibility: 'visible' })))[0]?.visibility,
    ).toBe('visible');
  });

  it.each([
    [null, 'Workbook views collection is invalid'],
    ['views', 'Workbook views collection is invalid'],
    [{}, 'Workbook views collection is empty'],
    [{ workbookView: [] }, 'Workbook views collection is empty'],
    [{ workbookView: 'view' }, 'Workbook views collection is empty'],
  ])('rejects invalid workbook views %#', (value, message) => {
    expect(capture(value).diagnostic).toMatchObject({
      code: 'invalid-document-structure',
      message,
      part: PART,
    });
  });

  it('requires at least one visible sheet', () => {
    expect(
      capture(undefined, [worksheet(0, 'hidden'), worksheet(1, 'very-hidden')])
        .diagnostic.message,
    ).toBe('Workbook must contain a visible sheet');
  });

  it.each([
    ['activeTab', '-1', 'Workbook active sheet reference is invalid'],
    ['activeTab', '01', 'Workbook active sheet reference is invalid'],
    ['activeTab', '3', 'Workbook active sheet reference is invalid'],
    ['firstSheet', '-1', 'Workbook first-visible sheet reference is invalid'],
    ['firstSheet', '3', 'Workbook first-visible sheet reference is invalid'],
  ] as const)('rejects invalid %s value %s', (attribute, value, message) => {
    expect(
      capture(collection(view({ [attribute]: value }))).diagnostic.message,
    ).toBe(message);
  });

  it('rejects hidden active and first-visible sheet references', () => {
    const sheets = [worksheet(0), worksheet(1, 'hidden'), worksheet(2)];
    expect(
      capture(collection(view({ activeTab: '1' })), sheets).diagnostic.message,
    ).toBe('Workbook active sheet reference is invalid');
    expect(
      capture(collection(view({ firstSheet: '1' })), sheets).diagnostic.message,
    ).toBe('Workbook first-visible sheet reference is invalid');
  });

  it.each([
    'autoFilterDateGrouping',
    'minimized',
    'showHorizontalScroll',
    'showSheetTabs',
    'showVerticalScroll',
  ] as const)('rejects invalid %s flag', (attribute) => {
    expect(
      capture(collection(view({ [attribute]: 'yes' }))).diagnostic.message,
    ).toContain('flag is invalid');
  });

  it.each(['bad', 'veryhidden', 'VISIBLE'])(
    'rejects visibility %s',
    (visibility) => {
      expect(capture(collection(view({ visibility }))).diagnostic.message).toBe(
        'Workbook view visibility is invalid',
      );
    },
  );

  it.each(['1001', '4294967295', '-1', '01'])(
    'rejects tab ratio %s',
    (tabRatio) => {
      expect(capture(collection(view({ tabRatio }))).diagnostic.message).toBe(
        'Workbook view tab ratio is invalid',
      );
    },
  );

  it.each([
    ['xWindow', '-2147483649', 'Workbook view horizontal position is invalid'],
    ['xWindow', '2147483648', 'Workbook view horizontal position is invalid'],
    ['xWindow', '-0', 'Workbook view horizontal position is invalid'],
    ['yWindow', '-2147483649', 'Workbook view vertical position is invalid'],
    ['yWindow', '2147483648', 'Workbook view vertical position is invalid'],
    ['yWindow', '1.0', 'Workbook view vertical position is invalid'],
  ] as const)('rejects invalid %s position %s', (attribute, value, message) => {
    expect(
      capture(collection(view({ [attribute]: value }))).diagnostic.message,
    ).toBe(message);
  });

  it.each([
    ['windowWidth', '-1', 'Workbook view width is invalid'],
    ['windowWidth', '4294967296', 'Workbook view width is invalid'],
    ['windowHeight', '01', 'Workbook view height is invalid'],
    ['windowHeight', '4294967296', 'Workbook view height is invalid'],
  ] as const)('rejects invalid %s %s', (attribute, value, message) => {
    expect(
      capture(collection(view({ [attribute]: value }))).diagnostic.message,
    ).toBe(message);
  });
});
