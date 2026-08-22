import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import type { XlsxXmlElement } from '../../src/formats/xlsx/internal/streaming-xml';
import {
  parseXlsxWorksheetPane,
  parseXlsxWorksheetView,
  parseXlsxWorksheetViewSelection,
  validateXlsxWorksheetView,
  xlsxWorksheetRangeContains,
} from '../../src/formats/xlsx/internal/worksheet-view';
import type { XlsxWorksheetView } from '../../src/formats/xlsx/types';

const PART = 'xl/worksheets/sheet1.xml';

function element(attributes: Record<string, string>): XlsxXmlElement {
  return {
    attributes: new Map(
      Object.entries(attributes).map(([name, value]) => [`{}${name}`, value]),
    ),
    localName: 'test',
    namespace: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  };
}

function capture(action: () => unknown): XlsxParseError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected worksheet view parsing to fail');
}

describe('XLSX worksheet views', () => {
  it('applies deterministic worksheet-view defaults', () => {
    expect(
      parseXlsxWorksheetView(element({ workbookViewId: '0' }), PART),
    ).toEqual({
      kind: 'normal',
      rightToLeft: false,
      selections: [],
      showGridLines: true,
      showRowColumnHeaders: true,
      tabSelected: false,
      workbookViewId: 0,
      zoomScale: 100,
    });
  });

  it('normalizes authored view state and exact zoom boundaries', () => {
    expect(
      parseXlsxWorksheetView(
        element({
          rightToLeft: 'true',
          showGridLines: '0',
          showRowColHeaders: 'false',
          tabSelected: '1',
          topLeftCell: 'xfd1048576',
          view: 'pageBreakPreview',
          workbookViewId: '4294967295',
          zoomScale: '10',
          zoomScaleNormal: '400',
          zoomScalePageLayoutView: '25',
          zoomScaleSheetLayoutView: '250',
        }),
        PART,
      ),
    ).toEqual({
      kind: 'page-break-preview',
      rightToLeft: true,
      selections: [],
      showGridLines: false,
      showRowColumnHeaders: false,
      tabSelected: true,
      topLeftCell: 'XFD1048576',
      workbookViewId: 4_294_967_295,
      zoomScale: 10,
      zoomScaleNormal: 400,
      zoomScalePageLayout: 25,
      zoomScaleSheetLayout: 250,
    });
    expect(
      parseXlsxWorksheetView(
        element({ view: 'pageLayout', workbookViewId: '1', zoomScale: '400' }),
        PART,
      ).kind,
    ).toBe('page-layout');
    expect(
      parseXlsxWorksheetView(
        element({ view: 'normal', workbookViewId: '2' }),
        PART,
      ).kind,
    ).toBe('normal');
  });

  it.each([
    [{}, 'Worksheet workbook view reference is missing'],
    [{ workbookViewId: '-1' }, 'Worksheet workbook view reference is invalid'],
    [
      { workbookViewId: '4294967296' },
      'Worksheet workbook view reference is invalid',
    ],
    [{ workbookViewId: '1x' }, 'Worksheet workbook view reference is invalid'],
    [{ workbookViewId: '01' }, 'Worksheet workbook view reference is invalid'],
    [{ view: 'bad', workbookViewId: '0' }, 'Worksheet view kind is invalid'],
    [
      { rightToLeft: 'yes', workbookViewId: '0' },
      'Worksheet view right-to-left flag is invalid',
    ],
    [
      { showGridLines: 'yes', workbookViewId: '0' },
      'Worksheet view gridline flag is invalid',
    ],
    [
      { showRowColHeaders: 'yes', workbookViewId: '0' },
      'Worksheet view header flag is invalid',
    ],
    [
      { tabSelected: 'yes', workbookViewId: '0' },
      'Worksheet view selected-tab flag is invalid',
    ],
    [
      { topLeftCell: '$A$1', workbookViewId: '0' },
      'Worksheet view top-left cell is invalid',
    ],
    [
      { workbookViewId: '0', zoomScale: '9' },
      'Worksheet view zoom scale is invalid',
    ],
    [
      { workbookViewId: '0', zoomScale: '401' },
      'Worksheet view zoom scale is invalid',
    ],
    [
      { workbookViewId: '0', zoomScale: '1.5' },
      'Worksheet view zoom scale is invalid',
    ],
  ] as const)('rejects invalid worksheet view %#', (attributes, message) => {
    expect(
      capture(() => parseXlsxWorksheetView(element({ ...attributes }), PART))
        .diagnostic,
    ).toMatchObject({ code: 'invalid-document-value', message, part: PART });
  });

  it('normalizes frozen and split panes', () => {
    const limits = defaultXlsxResourceLimits();
    expect(
      parseXlsxWorksheetPane(
        element({
          activePane: 'bottomRight',
          state: 'frozen',
          topLeftCell: 'b3',
          xSplit: '1',
          ySplit: '2',
        }),
        PART,
        limits,
      ),
    ).toEqual({
      activePane: 'bottom-right',
      state: 'frozen',
      topLeftCell: 'B3',
      xSplit: 1,
      ySplit: 2,
    });
    expect(
      parseXlsxWorksheetPane(
        element({ state: 'split', xSplit: '1.5' }),
        PART,
        limits,
      ),
    ).toEqual({
      activePane: 'top-right',
      state: 'split',
      xSplit: 1.5,
      ySplit: 0,
    });
    expect(
      parseXlsxWorksheetPane(
        element({ state: 'frozenSplit', ySplit: '1' }),
        PART,
        limits,
      ).activePane,
    ).toBe('bottom-left');
    expect(
      parseXlsxWorksheetPane(
        element({ xSplit: '12', ySplit: '.55' }),
        PART,
        limits,
      ).activePane,
    ).toBe('bottom-right');
    expect(
      parseXlsxWorksheetPane(
        element({ activePane: 'topLeft', xSplit: '1.' }),
        PART,
        limits,
      ).activePane,
    ).toBe('top-left');
  });

  it('accepts frozen splits exactly at configured grid limits', () => {
    expect(
      parseXlsxWorksheetPane(
        element({ state: 'frozen', xSplit: '2', ySplit: '3' }),
        PART,
        {
          ...defaultXlsxResourceLimits(),
          maxColumnsPerWorksheet: 2,
          maxRowsPerWorksheet: 3,
        },
      ),
    ).toMatchObject({ xSplit: 2, ySplit: 3 });
  });

  it.each([
    [{}, 'Worksheet pane must split at least one axis'],
    [{ activePane: 'bad', xSplit: '1' }, 'Worksheet pane position is invalid'],
    [{ state: 'bad', xSplit: '1' }, 'Worksheet pane state is invalid'],
    [{ xSplit: '-1' }, 'Worksheet pane split value is invalid'],
    [{ xSplit: '' }, 'Worksheet pane split value is invalid'],
    [{ xSplit: '1 ' }, 'Worksheet pane split value is invalid'],
    [{ xSplit: '1e309' }, 'Worksheet pane split value is invalid'],
    [{ xSplit: '1x' }, 'Worksheet pane split value is invalid'],
    [{ xSplit: '0b10' }, 'Worksheet pane split value is invalid'],
    [{ xSplit: '0o10' }, 'Worksheet pane split value is invalid'],
    [{ xSplit: '0x10' }, 'Worksheet pane split value is invalid'],
    [
      { state: 'frozen', xSplit: '1.5' },
      'Worksheet frozen pane split is invalid',
    ],
    [
      { state: 'frozen', xSplit: '3' },
      'Worksheet frozen pane split is invalid',
    ],
    [
      { state: 'frozen', ySplit: '4' },
      'Worksheet frozen pane split is invalid',
    ],
    [
      { topLeftCell: '$A$1', xSplit: '1' },
      'Worksheet pane top-left cell is invalid',
    ],
  ] as const)('rejects invalid worksheet pane %#', (attributes, message) => {
    expect(
      capture(() =>
        parseXlsxWorksheetPane(element({ ...attributes }), PART, {
          ...defaultXlsxResourceLimits(),
          maxColumnsPerWorksheet: 2,
          maxRowsPerWorksheet: 3,
        }),
      ).diagnostic.message,
    ).toBe(message);
  });

  it('normalizes pane selection areas and active-cell identity', () => {
    expect(
      parseXlsxWorksheetViewSelection(
        element({
          activeCell: 'c3',
          activeCellId: '1',
          pane: 'bottomRight',
          sqref: 'A1:B2 C3:D4',
        }),
        PART,
      ),
    ).toEqual({
      rangeAreaCount: 2,
      selection: {
        activeCell: 'C3',
        activeCellId: 1,
        pane: 'bottom-right',
        ranges: [
          {
            end: { column: 2, row: 2 },
            reference: 'A1:B2',
            start: { column: 1, row: 1 },
          },
          {
            end: { column: 4, row: 4 },
            reference: 'C3:D4',
            start: { column: 3, row: 3 },
          },
        ],
      },
    });
    expect(
      parseXlsxWorksheetViewSelection(element({ sqref: 'XFD1048576' }), PART),
    ).toMatchObject({
      rangeAreaCount: 1,
      selection: { pane: 'top-left' },
    });
    expect(
      parseXlsxWorksheetViewSelection(
        element({ activeCell: 'B2', sqref: '  B2:C3   D4  ' }),
        PART,
      ),
    ).toMatchObject({
      rangeAreaCount: 2,
      selection: { activeCell: 'B2' },
    });
    expect(
      parseXlsxWorksheetViewSelection(
        element({ activeCell: 'C3', sqref: 'B2:C3' }),
        PART,
      ).selection.activeCell,
    ).toBe('C3');
  });

  it.each([
    [1, 2, false],
    [4, 2, false],
    [2, 1, false],
    [2, 4, false],
    [2, 2, true],
    [3, 3, true],
  ] as const)(
    'classifies selected-range containment at row %i column %i as %s',
    (row, column, expected) => {
      expect(
        xlsxWorksheetRangeContains(
          {
            end: { column: 3, row: 3 },
            reference: 'B2:C3',
            start: { column: 2, row: 2 },
          },
          row,
          column,
        ),
      ).toBe(expected);
    },
  );

  it.each([
    [{}, 'Worksheet view selection range is missing'],
    [{ sqref: '  ' }, 'Worksheet view selection range is missing'],
    [{ sqref: 'B2:A1' }, 'Worksheet view selection range is invalid'],
    [{ sqref: '$A$1' }, 'Worksheet view selection range is invalid'],
    [
      { activeCell: '$A$1', sqref: 'A1' },
      'Worksheet view active cell is invalid',
    ],
    [
      { activeCell: 'A1', activeCellId: '-1', sqref: 'A1' },
      'Worksheet view active-cell index is invalid',
    ],
    [
      { activeCell: 'A1', activeCellId: '1', sqref: 'A1' },
      'Worksheet view active-cell index is invalid',
    ],
    [
      { activeCellId: '0', sqref: 'A1' },
      'Worksheet view active cell is missing',
    ],
    [
      { activeCell: 'B2', sqref: 'A1' },
      'Worksheet view active cell is outside its selected range',
    ],
    [
      { activeCell: 'B1', sqref: 'B2:C3' },
      'Worksheet view active cell is outside its selected range',
    ],
    [
      { activeCell: 'B4', sqref: 'B2:C3' },
      'Worksheet view active cell is outside its selected range',
    ],
    [
      { activeCell: 'A2', sqref: 'B2:C3' },
      'Worksheet view active cell is outside its selected range',
    ],
    [
      { activeCell: 'D2', sqref: 'B2:C3' },
      'Worksheet view active cell is outside its selected range',
    ],
    [{ pane: 'bad', sqref: 'A1' }, 'Worksheet pane position is invalid'],
  ] as const)('rejects invalid view selection %#', (attributes, message) => {
    expect(
      capture(() =>
        parseXlsxWorksheetViewSelection(element({ ...attributes }), PART),
      ).diagnostic.message,
    ).toBe(message);
  });

  it('validates one selection per existing pane', () => {
    const view = parseXlsxWorksheetView(element({ workbookViewId: '0' }), PART);
    view.pane = parseXlsxWorksheetPane(
      element({ xSplit: '1' }),
      PART,
      defaultXlsxResourceLimits(),
    );
    view.selections.push(
      parseXlsxWorksheetViewSelection(
        element({ pane: 'topRight', sqref: 'B1' }),
        PART,
      ).selection,
    );
    expect(() => validateXlsxWorksheetView(view, PART)).not.toThrow();
    view.selections.push({ ...view.selections[0]!, ranges: [] });
    expect(
      capture(() => validateXlsxWorksheetView(view, PART)).diagnostic.message,
    ).toBe('Worksheet view contains duplicate pane selections');
  });

  it('rejects a pane selection when no pane exists', () => {
    const view: XlsxWorksheetView = {
      ...parseXlsxWorksheetView(element({ workbookViewId: '0' }), PART),
      selections: [
        parseXlsxWorksheetViewSelection(
          element({ pane: 'bottomLeft', sqref: 'A1' }),
          PART,
        ).selection,
      ],
    };
    expect(
      capture(() => validateXlsxWorksheetView(view, PART)).diagnostic.message,
    ).toBe('Worksheet view selection references a missing pane');
  });
});
