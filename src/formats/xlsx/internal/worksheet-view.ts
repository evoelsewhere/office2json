import { XlsxParseError } from '../errors';
import type {
  XlsxPanePosition,
  XlsxRange,
  XlsxWorksheetPane,
  XlsxWorksheetView,
  XlsxWorksheetViewSelection,
} from '../types';
import {
  parseXlsxCellReference,
  parseXlsxRangeReference,
} from './cell-reference';
import { type ResolvedXlsxResourceLimits } from './resource-limits';
import type { XlsxXmlElement } from './streaming-xml';

export interface ParsedXlsxWorksheetViewSelection {
  rangeAreaCount: number;
  selection: XlsxWorksheetViewSelection;
}

function fail(part: string, message: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function attribute(
  element: XlsxXmlElement,
  localName: string,
): string | undefined {
  return element.attributes.get(`{}${localName}`);
}

function booleanAttribute(
  element: XlsxXmlElement,
  localName: string,
  fallback: boolean,
  part: string,
  message: string,
): boolean {
  const value = attribute(element, localName);
  if (value === undefined) return fallback;
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  fail(part, message);
}

function unsignedInteger(
  value: string | undefined,
  part: string,
  message: string,
): number | undefined {
  if (value === undefined) return undefined;
  const parsed = Number(value);
  if (
    String(parsed) !== value ||
    !Number.isSafeInteger(parsed) ||
    parsed < 0 ||
    parsed > 0xffff_ffff
  ) {
    fail(part, message);
  }
  return parsed;
}

function cellReference(
  value: string | undefined,
  part: string,
  message: string,
): string | undefined {
  if (value === undefined) return undefined;
  const parsed = parseXlsxCellReference(value);
  if (!parsed || parsed.absoluteColumn || parsed.absoluteRow)
    fail(part, message);
  return parsed.address;
}

function zoomScale(
  element: XlsxXmlElement,
  localName: string,
  part: string,
): number | undefined {
  const value = unsignedInteger(
    attribute(element, localName),
    part,
    'Worksheet view zoom scale is invalid',
  );
  if (value !== undefined && (value < 10 || value > 400)) {
    fail(part, 'Worksheet view zoom scale is invalid');
  }
  return value;
}

function panePosition(
  value: string | undefined,
  part: string,
): XlsxPanePosition | undefined {
  if (value === undefined) return undefined;
  if (value === 'bottomLeft') return 'bottom-left';
  if (value === 'bottomRight') return 'bottom-right';
  if (value === 'topLeft') return 'top-left';
  if (value === 'topRight') return 'top-right';
  fail(part, 'Worksheet pane position is invalid');
}

function splitValue(value: string | undefined, part: string): number {
  if (value === undefined) return 0;
  const folded = value.toLowerCase();
  if (
    value.length === 0 ||
    value.trim() !== value ||
    value.startsWith('-') ||
    folded.includes('b') ||
    folded.includes('o') ||
    folded.includes('x')
  ) {
    fail(part, 'Worksheet pane split value is invalid');
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    fail(part, 'Worksheet pane split value is invalid');
  }
  return parsed;
}

function derivedPanePosition(xSplit: number, ySplit: number): XlsxPanePosition {
  if (xSplit > 0 && ySplit > 0) return 'bottom-right';
  if (xSplit > 0) return 'top-right';
  return 'bottom-left';
}

function viewKind(
  value: string | undefined,
  part: string,
): XlsxWorksheetView['kind'] {
  if (value === undefined || value === 'normal') return 'normal';
  if (value === 'pageBreakPreview') return 'page-break-preview';
  if (value === 'pageLayout') return 'page-layout';
  fail(part, 'Worksheet view kind is invalid');
}

function paneState(
  value: string | undefined,
  part: string,
): XlsxWorksheetPane['state'] {
  if (value === undefined || value === 'split') return 'split';
  if (value === 'frozen') return 'frozen';
  if (value === 'frozenSplit') return 'frozen-split';
  fail(part, 'Worksheet pane state is invalid');
}

export function xlsxWorksheetRangeContains(
  range: XlsxRange,
  row: number,
  column: number,
): boolean {
  return (
    row >= range.start.row &&
    row <= range.end.row &&
    column >= range.start.column &&
    column <= range.end.column
  );
}

export function parseXlsxWorksheetView(
  element: XlsxXmlElement,
  part: string,
): XlsxWorksheetView {
  const workbookViewId = unsignedInteger(
    attribute(element, 'workbookViewId'),
    part,
    'Worksheet workbook view reference is invalid',
  );
  if (workbookViewId === undefined) {
    fail(part, 'Worksheet workbook view reference is missing');
  }
  const authoredZoom = zoomScale(element, 'zoomScale', part);
  const normalZoom = zoomScale(element, 'zoomScaleNormal', part);
  const pageLayoutZoom = zoomScale(element, 'zoomScalePageLayoutView', part);
  const sheetLayoutZoom = zoomScale(element, 'zoomScaleSheetLayoutView', part);
  const topLeftCell = cellReference(
    attribute(element, 'topLeftCell'),
    part,
    'Worksheet view top-left cell is invalid',
  );
  return {
    kind: viewKind(attribute(element, 'view'), part),
    rightToLeft: booleanAttribute(
      element,
      'rightToLeft',
      false,
      part,
      'Worksheet view right-to-left flag is invalid',
    ),
    selections: [],
    showGridLines: booleanAttribute(
      element,
      'showGridLines',
      true,
      part,
      'Worksheet view gridline flag is invalid',
    ),
    showRowColumnHeaders: booleanAttribute(
      element,
      'showRowColHeaders',
      true,
      part,
      'Worksheet view header flag is invalid',
    ),
    tabSelected: booleanAttribute(
      element,
      'tabSelected',
      false,
      part,
      'Worksheet view selected-tab flag is invalid',
    ),
    ...(topLeftCell === undefined ? {} : { topLeftCell }),
    workbookViewId,
    zoomScale: authoredZoom ?? 100,
    ...(normalZoom === undefined ? {} : { zoomScaleNormal: normalZoom }),
    ...(pageLayoutZoom === undefined
      ? {}
      : { zoomScalePageLayout: pageLayoutZoom }),
    ...(sheetLayoutZoom === undefined
      ? {}
      : { zoomScaleSheetLayout: sheetLayoutZoom }),
  };
}

export function parseXlsxWorksheetPane(
  element: XlsxXmlElement,
  part: string,
  limits: ResolvedXlsxResourceLimits,
): XlsxWorksheetPane {
  const xSplit = splitValue(attribute(element, 'xSplit'), part);
  const ySplit = splitValue(attribute(element, 'ySplit'), part);
  if (xSplit === 0 && ySplit === 0) {
    fail(part, 'Worksheet pane must split at least one axis');
  }
  const state = paneState(attribute(element, 'state'), part);
  if (
    state !== 'split' &&
    (!Number.isSafeInteger(xSplit) ||
      !Number.isSafeInteger(ySplit) ||
      xSplit > limits.maxColumnsPerWorksheet ||
      ySplit > limits.maxRowsPerWorksheet)
  ) {
    fail(part, 'Worksheet frozen pane split is invalid');
  }
  const topLeftCell = cellReference(
    attribute(element, 'topLeftCell'),
    part,
    'Worksheet pane top-left cell is invalid',
  );
  return {
    activePane:
      panePosition(attribute(element, 'activePane'), part) ??
      derivedPanePosition(xSplit, ySplit),
    state,
    ...(topLeftCell === undefined ? {} : { topLeftCell }),
    xSplit,
    ySplit,
  };
}

export function parseXlsxWorksheetViewSelection(
  element: XlsxXmlElement,
  part: string,
): ParsedXlsxWorksheetViewSelection {
  const source = attribute(element, 'sqref');
  if (source === undefined || source.trim().length === 0) {
    fail(part, 'Worksheet view selection range is missing');
  }
  const ranges = source
    .trim()
    .split(/\s+/u)
    .map((reference) => {
      const range = parseXlsxRangeReference(reference);
      if (!range || reference.includes('$')) {
        fail(part, 'Worksheet view selection range is invalid');
      }
      return range;
    });
  const activeCell = cellReference(
    attribute(element, 'activeCell'),
    part,
    'Worksheet view active cell is invalid',
  );
  const activeCellId = unsignedInteger(
    attribute(element, 'activeCellId'),
    part,
    'Worksheet view active-cell index is invalid',
  );
  if (activeCellId !== undefined && activeCellId >= ranges.length) {
    fail(part, 'Worksheet view active-cell index is invalid');
  }
  if (activeCellId !== undefined && activeCell === undefined) {
    fail(part, 'Worksheet view active cell is missing');
  }
  if (activeCell !== undefined) {
    const parsed = parseXlsxCellReference(activeCell)!;
    const target = ranges[activeCellId ?? 0]!;
    if (!xlsxWorksheetRangeContains(target, parsed.row, parsed.column)) {
      fail(part, 'Worksheet view active cell is outside its selected range');
    }
  }
  return {
    rangeAreaCount: ranges.length,
    selection: {
      ...(activeCell === undefined ? {} : { activeCell }),
      ...(activeCellId === undefined ? {} : { activeCellId }),
      pane: panePosition(attribute(element, 'pane'), part) ?? 'top-left',
      ranges,
    },
  };
}

export function validateXlsxWorksheetView(
  view: XlsxWorksheetView,
  part: string,
): void {
  const panes = new Set<XlsxPanePosition>();
  for (const selection of view.selections) {
    if (panes.has(selection.pane)) {
      fail(part, 'Worksheet view contains duplicate pane selections');
    }
    panes.add(selection.pane);
    if (!view.pane && selection.pane !== 'top-left') {
      fail(part, 'Worksheet view selection references a missing pane');
    }
  }
}
