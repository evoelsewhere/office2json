import { XlsxParseError } from '../errors';
import type { XlsxSheet, XlsxWorkbookView } from '../types';

type XmlRecord = Record<string, unknown>;

function fail(part: string, message: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function structureFailure(part: string, message: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-structure',
    message,
    part,
    severity: 'error',
  });
}

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function records(value: unknown): XmlRecord[] | undefined {
  const values = Array.isArray(value) ? value : [value];
  const output: XmlRecord[] = [];
  for (const item of values) {
    const parsed = record(item);
    if (!parsed) return undefined;
    output.push(parsed);
  }
  return output;
}

function attributes(value: XmlRecord): XmlRecord {
  return record(value.attrs) ?? {};
}

function booleanAttribute(
  value: unknown,
  fallback: boolean,
  part: string,
  message: string,
): boolean {
  if (value === undefined) return fallback;
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  fail(part, message);
}

function unsignedInteger(
  value: unknown,
  fallback: number,
  part: string,
  message: string,
): number {
  if (value === undefined) return fallback;
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

function signedInteger(
  value: unknown,
  part: string,
  message: string,
): number | undefined {
  if (value === undefined) return undefined;
  const parsed = Number(value);
  if (
    String(parsed) !== value ||
    !Number.isSafeInteger(parsed) ||
    parsed < -0x8000_0000 ||
    parsed > 0x7fff_ffff
  ) {
    fail(part, message);
  }
  return parsed;
}

function visibility(
  value: unknown,
  part: string,
): XlsxWorkbookView['visibility'] {
  if (value === undefined || value === 'visible') return 'visible';
  if (value === 'hidden') return 'hidden';
  if (value === 'veryHidden') return 'very-hidden';
  fail(part, 'Workbook view visibility is invalid');
}

function firstVisibleSheet(sheets: readonly XlsxSheet[], part: string): number {
  const index = sheets.findIndex((sheet) => sheet.state === 'visible');
  if (index < 0) fail(part, 'Workbook must contain a visible sheet');
  return index;
}

function sheetIndex(
  value: unknown,
  fallback: number,
  sheets: readonly XlsxSheet[],
  part: string,
  message: string,
): number {
  const index = unsignedInteger(value, fallback, part, message);
  if (index >= sheets.length || sheets[index]!.state !== 'visible') {
    fail(part, message);
  }
  return index;
}

function parseView(
  node: XmlRecord,
  sheets: readonly XlsxSheet[],
  fallbackSheet: number,
  part: string,
): XlsxWorkbookView {
  const attrs = attributes(node);
  const tabRatio = unsignedInteger(
    attrs.tabRatio,
    600,
    part,
    'Workbook view tab ratio is invalid',
  );
  if (tabRatio > 1_000) fail(part, 'Workbook view tab ratio is invalid');
  const xWindow = signedInteger(
    attrs.xWindow,
    part,
    'Workbook view horizontal position is invalid',
  );
  const yWindow = signedInteger(
    attrs.yWindow,
    part,
    'Workbook view vertical position is invalid',
  );
  const windowWidth =
    attrs.windowWidth === undefined
      ? undefined
      : unsignedInteger(
          attrs.windowWidth,
          0,
          part,
          'Workbook view width is invalid',
        );
  const windowHeight =
    attrs.windowHeight === undefined
      ? undefined
      : unsignedInteger(
          attrs.windowHeight,
          0,
          part,
          'Workbook view height is invalid',
        );
  return {
    activeSheetIndex: sheetIndex(
      attrs.activeTab,
      fallbackSheet,
      sheets,
      part,
      'Workbook active sheet reference is invalid',
    ),
    autoFilterDateGrouping: booleanAttribute(
      attrs.autoFilterDateGrouping,
      true,
      part,
      'Workbook view date-grouping flag is invalid',
    ),
    firstVisibleSheetIndex: sheetIndex(
      attrs.firstSheet,
      fallbackSheet,
      sheets,
      part,
      'Workbook first-visible sheet reference is invalid',
    ),
    minimized: booleanAttribute(
      attrs.minimized,
      false,
      part,
      'Workbook view minimized flag is invalid',
    ),
    showHorizontalScroll: booleanAttribute(
      attrs.showHorizontalScroll,
      true,
      part,
      'Workbook view horizontal-scroll flag is invalid',
    ),
    showSheetTabs: booleanAttribute(
      attrs.showSheetTabs,
      true,
      part,
      'Workbook view sheet-tab flag is invalid',
    ),
    showVerticalScroll: booleanAttribute(
      attrs.showVerticalScroll,
      true,
      part,
      'Workbook view vertical-scroll flag is invalid',
    ),
    tabRatio,
    visibility: visibility(attrs.visibility, part),
    ...(windowHeight === undefined ? {} : { windowHeight }),
    ...(windowWidth === undefined ? {} : { windowWidth }),
    ...(xWindow === undefined ? {} : { xWindow }),
    ...(yWindow === undefined ? {} : { yWindow }),
  };
}

export function parseXlsxWorkbookViews(
  value: unknown,
  prefix: string,
  part: string,
  sheets: readonly XlsxSheet[],
): XlsxWorkbookView[] {
  const fallbackSheet = firstVisibleSheet(sheets, part);
  if (value === undefined) {
    return [parseView({}, sheets, fallbackSheet, part)];
  }
  const container = record(value);
  if (!container) {
    structureFailure(part, 'Workbook views collection is invalid');
  }
  const nodes = records(
    container[prefix ? `${prefix}:workbookView` : 'workbookView'],
  );
  if (!nodes || nodes.length === 0) {
    structureFailure(part, 'Workbook views collection is empty');
  }
  return nodes.map((node) => parseView(node, sheets, fallbackSheet, part));
}
