import { XlsxParseError } from '../errors';
import type {
  XlsxColor,
  XlsxRange,
  XlsxWorksheetFormat,
  XlsxWorksheetOutline,
} from '../types';
import { parseXlsxRangeReference } from './cell-reference';
import { parseXlsxStyleColor } from './style-color';
import type { XlsxXmlElement } from './streaming-xml';

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
  fallback: number,
  part: string,
  message: string,
): number {
  if (value === undefined) return fallback;
  const parsed = Number(value);
  if (String(parsed) !== value || !Number.isSafeInteger(parsed) || parsed < 0) {
    fail(part, message);
  }
  return parsed;
}

function nonnegativeNumber(
  value: string | undefined,
  part: string,
  message: string,
): number | undefined {
  if (value === undefined) return undefined;
  const folded = value.toLowerCase();
  const parsed = Number(value);
  if (
    value.length === 0 ||
    value.trim() !== value ||
    value.startsWith('-') ||
    folded.includes('b') ||
    folded.includes('o') ||
    folded.includes('x') ||
    !Number.isFinite(parsed)
  ) {
    fail(part, message);
  }
  return parsed;
}

export function parseXlsxDeclaredDimension(
  element: XlsxXmlElement,
  part: string,
): XlsxRange {
  const source = attribute(element, 'ref');
  const range = parseXlsxRangeReference(source);
  if (!range || source?.includes('$')) {
    fail(part, 'Worksheet declared dimension is invalid');
  }
  return range;
}

export function parseXlsxWorksheetFormat(
  element: XlsxXmlElement,
  part: string,
): XlsxWorksheetFormat {
  const defaultRowHeight = nonnegativeNumber(
    attribute(element, 'defaultRowHeight'),
    part,
    'Worksheet default row height is invalid',
  );
  if (defaultRowHeight === undefined || defaultRowHeight > 409) {
    fail(part, 'Worksheet default row height is invalid');
  }
  const defaultColumnWidth = nonnegativeNumber(
    attribute(element, 'defaultColWidth'),
    part,
    'Worksheet default column width is invalid',
  );
  if (defaultColumnWidth !== undefined && defaultColumnWidth > 255) {
    fail(part, 'Worksheet default column width is invalid');
  }
  const baseColumnWidth = unsignedInteger(
    attribute(element, 'baseColWidth'),
    8,
    part,
    'Worksheet base column width is invalid',
  );
  if (baseColumnWidth > 255) {
    fail(part, 'Worksheet base column width is invalid');
  }
  const outlineColumnLevel = unsignedInteger(
    attribute(element, 'outlineLevelCol'),
    0,
    part,
    'Worksheet column outline level is invalid',
  );
  const outlineRowLevel = unsignedInteger(
    attribute(element, 'outlineLevelRow'),
    0,
    part,
    'Worksheet row outline level is invalid',
  );
  if (outlineColumnLevel > 7) {
    fail(part, 'Worksheet column outline level is invalid');
  }
  if (outlineRowLevel > 7) {
    fail(part, 'Worksheet row outline level is invalid');
  }
  return {
    baseColumnWidth,
    customHeight: booleanAttribute(
      element,
      'customHeight',
      false,
      part,
      'Worksheet custom-height flag is invalid',
    ),
    ...(defaultColumnWidth === undefined ? {} : { defaultColumnWidth }),
    defaultRowHeight,
    outlineColumnLevel,
    outlineRowLevel,
    thickBottom: booleanAttribute(
      element,
      'thickBottom',
      false,
      part,
      'Worksheet thick-bottom flag is invalid',
    ),
    thickTop: booleanAttribute(
      element,
      'thickTop',
      false,
      part,
      'Worksheet thick-top flag is invalid',
    ),
    zeroHeight: booleanAttribute(
      element,
      'zeroHeight',
      false,
      part,
      'Worksheet zero-height flag is invalid',
    ),
  };
}

export function parseXlsxWorksheetOutline(
  element: XlsxXmlElement,
  part: string,
): XlsxWorksheetOutline {
  return {
    applyStyles: booleanAttribute(
      element,
      'applyStyles',
      false,
      part,
      'Worksheet outline apply-styles flag is invalid',
    ),
    showOutlineSymbols: booleanAttribute(
      element,
      'showOutlineSymbols',
      true,
      part,
      'Worksheet outline-symbol flag is invalid',
    ),
    summaryBelow: booleanAttribute(
      element,
      'summaryBelow',
      true,
      part,
      'Worksheet outline summary-below flag is invalid',
    ),
    summaryRight: booleanAttribute(
      element,
      'summaryRight',
      true,
      part,
      'Worksheet outline summary-right flag is invalid',
    ),
  };
}

export function parseXlsxWorksheetTabColor(
  element: XlsxXmlElement,
  part: string,
): XlsxColor {
  const attrs = Object.fromEntries(
    [...element.attributes.entries()].map(([name, value]) => [
      name.slice(2),
      value,
    ]),
  );
  const color = parseXlsxStyleColor({ attrs }, part, 'Worksheet tab');
  if (!color) fail(part, 'Worksheet tab color is missing');
  return color;
}
