import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import type { XlsxXmlElement } from '../../src/formats/xlsx/internal/streaming-xml';
import {
  parseXlsxDeclaredDimension,
  parseXlsxWorksheetFormat,
  parseXlsxWorksheetOutline,
  parseXlsxWorksheetTabColor,
} from '../../src/formats/xlsx/internal/worksheet-metadata';

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
  throw new Error('Expected worksheet metadata parsing to fail');
}

describe('XLSX worksheet metadata', () => {
  it.each([
    [
      'A1',
      {
        end: { column: 1, row: 1 },
        reference: 'A1',
        start: { column: 1, row: 1 },
      },
    ],
    [
      'b2:xfd1048576',
      {
        end: { column: 16_384, row: 1_048_576 },
        reference: 'B2:XFD1048576',
        start: { column: 2, row: 2 },
      },
    ],
  ] as const)('normalizes declared dimension %s', (reference, expected) => {
    expect(
      parseXlsxDeclaredDimension(element({ ref: reference }), PART),
    ).toEqual(expected);
  });

  it.each([undefined, '', '$A$1', 'B2:A1', 'XFE1'])(
    'rejects declared dimension %#',
    (ref) => {
      expect(
        capture(() =>
          parseXlsxDeclaredDimension(
            element(ref === undefined ? {} : { ref }),
            PART,
          ),
        ).diagnostic.message,
      ).toBe('Worksheet declared dimension is invalid');
    },
  );

  it('normalizes sheet format defaults and exact boundaries', () => {
    expect(
      parseXlsxWorksheetFormat(element({ defaultRowHeight: '15' }), PART),
    ).toEqual({
      baseColumnWidth: 8,
      customHeight: false,
      defaultRowHeight: 15,
      outlineColumnLevel: 0,
      outlineRowLevel: 0,
      thickBottom: false,
      thickTop: false,
      zeroHeight: false,
    });
    expect(
      parseXlsxWorksheetFormat(
        element({
          baseColWidth: '255',
          customHeight: '1',
          defaultColWidth: '255',
          defaultRowHeight: '409',
          outlineLevelCol: '7',
          outlineLevelRow: '7',
          thickBottom: 'true',
          thickTop: '1',
          zeroHeight: 'true',
        }),
        PART,
      ),
    ).toEqual({
      baseColumnWidth: 255,
      customHeight: true,
      defaultColumnWidth: 255,
      defaultRowHeight: 409,
      outlineColumnLevel: 7,
      outlineRowLevel: 7,
      thickBottom: true,
      thickTop: true,
      zeroHeight: true,
    });
    expect(
      parseXlsxWorksheetFormat(
        element({
          baseColWidth: '0',
          defaultColWidth: '.5',
          defaultRowHeight: '1.',
        }),
        PART,
      ),
    ).toMatchObject({
      baseColumnWidth: 0,
      defaultColumnWidth: 0.5,
      defaultRowHeight: 1,
    });
  });

  it.each([
    [{}, 'Worksheet default row height is invalid'],
    [{ defaultRowHeight: '-1' }, 'Worksheet default row height is invalid'],
    [{ defaultRowHeight: '' }, 'Worksheet default row height is invalid'],
    [{ defaultRowHeight: '1 ' }, 'Worksheet default row height is invalid'],
    [{ defaultRowHeight: '0b10' }, 'Worksheet default row height is invalid'],
    [{ defaultRowHeight: '0o10' }, 'Worksheet default row height is invalid'],
    [{ defaultRowHeight: '0x10' }, 'Worksheet default row height is invalid'],
    [{ defaultRowHeight: '410' }, 'Worksheet default row height is invalid'],
    [{ defaultRowHeight: '1e309' }, 'Worksheet default row height is invalid'],
    [
      { defaultColWidth: '256', defaultRowHeight: '15' },
      'Worksheet default column width is invalid',
    ],
    [
      { defaultColWidth: '-1', defaultRowHeight: '15' },
      'Worksheet default column width is invalid',
    ],
    [
      { baseColWidth: '256', defaultRowHeight: '15' },
      'Worksheet base column width is invalid',
    ],
    [
      { baseColWidth: '-1', defaultRowHeight: '15' },
      'Worksheet base column width is invalid',
    ],
    [
      { baseColWidth: '01', defaultRowHeight: '15' },
      'Worksheet base column width is invalid',
    ],
    [
      { defaultRowHeight: '15', outlineLevelCol: '8' },
      'Worksheet column outline level is invalid',
    ],
    [
      { defaultRowHeight: '15', outlineLevelCol: '-1' },
      'Worksheet column outline level is invalid',
    ],
    [
      { defaultRowHeight: '15', outlineLevelRow: '8' },
      'Worksheet row outline level is invalid',
    ],
    [
      { defaultRowHeight: '15', outlineLevelRow: '-1' },
      'Worksheet row outline level is invalid',
    ],
    [
      { customHeight: 'yes', defaultRowHeight: '15' },
      'Worksheet custom-height flag is invalid',
    ],
    [
      { defaultRowHeight: '15', thickBottom: 'yes' },
      'Worksheet thick-bottom flag is invalid',
    ],
    [
      { defaultRowHeight: '15', thickTop: 'yes' },
      'Worksheet thick-top flag is invalid',
    ],
    [
      { defaultRowHeight: '15', zeroHeight: 'yes' },
      'Worksheet zero-height flag is invalid',
    ],
  ] as const)('rejects invalid sheet format %#', (attributes, message) => {
    expect(
      capture(() => parseXlsxWorksheetFormat(element({ ...attributes }), PART))
        .diagnostic.message,
    ).toBe(message);
  });

  it('normalizes outline defaults and authored flags', () => {
    expect(parseXlsxWorksheetOutline(element({}), PART)).toEqual({
      applyStyles: false,
      showOutlineSymbols: true,
      summaryBelow: true,
      summaryRight: true,
    });
    expect(
      parseXlsxWorksheetOutline(
        element({
          applyStyles: 'true',
          showOutlineSymbols: '0',
          summaryBelow: 'false',
          summaryRight: '0',
        }),
        PART,
      ),
    ).toEqual({
      applyStyles: true,
      showOutlineSymbols: false,
      summaryBelow: false,
      summaryRight: false,
    });
  });

  it.each([
    ['applyStyles', 'Worksheet outline apply-styles flag is invalid'],
    ['showOutlineSymbols', 'Worksheet outline-symbol flag is invalid'],
    ['summaryBelow', 'Worksheet outline summary-below flag is invalid'],
    ['summaryRight', 'Worksheet outline summary-right flag is invalid'],
  ] as const)('rejects invalid outline flag %s', (attribute, message) => {
    expect(
      capture(() =>
        parseXlsxWorksheetOutline(element({ [attribute]: 'yes' }), PART),
      ).diagnostic.message,
    ).toBe(message);
  });

  it('normalizes worksheet tab colors without exposing XML attributes', () => {
    expect(
      parseXlsxWorksheetTabColor(element({ rgb: 'ff00aabb' }), PART),
    ).toEqual({ argb: 'FF00AABB', kind: 'rgb' });
    expect(
      parseXlsxWorksheetTabColor(element({ theme: '2', tint: '-.5' }), PART),
    ).toEqual({ index: 2, kind: 'theme', tint: -0.5 });
  });

  it('rejects a missing worksheet tab color selector', () => {
    expect(
      capture(() => parseXlsxWorksheetTabColor(element({}), PART)).diagnostic
        .message,
    ).toBe('Worksheet tab color is missing');
    expect(
      capture(() =>
        parseXlsxWorksheetTabColor(element({ rgb: 'FFFFFF' }), PART),
      ).diagnostic.message,
    ).toBe('Worksheet tab color RGB is invalid');
  });
});
