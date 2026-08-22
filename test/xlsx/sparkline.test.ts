import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  parseXlsxWithDiagnostics,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
  XlsxParseError,
} from '../../src/formats/xlsx';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import { XlsxWorksheetExtensionsCapture } from '../../src/formats/xlsx/internal/sparkline';
import type { XlsxXmlElement } from '../../src/formats/xlsx/internal/streaming-xml';
import { createXlsxWorksheetBudget } from '../../src/formats/xlsx/internal/worksheet';
import {
  createIndependentXlsx,
  type XlsxBlackBoxOverrides,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';
const SPARKLINE_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';

function element(
  localName: string,
  namespace = XLSX_SPREADSHEET_NS,
): XlsxXmlElement {
  return { attributes: new Map(), localName, namespace };
}

function captureSync(action: () => unknown): XlsxParseError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected sparkline capture to fail');
}

function extensionCapture(): XlsxWorksheetExtensionsCapture {
  return new XlsxWorksheetExtensionsCapture(
    XLSX_SPREADSHEET_NS,
    { kind: 'full-sheet' },
    createXlsxWorksheetBudget({ part: null, values: [] }),
    defaultXlsxResourceLimits(),
    'xl/worksheets/sheet1.xml',
  );
}

function worksheet(extensionPayload: string): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:x14="${X14_NS}" xmlns:xm="${XM_NS}"><sheetData/><extLst>${extensionPayload}</extLst></worksheet>`;
}

function sparkline(formula: string, location: string): string {
  return `<x14:sparkline><xm:f>${formula}</xm:f><xm:sqref>${location}</xm:sqref></x14:sparkline>`;
}

function sparklineExtension(
  groupAttributes = 'manualMax="100" manualMin="-5" lineWeight="0.75" type="column" dateAxis="1" displayEmptyCellsAs="span" markers="true" high="1" low="0" first="true" last="false" negative="1" displayXAxis="true" displayHidden="1" rightToLeft="0" minAxisType="custom" maxAxisType="custom"',
  sparklines = `${sparkline('Sheet1!A1:A3', 'B1')}${sparkline('Sheet1!C1:C3', 'B2')}`,
  colors = '<x14:colorSeries rgb="FF112233"/><x14:colorNegative theme="2" tint="0.25"/><x14:colorAxis indexed="64"/><x14:colorMarkers rgb="FF445566"/><x14:colorFirst rgb="FF778899"/><x14:colorLast rgb="FFAABBCC"/><x14:colorHigh rgb="FFDDEEFF"/><x14:colorLow auto="1"/>',
): string {
  return `<ext uri="${SPARKLINE_URI}"><x14:sparklineGroups><x14:sparklineGroup ${groupAttributes}>${colors}<x14:sparklines>${sparklines}</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`;
}

function overrides(
  extensionPayload = sparklineExtension(),
): XlsxBlackBoxOverrides {
  return { 'xl/worksheets/sheet1.xml': worksheet(extensionPayload) };
}

async function capture(
  changes: XlsxBlackBoxOverrides,
  options: Parameters<typeof parseXlsx>[1] = { errorMode: 'strict' },
): Promise<XlsxParseError> {
  try {
    await parseXlsx(await createIndependentXlsx(changes), options);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected XLSX sparkline parsing to fail');
}

describe('XLSX sparklines', () => {
  it('parses x14 groups, ranges, axis options, flags, and colors', async () => {
    const document = await parseXlsx(await createIndependentXlsx(overrides()), {
      errorMode: 'strict',
    });
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    expect(
      sheet.kind === 'worksheet' ? (sheet.sparklineGroups ?? []) : [],
    ).toStrictEqual([
      {
        colors: {
          axis: { index: 64, kind: 'indexed' },
          first: { argb: 'FF778899', kind: 'rgb' },
          high: { argb: 'FFDDEEFF', kind: 'rgb' },
          last: { argb: 'FFAABBCC', kind: 'rgb' },
          low: { kind: 'automatic' },
          markers: { argb: 'FF445566', kind: 'rgb' },
          negative: { index: 2, kind: 'theme', tint: 0.25 },
          series: { argb: 'FF112233', kind: 'rgb' },
        },
        dateAxis: true,
        displayEmptyCellsAs: 'span',
        displayHidden: true,
        displayXAxis: true,
        first: true,
        high: true,
        last: false,
        lineWeight: 0.75,
        low: false,
        manualMaximum: 100,
        manualMinimum: -5,
        markers: true,
        maximumAxisType: 'custom',
        minimumAxisType: 'custom',
        negative: true,
        rightToLeft: false,
        sparklines: [
          {
            dataFormula: 'Sheet1!A1:A3',
            location: 'B1',
            selectionRelation: 'full-sheet',
          },
          {
            dataFormula: 'Sheet1!C1:C3',
            location: 'B2',
            selectionRelation: 'full-sheet',
          },
        ],
        type: 'column',
      },
    ]);
    expect(JSON.parse(JSON.stringify(document))).toEqual(document);
  });

  it('preserves defaults and filters sparkline locations by selection', async () => {
    const source = sparklineExtension(
      '',
      `${sparkline('A1:A3', 'B1')}${sparkline('C1:C3', 'B2')}`,
      '',
    );
    const document = await parseXlsx(
      await createIndependentXlsx(overrides(source)),
      {
        errorMode: 'strict',
        selection: { ranges: { Sheet1: ['B2'] } },
      },
    );
    const sheet = document.sheets[0]!;
    const groups =
      sheet.kind === 'worksheet' ? (sheet.sparklineGroups ?? []) : [];
    expect(groups).toStrictEqual([
      {
        colors: {},
        dateAxis: false,
        displayEmptyCellsAs: 'zero',
        displayHidden: false,
        displayXAxis: false,
        first: false,
        high: false,
        last: false,
        low: false,
        markers: false,
        maximumAxisType: 'individual',
        minimumAxisType: 'individual',
        negative: false,
        rightToLeft: false,
        sparklines: [
          {
            dataFormula: 'C1:C3',
            location: 'B2',
            selectionRelation: 'intersects-selection',
          },
        ],
        type: 'line',
      },
    ]);
  });

  it('diagnoses unrelated extension payloads without confusing nested names', async () => {
    const unrelated = `<ext uri="urn:other"><foreign xmlns="urn:foreign"><sparklineGroups><sparklineGroup/></sparklineGroups></foreign></ext>`;
    const source = await createIndependentXlsx(
      overrides(`${unrelated}${sparklineExtension()}`),
    );
    const result = await parseXlsxWithDiagnostics(source);
    expect(result.diagnostics).toStrictEqual([
      {
        code: 'unsupported-feature',
        message: 'Worksheet extension content was omitted',
        part: 'xl/worksheets/sheet1.xml',
        severity: 'warning',
        sheet: 'Sheet1',
      },
    ]);
    expect(JSON.stringify(result)).not.toContain('urn:other');
    const sheet = result.document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? (sheet.sparklineGroups ?? []) : [],
    ).toHaveLength(1);
    await expect(
      parseXlsx(source, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'unsupported-feature',
        message: 'Worksheet extension content was omitted',
        severity: 'error',
      },
    });
  });

  it('round-trips sparkline metadata through portable exact R0', async () => {
    const source = await createIndependentXlsx(overrides());
    const snapshot = await readXlsxRoundTrip(source);
    const output = await writeXlsxRoundTrip(
      await validateXlsxRoundTripJson(
        JSON.parse(JSON.stringify(snapshot)) as unknown,
      ),
    );
    expect(output.data).toEqual(source);
    expect(output.report.level).toBe('R0');
  });

  it('enforces sparkline formula and location budgets exactly', async () => {
    const one = sparklineExtension('', sparkline('A', 'B1'));
    await expect(
      parseXlsx(await createIndependentXlsx(overrides(one)), {
        errorMode: 'strict',
        limits: { maxFormulaCharacters: 1, maxRangeAreas: 1 },
      }),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(
          overrides(sparklineExtension('', sparkline('AA', 'B1'))),
          {
            errorMode: 'strict',
            limits: { maxFormulaCharacters: 1 },
          },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxFormulaCharacters',
    });
    expect(
      (
        await capture(overrides(), {
          errorMode: 'strict',
          limits: { maxRangeAreas: 1 },
        })
      ).diagnostic,
    ).toMatchObject({ actual: 2, limit: 1, limitName: 'maxRangeAreas' });
  });

  it.each([
    ['0', 0],
    ['.55', 0.55],
    ['1.', 1],
    ['1e30', 1e30],
    ['1e+30', 1e30],
    ['-0', 0],
  ] as const)(
    'parses sparkline numeric lexical form %s',
    async (source, expected) => {
      const document = await parseXlsx(
        await createIndependentXlsx(
          overrides(sparklineExtension(`lineWeight="${source}"`)),
        ),
        { errorMode: 'strict' },
      );
      const sheet = document.sheets[0]!;
      const groups =
        sheet.kind === 'worksheet' ? (sheet.sparklineGroups ?? []) : [];
      expect(groups[0]?.lineWeight).toBe(expected);
    },
  );

  it('ignores foreign lookalikes inside a known sparkline group', async () => {
    const source = sparklineExtension(
      '',
      `<x14:sparkline><f xmlns="urn:foreign">ignored</f><xm:f>A1:A3</xm:f><sqref xmlns="urn:foreign">ignored</sqref><xm:sqref>B1</xm:sqref></x14:sparkline>`,
      '',
    );
    const document = await parseXlsx(
      await createIndependentXlsx(overrides(source)),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const groups =
      sheet.kind === 'worksheet' ? (sheet.sparklineGroups ?? []) : [];
    expect(groups[0]?.sparklines).toStrictEqual([
      {
        dataFormula: 'A1:A3',
        location: 'B1',
        selectionRelation: 'full-sheet',
      },
    ]);
  });

  it('omits a group when selection excludes every sparkline location', async () => {
    const document = await parseXlsx(await createIndependentXlsx(overrides()), {
      errorMode: 'strict',
      selection: { ranges: { Sheet1: ['C10'] } },
    });
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet).not.toHaveProperty('sparklineGroups');
  });

  it('validates worksheet-extension capture root, nesting, text, and completion', () => {
    const wrongName = extensionCapture();
    expect(
      captureSync(() => wrongName.openElement(element('wrong'))).diagnostic
        .message,
    ).toBe('Worksheet extension capture root is invalid');

    const wrongNamespace = extensionCapture();
    expect(
      captureSync(() =>
        wrongNamespace.openElement(element('extLst', 'urn:wrong')),
      ).diagnostic.message,
    ).toBe('Worksheet extension capture root is invalid');

    const duplicateRoot = extensionCapture();
    duplicateRoot.openElement(element('extLst'));
    duplicateRoot.closeElement(element('extLst'));
    expect(
      captureSync(() => duplicateRoot.openElement(element('extLst'))).diagnostic
        .message,
    ).toBe('Worksheet extension capture root is invalid');

    const missingClose = extensionCapture();
    missingClose.openElement(element('extLst'));
    expect(captureSync(() => missingClose.result()).diagnostic.message).toBe(
      'Worksheet extension capture is incomplete',
    );

    const wrongClose = extensionCapture();
    wrongClose.openElement(element('extLst'));
    expect(
      captureSync(() => wrongClose.closeElement(element('wrong'))).diagnostic
        .message,
    ).toBe('Worksheet extension capture nesting is invalid');

    const wrongCloseNamespace = extensionCapture();
    wrongCloseNamespace.openElement(element('extLst'));
    expect(
      captureSync(() =>
        wrongCloseNamespace.closeElement(element('extLst', 'urn:wrong')),
      ).diagnostic.message,
    ).toBe('Worksheet extension capture nesting is invalid');

    const empty = extensionCapture();
    expect(
      captureSync(() => empty.closeElement(element('extLst'))).diagnostic
        .message,
    ).toBe('Worksheet extension capture nesting is invalid');
    expect(captureSync(() => empty.text('outside')).diagnostic.message).toBe(
      'Worksheet extension text is outside the root',
    );
    expect(captureSync(() => empty.result()).diagnostic.message).toBe(
      'Worksheet extension capture is incomplete',
    );
  });

  it('requires exactly one sparkline collection per group', async () => {
    const missing = `<ext uri="${SPARKLINE_URI}"><x14:sparklineGroups><x14:sparklineGroup/></x14:sparklineGroups></ext>`;
    expect(
      (
        await capture({
          'xl/worksheets/sheet1.xml': worksheet(missing),
        })
      ).diagnostic.message,
    ).toBe('Sparkline group must contain one sparkline collection');
  });

  it('rejects duplicate worksheet extension lists', async () => {
    const xml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><extLst/><extLst/></worksheet>`;
    expect(
      (await capture({ 'xl/worksheets/sheet1.xml': xml })).diagnostic.message,
    ).toBe('Worksheet contains duplicate extension lists');
  });

  it.each([
    [
      worksheet(`<ext uri="${SPARKLINE_URI}"><x14:wrong/></ext>`),
      'Sparkline extension payload is invalid',
    ],
    [
      worksheet(`${sparklineExtension()}${sparklineExtension()}`),
      'Worksheet contains duplicate sparkline extensions',
    ],
    [worksheet(sparklineExtension('', '')), 'Sparkline group is empty'],
    [
      worksheet(sparklineExtension('', sparkline('', 'B1'))),
      'Sparkline data formula is invalid',
    ],
    [
      worksheet(sparklineExtension('', sparkline('A1:A3', 'B1:B2'))),
      'Sparkline location is invalid',
    ],
    [worksheet(sparklineExtension('type="bad"')), 'Sparkline type is invalid'],
    [
      worksheet(sparklineExtension('markers="bad"')),
      'Sparkline marker flag is invalid',
    ],
    [
      worksheet(sparklineExtension('lineWeight="-1"')),
      'Sparkline line weight is invalid',
    ],
    [
      worksheet(sparklineExtension('minAxisType="custom"')),
      'Sparkline custom minimum requires a manual value',
    ],
    [
      worksheet(sparklineExtension('maxAxisType="custom"')),
      'Sparkline custom maximum requires a manual value',
    ],
    [
      worksheet(
        sparklineExtension(
          '',
          sparkline('A1:A3', 'B1'),
          '<x14:colorSeries rgb="FF112233"/><x14:colorSeries rgb="FF445566"/>',
        ),
      ),
      'Sparkline colorSeries element is duplicated',
    ],
    [
      worksheet(
        sparklineExtension(
          '',
          '<x14:sparkline><xm:sqref>B1</xm:sqref></x14:sparkline>',
        ),
      ),
      'Sparkline formula and location are required exactly once',
    ],
    [
      worksheet(
        sparklineExtension(
          '',
          '<x14:sparkline><xm:f>A1:A3</xm:f></x14:sparkline>',
        ),
      ),
      'Sparkline formula and location are required exactly once',
    ],
    [
      worksheet(
        sparklineExtension(
          '',
          '<x14:sparkline><xm:f>A1:A3</xm:f><xm:f>C1:C3</xm:f><xm:sqref>B1</xm:sqref></x14:sparkline>',
        ),
      ),
      'Sparkline formula and location are required exactly once',
    ],
    [
      worksheet(
        sparklineExtension(
          '',
          '<x14:sparkline><xm:f>A1:A3</xm:f><xm:sqref>B1</xm:sqref><xm:sqref>B2</xm:sqref></x14:sparkline>',
        ),
      ),
      'Sparkline formula and location are required exactly once',
    ],
    [
      worksheet(sparklineExtension('', sparkline(' A1:A3', 'B1'))),
      'Sparkline data formula is invalid',
    ],
    [
      worksheet(sparklineExtension('', sparkline('A1:A3', 'B1:C1'))),
      'Sparkline location is invalid',
    ],
    [
      worksheet(
        `<ext uri="${SPARKLINE_URI}"><x14:sparklineGroups><x14:sparklineGroup/></x14:sparklineGroups></ext>`,
      ),
      'Sparkline group must contain one sparkline collection',
    ],
    [
      worksheet(
        `<ext uri="${SPARKLINE_URI}"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines>${sparkline('A', 'B1')}</x14:sparklines><x14:sparklines>${sparkline('C', 'B2')}</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`,
      ),
      'Sparkline group must contain one sparkline collection',
    ],
    [
      worksheet(sparklineExtension('manualMax="bad"')),
      'Sparkline manual maximum is invalid',
    ],
    [
      worksheet(sparklineExtension('manualMin="bad"')),
      'Sparkline manual minimum is invalid',
    ],
    [
      worksheet(sparklineExtension('lineWeight=" 1"')),
      'Sparkline line weight is invalid',
    ],
    [
      worksheet(sparklineExtension('lineWeight="1 "')),
      'Sparkline line weight is invalid',
    ],
    [
      worksheet(sparklineExtension('lineWeight="1e309"')),
      'Sparkline line weight is invalid',
    ],
    [
      worksheet(sparklineExtension('minAxisType="bad"')),
      'Sparkline minimum-axis type is invalid',
    ],
    [
      worksheet(sparklineExtension('maxAxisType="bad"')),
      'Sparkline maximum-axis type is invalid',
    ],
    [
      worksheet(sparklineExtension('displayEmptyCellsAs="bad"')),
      'Sparkline empty-cell display mode is invalid',
    ],
    [
      worksheet(sparklineExtension('dateAxis="bad"')),
      'Sparkline date-axis flag is invalid',
    ],
    [
      worksheet(sparklineExtension('displayHidden="bad"')),
      'Sparkline hidden-data flag is invalid',
    ],
    [
      worksheet(sparklineExtension('displayXAxis="bad"')),
      'Sparkline X-axis flag is invalid',
    ],
    [
      worksheet(sparklineExtension('first="bad"')),
      'Sparkline first-point flag is invalid',
    ],
    [
      worksheet(sparklineExtension('high="bad"')),
      'Sparkline high-point flag is invalid',
    ],
    [
      worksheet(sparklineExtension('last="bad"')),
      'Sparkline last-point flag is invalid',
    ],
    [
      worksheet(sparklineExtension('low="bad"')),
      'Sparkline low-point flag is invalid',
    ],
    [
      worksheet(sparklineExtension('negative="bad"')),
      'Sparkline negative-point flag is invalid',
    ],
    [
      worksheet(sparklineExtension('rightToLeft="bad"')),
      'Sparkline right-to-left flag is invalid',
    ],
    [
      worksheet(
        sparklineExtension(
          '',
          sparkline('A1:A3', 'B1'),
          '<x14:colorSeries rgb="bad"/>',
        ),
      ),
      'Sparkline colorSeries color RGB is invalid',
    ],
  ] as const)('rejects invalid sparkline contract %#', async (xml, message) => {
    expect(
      (await capture({ 'xl/worksheets/sheet1.xml': xml })).diagnostic,
    ).toMatchObject({ message });
  });
});
