import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
  XlsxParseError,
} from '../../src/formats/xlsx';
import {
  parseXlsxConditionalFormatting,
  XlsxConditionalFormattingCapture,
} from '../../src/formats/xlsx/internal/conditional-format';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from '../../src/formats/xlsx/internal/resource-limits';
import type { XlsxResolvedSheetSelection } from '../../src/formats/xlsx/internal/selection';
import type { XlsxXmlElement } from '../../src/formats/xlsx/internal/streaming-xml';
import type { XlsxWorksheetBudget } from '../../src/formats/xlsx/internal/worksheet';
import {
  createIndependentXlsx,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const PART = 'xl/worksheets/sheet1.xml';
const FULL: XlsxResolvedSheetSelection = { kind: 'full-sheet' };

const STYLES_WITH_DXF = `<styleSheet xmlns="${XLSX_SPREADSHEET_NS}">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="1"><dxf><font><b/></font></dxf></dxfs>
</styleSheet>`;

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

function parseTree(
  value: unknown,
  options: {
    differentialStyles?: number;
    limits?: Partial<ResolvedXlsxResourceLimits>;
    priorities?: Set<number>;
    selection?: XlsxResolvedSheetSelection;
  } = {},
) {
  return parseXlsxConditionalFormatting(
    value,
    options.differentialStyles ?? 0,
    options.priorities ?? new Set(),
    options.selection ?? FULL,
    budget(),
    { ...defaultXlsxResourceLimits(), ...options.limits },
    PART,
  );
}

function captureTree(action: () => unknown): XlsxParseError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected normalized conditional formatting to fail');
}

function captureLimit(action: () => unknown): XlsxResourceLimitError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxResourceLimitError);
    return error as XlsxResourceLimitError;
  }
  throw new Error('Expected conditional-format resource failure');
}

function worksheet(formats: string): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/>${formats}</worksheet>`;
}

async function parseFormats(
  formats: string,
  options: Parameters<typeof parseXlsx>[1] = {},
  overrides: Record<string, string | null> = {},
) {
  return parseXlsx(
    await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': worksheet(formats),
      ...overrides,
    }),
    options,
  );
}

async function capture(
  formats: string,
  options: Parameters<typeof parseXlsx>[1] = {},
  overrides: Record<string, string | null> = {},
): Promise<XlsxParseError> {
  try {
    await parseFormats(formats, options, overrides);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected public conditional formatting parsing to fail');
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

const COMPLETE_FORMATS = `<conditionalFormatting sqref="A1:C5" pivot="1">
  <cfRule type="cellIs" dxfId="0" priority="1" stopIfTrue="1" operator="between"><formula>1</formula><formula>10</formula></cfRule>
  <cfRule type="expression" priority="2"><formula>MOD(A1,2)=0</formula></cfRule>
  <cfRule type="aboveAverage" priority="3" aboveAverage="0" equalAverage="1" stdDev="2"/>
  <cfRule type="top10" priority="4" rank="5" bottom="1" percent="1"/>
  <cfRule type="containsText" priority="5" text="Open"/>
  <cfRule type="timePeriod" priority="6" timePeriod="last7Days"/>
  <cfRule type="colorScale" priority="7"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/></colorScale></cfRule>
  <cfRule type="dataBar" priority="8"><dataBar minLength="5" maxLength="95" showValue="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>
  <cfRule type="iconSet" priority="9"><iconSet iconSet="3Arrows" percent="0" reverse="1" showValue="0"><cfvo type="num" val="0"/><cfvo type="num" val="10"/><cfvo type="num" val="20"/></iconSet></cfRule>
</conditionalFormatting>`;

describe('XLSX conditional formatting', () => {
  it('preserves authored order, priorities, formulas, dxf, and visual rules', async () => {
    const document = await parseFormats(
      COMPLETE_FORMATS,
      {},
      {
        'xl/styles.xml': STYLES_WITH_DXF,
      },
    );
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    const formats =
      sheet.kind === 'worksheet' ? sheet.conditionalFormattings : [];
    expect(formats).toHaveLength(1);
    expect(formats[0]).toMatchObject({
      pivot: true,
      ranges: [{ reference: 'A1:C5' }],
      selectionRelation: 'full-sheet',
    });
    expect(formats[0]?.rules).toMatchObject([
      {
        differentialStyle: 0,
        formulas: ['1', '10'],
        operator: 'between',
        priority: 1,
        stopIfTrue: true,
        type: 'cell-is',
      },
      { formulas: ['MOD(A1,2)=0'], priority: 2, type: 'expression' },
      {
        aboveAverage: false,
        equalAverage: true,
        priority: 3,
        standardDeviations: 2,
        type: 'above-average',
      },
      { bottom: true, percent: true, priority: 4, rank: 5, type: 'top' },
      { priority: 5, text: 'Open', type: 'contains-text' },
      { priority: 6, timePeriod: 'last-7-days', type: 'time-period' },
      {
        colorScale: {
          stops: [
            {
              color: { argb: 'FFFF0000', kind: 'rgb' },
              threshold: { kind: 'minimum' },
            },
            {
              color: { argb: 'FFFFFF00', kind: 'rgb' },
              threshold: { kind: 'percentile', value: 50 },
            },
            {
              color: { argb: 'FF00FF00', kind: 'rgb' },
              threshold: { kind: 'maximum' },
            },
          ],
        },
        priority: 7,
        type: 'color-scale',
      },
      {
        dataBar: {
          color: { argb: 'FF638EC6', kind: 'rgb' },
          maximumLength: 95,
          minimumLength: 5,
          showValue: false,
        },
        priority: 8,
        type: 'data-bar',
      },
      {
        iconSet: {
          iconSet: '3Arrows',
          percent: false,
          reverse: true,
          showValue: false,
          thresholds: [
            { greaterThanOrEqual: true, kind: 'number', value: 0 },
            { greaterThanOrEqual: true, kind: 'number', value: 10 },
            { greaterThanOrEqual: true, kind: 'number', value: 20 },
          ],
        },
        priority: 9,
        type: 'icon-set',
      },
    ]);
    expect(JSON.parse(JSON.stringify(document))).toEqual(document);
  });

  it('filters emitted containers by selection while validating every priority', async () => {
    const included = await parseFormats(
      COMPLETE_FORMATS,
      {
        selection: { ranges: { Sheet1: ['C5'] } },
      },
      { 'xl/styles.xml': STYLES_WITH_DXF },
    );
    const includedSheet = included.sheets[0]!;
    expect(
      includedSheet.kind === 'worksheet'
        ? includedSheet.conditionalFormattings
        : [],
    ).toMatchObject([{ selectionRelation: 'intersects-selection' }]);
    const excluded = await parseFormats(
      COMPLETE_FORMATS,
      {
        selection: { ranges: { Sheet1: ['Z1'] } },
      },
      { 'xl/styles.xml': STYLES_WITH_DXF },
    );
    const excludedSheet = excluded.sheets[0]!;
    expect(
      excludedSheet.kind === 'worksheet'
        ? excludedSheet.conditionalFormattings
        : [],
    ).toEqual([]);
  });

  it.each([
    ['0', false],
    ['false', false],
    ['1', true],
    ['true', true],
  ] as const)(
    'parses every conditional boolean lexical form %s',
    async (source, expected) => {
      const xml = `<conditionalFormatting sqref="A1" pivot="${source}">
      <cfRule type="aboveAverage" priority="1" aboveAverage="${source}" equalAverage="${source}" stopIfTrue="${source}"/>
      <cfRule type="top10" priority="2" rank="1" bottom="${source}" percent="${source}"/>
      <cfRule type="dataBar" priority="3"><dataBar showValue="${source}"><cfvo type="min" gte="${source}"/><cfvo type="max"/><color rgb="FF000000"/></dataBar></cfRule>
      <cfRule type="iconSet" priority="4"><iconSet percent="${source}" reverse="${source}" showValue="${source}"><cfvo type="min"/><cfvo type="percent" val="33"/><cfvo type="max"/></iconSet></cfRule>
    </conditionalFormatting>`;
      const document = await parseFormats(xml);
      const sheet = document.sheets[0]!;
      const format =
        sheet.kind === 'worksheet'
          ? sheet.conditionalFormattings[0]
          : undefined;
      expect(format?.pivot).toBe(expected);
      expect(format?.rules[0]).toMatchObject({
        aboveAverage: expected,
        equalAverage: expected,
        stopIfTrue: expected,
      });
      expect(format?.rules[1]).toMatchObject({
        bottom: expected,
        percent: expected,
      });
      expect(format?.rules[2]?.dataBar).toMatchObject({ showValue: expected });
      expect(format?.rules[2]?.dataBar?.thresholds[0]).toMatchObject({
        greaterThanOrEqual: expected,
      });
      expect(format?.rules[3]?.iconSet).toMatchObject({
        percent: expected,
        reverse: expected,
        showValue: expected,
      });
    },
  );

  it('round-trips conditional formats through portable exact R0', async () => {
    const bytes = await createIndependentXlsx({
      'xl/styles.xml': STYLES_WITH_DXF,
      'xl/worksheets/sheet1.xml': worksheet(COMPLETE_FORMATS),
    });
    const snapshot = await readXlsxRoundTrip(bytes);
    const validated = await validateXlsxRoundTripJson(
      JSON.parse(JSON.stringify(snapshot)) as unknown,
    );
    const result = await writeXlsxRoundTrip(validated);
    expect(result.data).toEqual(bytes);
    expect(result.report.level).toBe('R0');
    expect(result.report.outputSha256).toBe(result.report.sourceSha256);
  });

  it('parses prefixed Strict conditional formatting', async () => {
    const strictSheet = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictRelationship =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const bytes = await createIndependentXlsx({
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelationship}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${strictRelationship}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheet}" xmlns:r="${strictRelationship}"><s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets></s:workbook>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheet}"><s:sheetData/><s:conditionalFormatting sqref="A1"><s:cfRule type="expression" priority="1"><s:formula>A1&gt;0</s:formula></s:cfRule></s:conditionalFormatting></s:worksheet>`,
    });
    const document = await parseXlsx(bytes);
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.conditionalFormattings : [],
    ).toMatchObject([{ rules: [{ formulas: ['A1>0'], type: 'expression' }] }]);
  });

  it.each([
    ['aboveAverage', 'above-average', '', ''],
    ['beginsWith', 'begins-with', ' text="x"', ''],
    ['cellIs', 'cell-is', ' operator="equal"', '<formula>1</formula>'],
    [
      'colorScale',
      'color-scale',
      '',
      '<colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale>',
    ],
    ['containsBlanks', 'contains-blanks', '', ''],
    ['containsErrors', 'contains-errors', '', ''],
    ['containsText', 'contains-text', ' text="x"', ''],
    [
      'dataBar',
      'data-bar',
      '',
      '<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/></dataBar>',
    ],
    ['duplicateValues', 'duplicate-values', '', ''],
    ['endsWith', 'ends-with', ' text="x"', ''],
    ['expression', 'expression', '', '<formula>TRUE</formula>'],
    [
      'iconSet',
      'icon-set',
      '',
      '<iconSet><cfvo type="min"/><cfvo type="percent" val="33"/><cfvo type="max"/></iconSet>',
    ],
    ['notContainsBlanks', 'not-contains-blanks', '', ''],
    ['notContainsErrors', 'not-contains-errors', '', ''],
    ['notContainsText', 'not-contains-text', ' text="x"', ''],
    ['timePeriod', 'time-period', ' timePeriod="today"', ''],
    ['top10', 'top', ' rank="1"', ''],
    ['uniqueValues', 'unique-values', '', ''],
  ] as const)(
    'normalizes conditional rule type %s',
    async (source, expected, attrs, children) => {
      const rule =
        children.length === 0
          ? `<cfRule type="${source}" priority="1"${attrs}/>`
          : `<cfRule type="${source}" priority="1"${attrs}>${children}</cfRule>`;
      const document = await parseFormats(
        `<conditionalFormatting sqref="A1">${rule}</conditionalFormatting>`,
      );
      const sheet = document.sheets[0]!;
      expect(
        sheet.kind === 'worksheet'
          ? sheet.conditionalFormattings[0]?.rules[0]?.type
          : undefined,
      ).toBe(expected);
    },
  );

  it.each([
    ['between', 2],
    ['equal', 1],
    ['greaterThan', 1],
    ['greaterThanOrEqual', 1],
    ['lessThan', 1],
    ['lessThanOrEqual', 1],
    ['notBetween', 2],
    ['notEqual', 1],
  ] as const)('normalizes cell operator %s', async (operator, count) => {
    const formulas = Array.from(
      { length: count },
      (_, index) => `<formula>${index + 1}</formula>`,
    ).join('');
    const document = await parseFormats(
      `<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="${operator}">${formulas}</cfRule></conditionalFormatting>`,
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.conditionalFormattings[0]?.rules[0]?.formulas
        : [],
    ).toHaveLength(count);
  });

  it.each([
    'last7Days',
    'lastMonth',
    'lastWeek',
    'nextMonth',
    'nextWeek',
    'thisMonth',
    'thisWeek',
    'today',
    'tomorrow',
    'yesterday',
  ] as const)('accepts time period %s', async (timePeriod) => {
    await expect(
      parseFormats(
        `<conditionalFormatting sqref="A1"><cfRule type="timePeriod" priority="1" timePeriod="${timePeriod}"/></conditionalFormatting>`,
      ),
    ).resolves.toBeDefined();
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
  ] as const)('accepts conditional icon set %s', async (iconSet) => {
    const count = Number(iconSet[0]);
    const thresholds = Array.from(
      { length: count },
      (_, index) => `<cfvo type="num" val="${index}"/>`,
    ).join('');
    await expect(
      parseFormats(
        `<conditionalFormatting sqref="A1"><cfRule type="iconSet" priority="1"><iconSet iconSet="${iconSet}">${thresholds}</iconSet></cfRule></conditionalFormatting>`,
      ),
    ).resolves.toBeDefined();
  });

  it('enforces rule, range, formula, text, dxf, and grid limits', async () => {
    const twoRules = `<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" priority="1"/><cfRule type="duplicateValues" priority="2"/></conditionalFormatting>`;
    await expect(
      parseFormats(twoRules, { limits: { maxConditionalFormattingRules: 2 } }),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(twoRules, {
          limits: { maxConditionalFormattingRules: 1 },
        })
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxConditionalFormattingRules',
    });
    const ranges = `<conditionalFormatting sqref="A1 B1"><cfRule type="uniqueValues" priority="1"/></conditionalFormatting>`;
    expect(
      (await capture(ranges, { limits: { maxRangeAreas: 1 } })).diagnostic,
    ).toMatchObject({ actual: 2, limit: 1, limitName: 'maxRangeAreas' });
    const formula = `<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>ABC</formula></cfRule></conditionalFormatting>`;
    expect(
      (
        await capture(formula, {
          limits: { maxFormulaCharacters: 2, maxTotalFormulaCharacters: 3 },
        })
      ).diagnostic,
    ).toMatchObject({ actual: 3, limit: 2, limitName: 'maxFormulaCharacters' });
    const text = `<conditionalFormatting sqref="A1"><cfRule type="containsText" priority="1" text="abc"/></conditionalFormatting>`;
    expect(
      (await capture(text, { limits: { maxTextCharacters: 11 } })).diagnostic,
    ).toMatchObject({ actual: 12, limit: 11, limitName: 'maxTextCharacters' });
    const dxf = `<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" priority="1" dxfId="1"/></conditionalFormatting>`;
    expect(
      (await capture(dxf, {}, { 'xl/styles.xml': STYLES_WITH_DXF })).diagnostic
        .message,
    ).toBe('Conditional-format differential-style reference is invalid');
    for (const [sqref, limits, limitName] of [
      ['A2', { maxRowsPerWorksheet: 1 }, 'maxRowsPerWorksheet'],
      ['B1', { maxColumnsPerWorksheet: 1 }, 'maxColumnsPerWorksheet'],
    ] as const) {
      expect(
        (
          await capture(
            `<conditionalFormatting sqref="${sqref}"><cfRule type="uniqueValues" priority="1"/></conditionalFormatting>`,
            { limits },
          )
        ).diagnostic,
      ).toMatchObject({ actual: 2, limit: 1, limitName });
    }
  });

  it.each([
    [
      '<conditionalFormatting><cfRule type="uniqueValues" priority="1"/></conditionalFormatting>',
      'Conditional-format range list is invalid',
    ],
    [
      '<conditionalFormatting sqref=""><cfRule type="uniqueValues" priority="1"/></conditionalFormatting>',
      'Conditional-format range list is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"/>',
      'Conditional-format rule collection is invalid',
    ],
    [
      '<conditionalFormatting sqref="bad"><cfRule type="uniqueValues" priority="1"/></conditionalFormatting>',
      'Conditional-format range is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1 A1"><cfRule type="uniqueValues" priority="1"/></conditionalFormatting>',
      'Conditional-format range list contains duplicates',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="bad" priority="1"/></conditionalFormatting>',
      'Conditional-format rule type is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" priority="0"/></conditionalFormatting>',
      'Conditional-format priority is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1"><formula>1</formula></cfRule></conditionalFormatting>',
      'Cell conditional operator is missing',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="between"><formula>1</formula></cfRule></conditionalFormatting>',
      'Cell conditional formula count is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"/></conditionalFormatting>',
      'Expression conditional formula is missing',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="containsText" priority="1"/></conditionalFormatting>',
      'Text conditional comparison text is missing',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="containsText" priority="1" text=""/></conditionalFormatting>',
      'Text conditional comparison text is missing',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="beginsWith" priority="1"/></conditionalFormatting>',
      'Text conditional comparison text is missing',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="endsWith" priority="1"/></conditionalFormatting>',
      'Text conditional comparison text is missing',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="notContainsText" priority="1"/></conditionalFormatting>',
      'Text conditional comparison text is missing',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="timePeriod" priority="1" timePeriod="bad"/></conditionalFormatting>',
      'Time-period conditional value is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="top10" priority="1"/></conditionalFormatting>',
      'Top conditional rank is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="top10" priority="1" rank="0"/></conditionalFormatting>',
      'Top conditional rank is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="top10" priority="1" rank="01"/></conditionalFormatting>',
      'Top conditional rank is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><color rgb="FF000000"/></colorScale></cfRule></conditionalFormatting>',
      'Conditional color scale stop count is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/></colorScale></cfRule></conditionalFormatting>',
      'Conditional color scale stop count is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><color rgb="FF000000"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional data-bar threshold count is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="min"/></iconSet></cfRule></conditionalFormatting>',
      'Conditional icon-set threshold count is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>1</formula><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional-format visual definition mismatches its rule',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>1</formula><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale></cfRule></conditionalFormatting>',
      'Conditional-format visual definition mismatches its rule',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>1</formula><iconSet><cfvo type="min"/><cfvo type="percent" val="33"/><cfvo type="max"/></iconSet></cfRule></conditionalFormatting>',
      'Conditional-format visual definition mismatches its rule',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" priority="1" stopIfTrue="bad"/></conditionalFormatting>',
      'Conditional-format stop flag is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" priority="01"/></conditionalFormatting>',
      'Conditional-format priority is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" priority="4294967296"/></conditionalFormatting>',
      'Conditional-format priority is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" priority="1" dxfId="01"/></conditionalFormatting>',
      'Conditional-format differential-style reference is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="aboveAverage" priority="1" stdDev="01"/></conditionalFormatting>',
      'Average conditional standard deviations are invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" priority="1" operator="bad"/></conditionalFormatting>',
      'Conditional-format operator is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" priority="1"><formula>1</formula><formula>2</formula><formula>3</formula><formula>4</formula></cfRule></conditionalFormatting>',
      'Conditional-format formula collection is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula/></cfRule></conditionalFormatting>',
      'Conditional-format formula is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional-format rule has multiple visual definitions',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="bad"/><cfvo type="max"/><color rgb="FF000000"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional-format threshold type is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="percent" val="101"/><cfvo type="max"/><color rgb="FF000000"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional-format threshold value is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="num" val="bad"/><cfvo type="max"/><color rgb="FF000000"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional-format threshold value is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="formula" val=""/><cfvo type="max"/><color rgb="FF000000"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional-format threshold formula is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min" gte="bad"/><cfvo type="max"/><color rgb="FF000000"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional-format threshold inclusive flag is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color/></dataBar></cfRule></conditionalFormatting>',
      'Conditional data bar color is missing',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional data-bar color count is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar minLength="90" maxLength="10"><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional data-bar lengths are inconsistent',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar minLength="101"><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional data-bar minimum length is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar maxLength="bad"><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional data-bar maximum length is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><dataBar showValue="bad"><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/></dataBar></cfRule></conditionalFormatting>',
      'Conditional data-bar show-value flag is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="iconSet" priority="1"><iconSet iconSet="bad"><cfvo type="min"/><cfvo type="percent" val="33"/><cfvo type="max"/></iconSet></cfRule></conditionalFormatting>',
      'Conditional icon set kind is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="iconSet" priority="1"><iconSet percent="bad"><cfvo type="min"/><cfvo type="percent" val="33"/><cfvo type="max"/></iconSet></cfRule></conditionalFormatting>',
      'Conditional icon-set percent flag is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="iconSet" priority="1"><iconSet reverse="bad"><cfvo type="min"/><cfvo type="percent" val="33"/><cfvo type="max"/></iconSet></cfRule></conditionalFormatting>',
      'Conditional icon-set reverse flag is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="iconSet" priority="1"><iconSet showValue="bad"><cfvo type="min"/><cfvo type="percent" val="33"/><cfvo type="max"/></iconSet></cfRule></conditionalFormatting>',
      'Conditional icon-set show-value flag is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="aboveAverage" priority="1" aboveAverage="bad"/></conditionalFormatting>',
      'Average conditional direction flag is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="aboveAverage" priority="1" equalAverage="bad"/></conditionalFormatting>',
      'Average conditional equality flag is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="top10" priority="1" rank="1" bottom="bad"/></conditionalFormatting>',
      'Top conditional bottom flag is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="top10" priority="1" rank="1" percent="bad"/></conditionalFormatting>',
      'Top conditional percent flag is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1" pivot="bad"><cfRule type="uniqueValues" priority="1"/></conditionalFormatting>',
      'Conditional-format pivot flag is invalid',
    ],
    [
      '<conditionalFormatting sqref="A1"><cfRule type="uniqueValues" priority="1"/></conditionalFormatting><conditionalFormatting sqref="B1"><cfRule type="duplicateValues" priority="1"/></conditionalFormatting>',
      'Worksheet contains duplicate conditional-format priorities',
    ],
    [
      '<conditionalFormatting sqref="A1">bad<cfRule type="uniqueValues" priority="1"/></conditionalFormatting>',
      'Conditional-format text content is invalid',
    ],
  ] as const)(
    'rejects invalid conditional-format contract %#',
    async (xml, message) => {
      expect((await capture(xml)).diagnostic.message).toBe(message);
    },
  );
});

describe('XLSX normalized conditional-format parser and capture', () => {
  it.each(['+1', '-1.5', '.5', '.55', '1.', '1e3', '1e30', '-2.5E-2'])(
    'accepts finite conditional numeric lexical form %s',
    (val) => {
      expect(
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'dataBar' },
            dataBar: {
              cfvo: [
                { attrs: { type: 'num', val } },
                { attrs: { type: 'max' } },
              ],
              color: { attrs: { rgb: 'FF000000' } },
            },
          },
        })?.rules[0]?.dataBar?.thresholds[0],
      ).toMatchObject({ kind: 'number', value: Number(val) });
    },
  );

  it.each([
    '',
    ' ',
    '1x',
    '0x1',
    '1e',
    '1e+',
    '1_0',
    'NaN',
    'Infinity',
    '1e309',
  ])('rejects invalid conditional numeric lexical form %#', (val) => {
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'dataBar' },
            dataBar: {
              cfvo: [
                { attrs: { type: 'num', val } },
                { attrs: { type: 'max' } },
              ],
              color: { attrs: { rgb: 'FF000000' } },
            },
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional-format threshold value is invalid');
  });

  it.each([
    ['percent', '0'],
    ['percent', '100'],
    ['percentile', '0'],
    ['percentile', '100'],
  ] as const)('accepts exact threshold boundary %s %s', (type, val) => {
    expect(
      parseTree({
        attrs: { sqref: 'A1' },
        cfRule: {
          attrs: { priority: '1', type: 'dataBar' },
          dataBar: {
            cfvo: [{ attrs: { type, val } }, { attrs: { type: 'max' } }],
            color: { attrs: { rgb: 'FF000000' } },
          },
        },
      })?.rules[0]?.dataBar?.thresholds[0],
    ).toMatchObject({ value: Number(val) });
  });

  it.each([
    ['percent', '-1'],
    ['percent', '101'],
    ['percentile', '-1'],
    ['percentile', '101'],
  ] as const)('rejects threshold one over %s %s', (type, val) => {
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'dataBar' },
            dataBar: {
              cfvo: [{ attrs: { type, val } }, { attrs: { type: 'max' } }],
              color: { attrs: { rgb: 'FF000000' } },
            },
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional-format threshold value is invalid');
  });

  it('preserves rule-family defaults without leaking unrelated fields', () => {
    expect(
      parseTree({
        attrs: { sqref: 'A1' },
        cfRule: { attrs: { priority: '1', type: 'uniqueValues' } },
      })?.rules[0],
    ).toEqual({
      formulas: [],
      priority: 1,
      stopIfTrue: false,
      type: 'unique-values',
    });
    expect(
      parseTree({
        attrs: { sqref: 'A1' },
        cfRule: { attrs: { priority: '1', type: 'aboveAverage' } },
      })?.rules[0],
    ).toEqual({
      aboveAverage: true,
      equalAverage: false,
      formulas: [],
      priority: 1,
      stopIfTrue: false,
      type: 'above-average',
    });
    expect(
      parseTree({
        attrs: { sqref: 'A1' },
        cfRule: { attrs: { priority: '1', rank: '1', type: 'top10' } },
      })?.rules[0],
    ).toEqual({
      bottom: false,
      formulas: [],
      percent: false,
      priority: 1,
      rank: 1,
      stopIfTrue: false,
      type: 'top',
    });
  });

  it('decodes XML entities and accepts maximum priority and grid bounds', () => {
    expect(
      parseTree(
        {
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: {
              priority: '4294967295',
              text: '&amp;&apos;&gt;&lt;&quot;&#65;&#x1F600;',
              type: 'containsText',
            },
          },
        },
        { limits: { maxColumnsPerWorksheet: 1, maxRowsPerWorksheet: 1 } },
      )?.rules[0],
    ).toMatchObject({
      priority: 0xffff_ffff,
      text: '&\'><"A😀',
    });
  });

  it('exposes every visual and rule default explicitly', () => {
    const dataBar = parseTree({
      attrs: { sqref: 'A1' },
      cfRule: {
        attrs: { priority: '1', type: 'dataBar' },
        dataBar: {
          cfvo: [{ attrs: { type: 'min' } }, { attrs: { type: 'max' } }],
          color: { attrs: { rgb: 'FF000000' } },
        },
      },
    });
    expect(dataBar).toMatchObject({
      pivot: false,
      rules: [
        {
          dataBar: {
            maximumLength: 90,
            minimumLength: 10,
            showValue: true,
          },
          stopIfTrue: false,
        },
      ],
    });
    const icons = parseTree({
      attrs: { sqref: 'A1' },
      cfRule: {
        attrs: { priority: '1', type: 'iconSet' },
        iconSet: {
          cfvo: [
            { attrs: { type: 'min' } },
            { attrs: { type: 'percent', val: '33' } },
            { attrs: { type: 'max' } },
          ],
        },
      },
    });
    expect(icons?.rules[0]?.iconSet).toMatchObject({
      iconSet: '3TrafficLights1',
      percent: true,
      reverse: false,
      showValue: true,
    });
  });

  it('accepts exact data-bar length and formula cardinality boundaries', () => {
    expect(
      parseTree({
        attrs: { sqref: 'A1' },
        cfRule: {
          attrs: { priority: '1', type: 'dataBar' },
          dataBar: {
            attrs: { maxLength: '50', minLength: '50' },
            cfvo: [{ attrs: { type: 'min' } }, { attrs: { type: 'max' } }],
            color: { attrs: { rgb: 'FF000000' } },
          },
        },
      })?.rules[0]?.dataBar,
    ).toMatchObject({ maximumLength: 50, minimumLength: 50 });
    expect(
      parseTree({
        attrs: { sqref: 'A1' },
        cfRule: {
          attrs: { priority: '1', type: 'uniqueValues' },
          formula: [{ value: '1' }, { value: '2' }, { value: '3' }],
        },
      })?.rules[0]?.formulas,
    ).toEqual(['1', '2', '3']);
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'expression' },
            formula: { value: '' },
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional-format formula is invalid');
  });

  it('normalizes repeated XML whitespace in range lists', () => {
    expect(
      parseTree({
        attrs: { sqref: '\n A1  \t B2 \r' },
        cfRule: { attrs: { priority: '1', type: 'uniqueValues' } },
      })?.ranges,
    ).toMatchObject([{ reference: 'A1' }, { reference: 'B2' }]);
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: '' },
          cfRule: { attrs: { priority: '1', type: 'uniqueValues' } },
        }),
      ).diagnostic.message,
    ).toBe('Conditional-format range list is invalid');
  });

  it('reports exact malformed visual collections', () => {
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'colorScale' },
            colorScale: {
              cfvo: 'bad',
              color: [
                { attrs: { rgb: 'FF000000' } },
                { attrs: { rgb: 'FFFFFFFF' } },
              ],
            },
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional color scale threshold collection is invalid');
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'colorScale' },
            colorScale: {
              cfvo: [{ attrs: { type: 'min' } }, { attrs: { type: 'max' } }],
              color: 'bad',
            },
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional color scale color collection is invalid');
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'dataBar' },
            dataBar: {
              cfvo: [
                { attrs: { type: 'num', val: 1 } },
                { attrs: { type: 'max' } },
              ],
              color: { attrs: { rgb: 'FF000000' } },
            },
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional-format threshold value is invalid');
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'dataBar' },
            dataBar: {
              cfvo: 'bad',
              color: { attrs: { rgb: 'FF000000' } },
            },
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional data bar threshold collection is invalid');
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'dataBar' },
            dataBar: {
              cfvo: [{ attrs: { type: 'min' } }, { attrs: { type: 'max' } }],
              color: 'bad',
            },
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional data bar color collection is invalid');
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'colorScale' },
            colorScale: [
              {
                cfvo: [{ attrs: { type: 'min' } }, { attrs: { type: 'max' } }],
                color: [
                  { attrs: { rgb: 'FF000000' } },
                  { attrs: { rgb: 'FFFFFFFF' } },
                ],
              },
            ],
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional color scale is invalid');
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'iconSet' },
            iconSet: [
              {
                cfvo: [
                  { attrs: { type: 'min' } },
                  { attrs: { type: 'percent', val: '33' } },
                  { attrs: { type: 'max' } },
                ],
              },
            ],
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional icon set is invalid');
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'dataBar' },
            dataBar: [
              {
                cfvo: [{ attrs: { type: 'min' } }, { attrs: { type: 'max' } }],
                color: { attrs: { rgb: 'FF000000' } },
              },
            ],
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional data bar is invalid');
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'iconSet' },
            iconSet: {
              attrs: { iconSet: 1 },
              cfvo: [
                { attrs: { type: 'min' } },
                { attrs: { type: 'percent', val: '33' } },
                { attrs: { type: 'max' } },
              ],
            },
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional icon set kind is invalid');
    expect(
      captureTree(() =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: '1', type: 'iconSet' },
            iconSet: { cfvo: 'bad' },
          },
        }),
      ).diagnostic.message,
    ).toBe('Conditional icon set threshold collection is invalid');
  });

  it('parses every threshold kind and default flags', () => {
    const thresholds = [
      { attrs: { type: 'min' } },
      { attrs: { type: 'max' } },
      { attrs: { type: 'num', val: '-0' } },
      { attrs: { type: 'percent', val: '25' } },
      { attrs: { type: 'percentile', val: '75' } },
      { attrs: { type: 'formula', val: 'A1' } },
    ];
    const outputs = thresholds.map(
      (cfvo, index) =>
        parseTree({
          attrs: { sqref: 'A1' },
          cfRule: {
            attrs: { priority: String(index + 1), type: 'dataBar' },
            dataBar: {
              cfvo: [cfvo, { attrs: { type: 'max' } }],
              color: { attrs: { rgb: 'FF000000' } },
            },
          },
        })?.rules[0]?.dataBar?.thresholds[0],
    );
    expect(outputs).toMatchObject([
      { greaterThanOrEqual: true, kind: 'minimum' },
      { greaterThanOrEqual: true, kind: 'maximum' },
      { kind: 'number', value: 0 },
      { kind: 'percent', value: 25 },
      { kind: 'percentile', value: 75 },
      { expression: 'A1', kind: 'formula' },
    ]);
  });

  it('distinguishes selection rectangles and charges scan work', () => {
    const value = {
      attrs: { sqref: 'B2:C3' },
      cfRule: { attrs: { priority: '1', type: 'uniqueValues' } },
    };
    const selected = (reference: string): XlsxResolvedSheetSelection => {
      const column = reference.codePointAt(0)! - 0x40;
      const row = Number(reference.slice(1));
      return {
        endRowPrefix: [row],
        kind: 'selected-ranges',
        ranges: [{ end: { column, row }, reference, start: { column, row } }],
      };
    };
    for (const reference of ['B1', 'B4', 'A2', 'D2']) {
      expect(parseTree(value, { selection: selected(reference) })).toBeNull();
    }
    for (const reference of ['B2', 'C2', 'B3', 'C3']) {
      expect(
        parseTree(value, { selection: selected(reference) }),
      ).toMatchObject({ selectionRelation: 'intersects-selection' });
    }
    expect(
      parseTree(value, { selection: { kind: 'not-selected' } }),
    ).toBeNull();
    expect(
      captureLimit(() =>
        parseTree(
          { ...value, attrs: { sqref: 'A1 B2' } },
          {
            limits: { maxScannedCells: 1 },
            selection: selected('C3'),
          },
        ),
      ),
    ).toMatchObject({ actual: 2, limit: 1, limitName: 'maxScannedCells' });
  });

  it('captures split formula text and rejects incomplete or malformed capture', () => {
    const capture = new XlsxConditionalFormattingCapture(
      0,
      new Set(),
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    capture.openElement(element('conditionalFormatting', { sqref: 'A1' }));
    capture.openElement(
      element('cfRule', { priority: '1', type: 'expression' }),
    );
    capture.openElement(element('formula'));
    capture.text('A1');
    capture.text('>0');
    capture.closeElement(element('formula'));
    capture.closeElement(element('cfRule'));
    capture.closeElement(element('conditionalFormatting'));
    expect(capture.result()).toMatchObject({
      rules: [{ formulas: ['A1>0'], type: 'expression' }],
    });

    const invalid = new XlsxConditionalFormattingCapture(
      0,
      new Set(),
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(
      captureTree(() => invalid.openElement(element('wrong'))).diagnostic
        .message,
    ).toBe('Conditional-format capture root is invalid');
    expect(captureTree(() => parseTree([])).diagnostic.message).toBe(
      'Conditional-format collection is invalid',
    );
    const incomplete = new XlsxConditionalFormattingCapture(
      0,
      new Set(),
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    incomplete.openElement(element('conditionalFormatting'));
    expect(captureTree(() => incomplete.result()).diagnostic.message).toBe(
      'Conditional-format capture is incomplete',
    );

    const duplicate = new XlsxConditionalFormattingCapture(
      0,
      new Set(),
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    duplicate.openElement(element('conditionalFormatting'));
    duplicate.closeElement(element('conditionalFormatting'));
    expect(
      captureTree(() => duplicate.openElement(element('conditionalFormatting')))
        .diagnostic.message,
    ).toBe('Conditional-format capture root is invalid');

    const nesting = new XlsxConditionalFormattingCapture(
      0,
      new Set(),
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    nesting.openElement(element('conditionalFormatting'));
    expect(
      captureTree(() => nesting.closeElement(element('wrong'))).diagnostic
        .message,
    ).toBe('Conditional-format capture nesting is invalid');
  });

  it('preserves raw expanded attribute keys during capture', () => {
    const capture = new XlsxConditionalFormattingCapture(
      0,
      new Set(),
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    const root = element('conditionalFormatting');
    root.attributes = new Map([
      ['{}sqref', 'A1'],
      ['pivot', '1'],
    ]);
    capture.openElement(root);
    capture.openElement(
      element('cfRule', { priority: '1', type: 'uniqueValues' }),
    );
    capture.closeElement(element('cfRule'));
    capture.closeElement(element('conditionalFormatting'));
    expect(capture.result()).toMatchObject({ pivot: true });
  });
});
