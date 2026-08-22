import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
  XlsxParseError,
} from '../../src/formats/xlsx';
import {
  parseXlsxDataValidations,
  XlsxDataValidationsCapture,
} from '../../src/formats/xlsx/internal/data-validation';
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
    limits?: Partial<ResolvedXlsxResourceLimits>;
    selection?: XlsxResolvedSheetSelection;
  } = {},
) {
  return parseXlsxDataValidations(
    value,
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
  throw new Error('Expected normalized data validation to fail');
}

function captureLimit(action: () => unknown): XlsxResourceLimitError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxResourceLimitError);
    return error as XlsxResourceLimitError;
  }
  throw new Error('Expected normalized data validation resource failure');
}

function worksheet(validations: string): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/>${validations}</worksheet>`;
}

async function parseValidations(
  validations: string,
  options: Parameters<typeof parseXlsx>[1] = {},
) {
  return parseXlsx(
    await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': worksheet(validations),
    }),
    options,
  );
}

async function capture(
  validations: string,
  options: Parameters<typeof parseXlsx>[1] = {},
): Promise<XlsxParseError> {
  try {
    await parseValidations(validations, options);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected public XLSX data validation parsing to fail');
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

const COMPLETE_VALIDATIONS = `<dataValidations count="4" disablePrompts="1" xWindow="12" yWindow="34">
  <dataValidation type="whole" operator="between" allowBlank="1" showDropDown="1" showErrorMessage="1" showInputMessage="1" errorStyle="warning" imeMode="hiragana" errorTitle="Invalid" error="Use 1-10" promptTitle="Input" prompt="Enter a number" sqref="A1:A3 C1">
    <formula1>1</formula1><formula2>10</formula2>
  </dataValidation>
  <dataValidation type="list" sqref="B1:B3"><formula1>&quot;One,Two&quot;</formula1></dataValidation>
  <dataValidation type="custom" sqref="D1"><formula1>AND(D1&gt;0,D1&lt;5)</formula1></dataValidation>
  <dataValidation sqref="E1"/>
</dataValidations>`;

describe('XLSX data validations', () => {
  it('preserves typed rules, multi-ranges, formulas, prompts, and settings', async () => {
    const document = await parseValidations(COMPLETE_VALIDATIONS);
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    expect(
      sheet.kind === 'worksheet' ? sheet.dataValidationSettings : null,
    ).toEqual({
      disablePrompts: true,
      xWindow: 12,
      yWindow: 34,
    });
    expect(sheet.kind === 'worksheet' ? sheet.dataValidations : []).toEqual([
      {
        allowBlank: true,
        error: 'Use 1-10',
        errorStyle: 'warning',
        errorTitle: 'Invalid',
        formula1: '1',
        formula2: '10',
        imeMode: 'hiragana',
        operator: 'between',
        prompt: 'Enter a number',
        promptTitle: 'Input',
        ranges: [
          {
            end: { column: 1, row: 3 },
            reference: 'A1:A3',
            start: { column: 1, row: 1 },
          },
          {
            end: { column: 3, row: 1 },
            reference: 'C1',
            start: { column: 3, row: 1 },
          },
        ],
        selectionRelation: 'full-sheet',
        showDropDown: true,
        showErrorMessage: true,
        showInputMessage: true,
        type: 'whole',
      },
      {
        allowBlank: false,
        errorStyle: 'stop',
        formula1: '"One,Two"',
        imeMode: 'no-control',
        operator: 'between',
        ranges: [
          {
            end: { column: 2, row: 3 },
            reference: 'B1:B3',
            start: { column: 2, row: 1 },
          },
        ],
        selectionRelation: 'full-sheet',
        showDropDown: false,
        showErrorMessage: false,
        showInputMessage: false,
        type: 'list',
      },
      expect.objectContaining({
        formula1: 'AND(D1>0,D1<5)',
        type: 'custom',
      }),
      expect.objectContaining({ type: 'none' }),
    ]);
    expect(JSON.parse(JSON.stringify(document))).toEqual(document);
  });

  it('filters emitted rules by selection while validating every rule', async () => {
    const included = await parseValidations(COMPLETE_VALIDATIONS, {
      selection: { ranges: { Sheet1: ['C1'] } },
    });
    const includedSheet = included.sheets[0]!;
    expect(
      includedSheet.kind === 'worksheet' ? includedSheet.dataValidations : [],
    ).toMatchObject([
      { selectionRelation: 'intersects-selection', type: 'whole' },
    ]);
    const excluded = await parseValidations(COMPLETE_VALIDATIONS, {
      selection: { ranges: { Sheet1: ['Z1'] } },
    });
    const excludedSheet = excluded.sheets[0]!;
    expect(
      excludedSheet.kind === 'worksheet' ? excludedSheet.dataValidations : [],
    ).toEqual([]);
    expect(
      excludedSheet.kind === 'worksheet'
        ? excludedSheet.dataValidationSettings
        : undefined,
    ).toEqual({ disablePrompts: true, xWindow: 12, yWindow: 34 });
  });

  it('round-trips validation rules through portable exact R0', async () => {
    const bytes = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': worksheet(COMPLETE_VALIDATIONS),
    });
    const snapshot = await readXlsxRoundTrip(bytes);
    const result = await writeXlsxRoundTrip(
      await validateXlsxRoundTripJson(
        JSON.parse(JSON.stringify(snapshot)) as unknown,
      ),
    );
    expect(result.data).toEqual(bytes);
    expect(result.report.level).toBe('R0');
    expect(result.report.outputSha256).toBe(result.report.sourceSha256);
    expect(
      snapshot.document.sheets[0]?.kind === 'worksheet'
        ? snapshot.document.sheets[0].dataValidations
        : [],
    ).toHaveLength(4);
  });

  it('parses prefixed Strict data validations', async () => {
    const strictSheet = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictRelationship =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const bytes = await createIndependentXlsx({
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelationship}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${strictRelationship}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheet}" xmlns:r="${strictRelationship}"><s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets></s:workbook>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheet}"><s:sheetData/><s:dataValidations count="1"><s:dataValidation type="custom" sqref="A1"><s:formula1>A1&gt;0</s:formula1></s:dataValidation></s:dataValidations></s:worksheet>`,
    });
    const document = await parseXlsx(bytes);
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.dataValidations : [],
    ).toMatchObject([{ formula1: 'A1>0', type: 'custom' }]);
  });

  it.each([
    ['0', false],
    ['false', false],
    ['1', true],
    ['true', true],
  ] as const)(
    'parses every validation boolean lexical form %s',
    async (source, expected) => {
      const document = await parseValidations(
        `<dataValidations disablePrompts="${source}"><dataValidation allowBlank="${source}" showDropDown="${source}" showErrorMessage="${source}" showInputMessage="${source}" sqref="A1"/></dataValidations>`,
      );
      const sheet = document.sheets[0]!;
      expect(
        sheet.kind === 'worksheet' ? sheet.dataValidationSettings : null,
      ).toEqual({ disablePrompts: expected });
      expect(
        sheet.kind === 'worksheet' ? sheet.dataValidations[0] : undefined,
      ).toMatchObject({
        allowBlank: expected,
        showDropDown: expected,
        showErrorMessage: expected,
        showInputMessage: expected,
      });
    },
  );

  it.each([
    ['custom', 'custom'],
    ['date', 'date'],
    ['decimal', 'decimal'],
    ['list', 'list'],
    ['none', 'none'],
    ['textLength', 'text-length'],
    ['time', 'time'],
    ['whole', 'whole'],
  ] as const)('normalizes validation type %s', async (source, expected) => {
    const formulas =
      source === 'none'
        ? ''
        : source === 'list' || source === 'custom'
          ? '<formula1>1</formula1>'
          : '<formula1>1</formula1><formula2>2</formula2>';
    const document = await parseValidations(
      `<dataValidations><dataValidation type="${source}" sqref="A1">${formulas}</dataValidation></dataValidations>`,
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.dataValidations[0]?.type : undefined,
    ).toBe(expected);
  });

  it.each([
    ['between', 'between'],
    ['equal', 'equal'],
    ['greaterThan', 'greater-than'],
    ['greaterThanOrEqual', 'greater-than-or-equal'],
    ['lessThan', 'less-than'],
    ['lessThanOrEqual', 'less-than-or-equal'],
    ['notBetween', 'not-between'],
    ['notEqual', 'not-equal'],
  ] as const)('normalizes validation operator %s', async (source, expected) => {
    const formula2 =
      source === 'between' || source === 'notBetween'
        ? '<formula2>2</formula2>'
        : '';
    const document = await parseValidations(
      `<dataValidations><dataValidation type="whole" operator="${source}" sqref="A1"><formula1>1</formula1>${formula2}</dataValidation></dataValidations>`,
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.dataValidations[0]?.operator
        : undefined,
    ).toBe(expected);
  });

  it.each([
    ['disabled', 'disabled'],
    ['fullAlpha', 'full-alpha'],
    ['fullHangul', 'full-hangul'],
    ['fullKatakana', 'full-katakana'],
    ['halfAlpha', 'half-alpha'],
    ['halfHangul', 'half-hangul'],
    ['halfKatakana', 'half-katakana'],
    ['hiragana', 'hiragana'],
    ['noControl', 'no-control'],
    ['off', 'off'],
    ['on', 'on'],
  ] as const)('normalizes validation IME mode %s', async (source, expected) => {
    const document = await parseValidations(
      `<dataValidations><dataValidation imeMode="${source}" sqref="A1"/></dataValidations>`,
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.dataValidations[0]?.imeMode
        : undefined,
    ).toBe(expected);
  });

  it.each(['information', 'stop', 'warning'] as const)(
    'preserves validation error style %s',
    async (errorStyle) => {
      const document = await parseValidations(
        `<dataValidations><dataValidation errorStyle="${errorStyle}" sqref="A1"/></dataValidations>`,
      );
      const sheet = document.sheets[0]!;
      expect(
        sheet.kind === 'worksheet'
          ? sheet.dataValidations[0]?.errorStyle
          : undefined,
      ).toBe(errorStyle);
    },
  );

  it('enforces validation, range, formula, and configured grid limits', async () => {
    const twoRules = `<dataValidations count="2"><dataValidation sqref="A1"/><dataValidation sqref="B1"/></dataValidations>`;
    await expect(
      parseValidations(twoRules, { limits: { maxValidationRules: 2 } }),
    ).resolves.toBeDefined();
    expect(
      (await capture(twoRules, { limits: { maxValidationRules: 1 } }))
        .diagnostic,
    ).toMatchObject({ actual: 2, limit: 1, limitName: 'maxValidationRules' });
    const ranges = `<dataValidations><dataValidation sqref="A1 B1"/></dataValidations>`;
    await expect(
      parseValidations(ranges, { limits: { maxRangeAreas: 2 } }),
    ).resolves.toBeDefined();
    expect(
      (await capture(ranges, { limits: { maxRangeAreas: 1 } })).diagnostic,
    ).toMatchObject({ actual: 2, limit: 1, limitName: 'maxRangeAreas' });
    const formulaRule = `<dataValidations><dataValidation type="custom" sqref="A1"><formula1>ABC</formula1></dataValidation></dataValidations>`;
    await expect(
      parseValidations(formulaRule, {
        limits: { maxFormulaCharacters: 3, maxTotalFormulaCharacters: 3 },
      }),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(formulaRule, {
          limits: { maxFormulaCharacters: 2, maxTotalFormulaCharacters: 3 },
        })
      ).diagnostic,
    ).toMatchObject({ actual: 3, limit: 2, limitName: 'maxFormulaCharacters' });
    for (const [sqref, limits, limitName] of [
      ['A2', { maxRowsPerWorksheet: 1 }, 'maxRowsPerWorksheet'],
      ['B1', { maxColumnsPerWorksheet: 1 }, 'maxColumnsPerWorksheet'],
    ] as const) {
      expect(
        (
          await capture(
            `<dataValidations><dataValidation sqref="${sqref}"/></dataValidations>`,
            { limits },
          )
        ).diagnostic,
      ).toMatchObject({ actual: 2, limit: 1, limitName });
    }
  });

  it.each([
    [
      '<dataValidations count="2"><dataValidation sqref="A1"/></dataValidations>',
      'Data-validation count does not match',
    ],
    [
      '<dataValidations count="01"><dataValidation sqref="A1"/></dataValidations>',
      'Data-validation count is invalid',
    ],
    [
      '<dataValidations><dataValidation type="bad" sqref="A1"/></dataValidations>',
      'Data-validation type is invalid',
    ],
    [
      '<dataValidations><dataValidation type="toString" sqref="A1"/></dataValidations>',
      'Data-validation type is invalid',
    ],
    [
      '<dataValidations><dataValidation operator="bad" sqref="A1"/></dataValidations>',
      'Data-validation operator is invalid',
    ],
    [
      '<dataValidations><dataValidation operator="constructor" sqref="A1"/></dataValidations>',
      'Data-validation operator is invalid',
    ],
    [
      '<dataValidations><dataValidation errorStyle="bad" sqref="A1"/></dataValidations>',
      'Data-validation error style is invalid',
    ],
    [
      '<dataValidations><dataValidation imeMode="bad" sqref="A1"/></dataValidations>',
      'Data-validation IME mode is invalid',
    ],
    [
      '<dataValidations><dataValidation imeMode="valueOf" sqref="A1"/></dataValidations>',
      'Data-validation IME mode is invalid',
    ],
    [
      '<dataValidations><dataValidation sqref="bad"/></dataValidations>',
      'Data-validation range is invalid',
    ],
    [
      '<dataValidations><dataValidation/></dataValidations>',
      'Data-validation range list is invalid',
    ],
    [
      '<dataValidations><dataValidation sqref=""/></dataValidations>',
      'Data-validation range list is invalid',
    ],
    [
      '<dataValidations><dataValidation sqref="$A$1"/></dataValidations>',
      'Data-validation range is invalid',
    ],
    [
      '<dataValidations><dataValidation sqref="A1 A1"/></dataValidations>',
      'Data-validation range list contains duplicates',
    ],
    [
      '<dataValidations><dataValidation type="list" sqref="A1"/></dataValidations>',
      'Data-validation first formula is missing',
    ],
    [
      '<dataValidations><dataValidation type="whole" sqref="A1"><formula1>1</formula1></dataValidation></dataValidations>',
      'Data-validation second formula is missing',
    ],
    [
      '<dataValidations><dataValidation type="whole" operator="notBetween" sqref="A1"><formula1>1</formula1></dataValidation></dataValidations>',
      'Data-validation second formula is missing',
    ],
    [
      '<dataValidations><dataValidation type="custom" sqref="A1"><formula1/></dataValidation></dataValidations>',
      'Data-validation first formula is invalid',
    ],
    [
      '<dataValidations><dataValidation type="whole" sqref="A1"><formula1>1</formula1><formula2/></dataValidation></dataValidations>',
      'Data-validation second formula is invalid',
    ],
    [
      '<dataValidations><dataValidation type="custom" sqref="A1"><formula1>1</formula1><formula1>2</formula1></dataValidation></dataValidations>',
      'Data-validation first formula is invalid',
    ],
    [
      '<dataValidations><dataValidation allowBlank="bad" sqref="A1"/></dataValidations>',
      'Data-validation allow-blank flag is invalid',
    ],
    [
      '<dataValidations><dataValidation showDropDown="bad" sqref="A1"/></dataValidations>',
      'Data-validation drop-down flag is invalid',
    ],
    [
      '<dataValidations><dataValidation showErrorMessage="bad" sqref="A1"/></dataValidations>',
      'Data-validation error-message flag is invalid',
    ],
    [
      '<dataValidations><dataValidation showInputMessage="bad" sqref="A1"/></dataValidations>',
      'Data-validation input-message flag is invalid',
    ],
    [
      '<dataValidations disablePrompts="bad"/>',
      'Data-validation disable-prompts flag is invalid',
    ],
    [
      '<dataValidations xWindow="01"/>',
      'Data-validation prompt X position is invalid',
    ],
    [
      '<dataValidations yWindow="4294967296"/>',
      'Data-validation prompt Y position is invalid',
    ],
    [
      '<dataValidations>bad</dataValidations>',
      'Data-validations text content is invalid',
    ],
    [
      '<dataValidations/><dataValidations/>',
      'Worksheet contains duplicate dataValidations elements',
    ],
  ] as const)(
    'rejects invalid data validation contract %#',
    async (xml, message) => {
      expect((await capture(xml)).diagnostic.message).toBe(message);
    },
  );
});

describe('XLSX normalized data-validation parser and capture', () => {
  it('decodes every XML entity and numeric-reference form', () => {
    expect(
      parseTree({
        attrs: { count: '1' },
        dataValidation: {
          attrs: {
            error: '&amp;&apos;&gt;&lt;&quot;&#65;&#x1F600;',
            sqref: 'A1',
          },
        },
      }).rules[0]?.error,
    ).toBe('&\'><"A😀');
  });

  it('rejects duplicate formulas as structure and normalized empty text as value', () => {
    expect(
      captureTree(() =>
        parseTree({
          attrs: { count: '1' },
          dataValidation: {
            attrs: { sqref: 'A1', type: 'custom' },
            formula1: [{ value: '1' }, { value: '2' }],
          },
        }),
      ).diagnostic,
    ).toMatchObject({
      code: 'invalid-document-structure',
      message: 'Data-validation first formula is invalid',
    });
    expect(
      captureTree(() =>
        parseTree({
          attrs: { count: '1' },
          dataValidation: {
            attrs: { sqref: 'A1', type: 'custom' },
            formula1: { value: '' },
          },
        }),
      ).diagnostic.message,
    ).toBe('Data-validation first formula is invalid');
  });

  it('accepts exact configured row and column bounds', () => {
    expect(
      parseTree(
        {
          attrs: { count: '1' },
          dataValidation: { attrs: { sqref: 'A1' } },
        },
        {
          limits: { maxColumnsPerWorksheet: 1, maxRowsPerWorksheet: 1 },
        },
      ).rules,
    ).toHaveLength(1);
  });

  it('distinguishes every validation selection rectangle boundary', () => {
    const value = {
      attrs: { count: '1' },
      dataValidation: { attrs: { sqref: 'B2:C3' } },
    };
    const selection = (reference: string): XlsxResolvedSheetSelection => {
      const column = reference.codePointAt(0)! - 0x40;
      const row = Number(reference.slice(1));
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
    for (const reference of ['A1', 'B1', 'D1', 'A2', 'D2', 'A4', 'B4', 'D4']) {
      expect(
        parseTree(value, { selection: selection(reference) }).rules,
      ).toEqual([]);
    }
    for (const reference of ['B2', 'C2', 'B3', 'C3']) {
      expect(
        parseTree(value, { selection: selection(reference) }).rules,
      ).toMatchObject([{ selectionRelation: 'intersects-selection' }]);
    }
    expect(
      parseTree(value, { selection: { kind: 'not-selected' } }).rules,
    ).toEqual([]);
  });

  it('returns explicit default collection settings', () => {
    expect(parseTree({ attrs: {} }).settings).toEqual({
      disablePrompts: false,
    });
  });

  it('accounts public text and selection work without emitting excluded rules', () => {
    const value = {
      attrs: { count: '1' },
      dataValidation: {
        attrs: {
          error: 'abc',
          prompt: 'def',
          sqref: 'A1 B2',
        },
      },
    };
    const selection: XlsxResolvedSheetSelection = {
      endRowPrefix: [3],
      kind: 'selected-ranges',
      ranges: [
        {
          end: { column: 3, row: 3 },
          reference: 'C3',
          start: { column: 3, row: 3 },
        },
      ],
    };
    expect(
      parseTree(value, {
        limits: { maxScannedCells: 2, maxTextCharacters: 6 },
        selection,
      }).rules,
    ).toEqual([]);
    expect(
      captureLimit(() =>
        parseTree(value, { limits: { maxTextCharacters: 5 } }),
      ),
    ).toMatchObject({
      actual: 6,
      limit: 5,
      limitName: 'maxTextCharacters',
    });
    expect(
      captureLimit(() =>
        parseTree(value, {
          limits: { maxScannedCells: 1 },
          selection,
        }),
      ),
    ).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxScannedCells',
    });
  });

  it('accepts maximum prompt coordinates and normalized XML whitespace', () => {
    expect(
      parseTree({
        attrs: { count: '1', xWindow: '4294967295', yWindow: '4294967295' },
        dataValidation: { attrs: { sqref: '\n A1  \t B2 \r' } },
      }),
    ).toMatchObject({
      rules: [{ ranges: [{ reference: 'A1' }, { reference: 'B2' }] }],
      settings: { xWindow: 0xffff_ffff, yWindow: 0xffff_ffff },
    });
  });

  it('captures formulas split across text chunks and complete settings', () => {
    const capture = new XlsxDataValidationsCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    capture.openElement(
      element('dataValidations', {
        count: '1',
        disablePrompts: '0',
        xWindow: '1',
      }),
    );
    capture.openElement(
      element('dataValidation', { sqref: 'A1', type: 'custom' }),
    );
    capture.openElement(element('formula1'));
    capture.text('A1');
    capture.text('>0');
    capture.closeElement(element('formula1'));
    capture.closeElement(element('dataValidation'));
    capture.closeElement(element('dataValidations'));
    expect(capture.result()).toMatchObject({
      rules: [{ formula1: 'A1>0', type: 'custom' }],
      settings: { disablePrompts: false, xWindow: 1 },
    });
  });

  it('rejects invalid normalized roots, children, and incomplete captures', () => {
    expect(captureTree(() => parseTree(undefined)).diagnostic.message).toBe(
      'Data-validations collection is invalid',
    );
    expect(
      captureTree(() => parseTree({ attrs: {}, dataValidation: 'bad' }))
        .diagnostic.message,
    ).toBe('Data-validation rule collection is invalid');
    const capture = new XlsxDataValidationsCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(
      captureTree(() => capture.openElement(element('wrong'))).diagnostic
        .message,
    ).toBe('Data-validations capture root is invalid');
    const incomplete = new XlsxDataValidationsCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    incomplete.openElement(element('dataValidations'));
    expect(captureTree(() => incomplete.result()).diagnostic.message).toBe(
      'Data-validations capture is incomplete',
    );

    const duplicateRoot = new XlsxDataValidationsCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    duplicateRoot.openElement(element('dataValidations'));
    duplicateRoot.closeElement(element('dataValidations'));
    expect(
      captureTree(() => duplicateRoot.openElement(element('dataValidations')))
        .diagnostic.message,
    ).toBe('Data-validations capture root is invalid');

    const nesting = new XlsxDataValidationsCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    nesting.openElement(element('dataValidations'));
    expect(
      captureTree(() => nesting.closeElement(element('wrong'))).diagnostic
        .message,
    ).toBe('Data-validations capture nesting is invalid');
  });

  it('preserves raw expanded attribute keys during capture', () => {
    const capture = new XlsxDataValidationsCapture(
      FULL,
      budget(),
      defaultXlsxResourceLimits(),
      PART,
    );
    const root = element('dataValidations');
    root.attributes = new Map([['count', '2']]);
    capture.openElement(root);
    capture.openElement(element('dataValidation', { sqref: 'A1' }));
    capture.closeElement(element('dataValidation'));
    capture.closeElement(element('dataValidations'));
    expect(captureTree(() => capture.result()).diagnostic.message).toBe(
      'Data-validation count does not match',
    );
  });
});
