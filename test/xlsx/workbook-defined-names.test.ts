import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from '../../src/formats/xlsx/internal/resource-limits';
import {
  isValidXlsxDefinedName,
  parseXlsxDefinedNames,
  xlsxDefinedNameFormulaCharacters,
  xlsxDefinedNameTextCharacters,
} from '../../src/formats/xlsx/internal/workbook-defined-names';

const PART = 'xl/workbook.xml';

function limits(
  overrides: Partial<ResolvedXlsxResourceLimits> = {},
): ResolvedXlsxResourceLimits {
  return { ...defaultXlsxResourceLimits(), ...overrides };
}

function definedName(
  attrs: Record<string, unknown>,
  value: unknown = 'Sheet1!$A$1',
): Record<string, unknown> {
  return { attrs, value };
}

function collection(...values: unknown[]): Record<string, unknown> {
  return { definedName: values.length === 1 ? values[0] : values };
}

function parse(
  value: unknown,
  options: Partial<ResolvedXlsxResourceLimits> = {},
  sheetCount = 2,
) {
  return parseXlsxDefinedNames(value, '', PART, sheetCount, limits(options));
}

function capture(
  value: unknown,
  options: Partial<ResolvedXlsxResourceLimits> = {},
  sheetCount = 2,
): XlsxParseError {
  try {
    parse(value, options, sheetCount);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected defined-name parsing to fail');
}

function captureResource(action: () => unknown): XlsxResourceLimitError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxResourceLimitError);
    return error as XlsxResourceLimitError;
  }
  throw new Error('Expected defined-name resource limit to fail');
}

describe('XLSX workbook defined names', () => {
  it('returns empty portable accounting when the collection is absent', () => {
    expect(parse(undefined)).toEqual({
      definedNames: [],
      formulaCharacters: 0,
      textCharacters: 0,
    });
  });

  it('normalizes workbook, sheet, builtin, Unicode, and function metadata', () => {
    const result = parse(
      collection(
        definedName(
          {
            comment: 'Comment',
            customMenu: 'Menu',
            description: 'Description',
            function: 'true',
            functionGroupId: '4294967295',
            help: 'Help',
            hidden: '1',
            name: 'Sales.Total_2026',
            publishToServer: 'false',
            shortcutKey: 'S',
            statusBar: 'Ready',
            vbProcedure: '0',
            workbookParameter: '1',
            xlm: 'false',
          },
          'SUM(Sheet1!$A$1:$A$2)',
        ),
        definedName(
          { localSheetId: '0', name: 'Sales.Total_2026' },
          'Sheet1!$A$1',
        ),
        definedName({ name: '_xlnm.Print_Area' }, 'Sheet1!$A$1:$D$20'),
        definedName({ name: '売上_2026' }, 'Sheet1!$B$2'),
        definedName({ name: '\\LocalName' }, 'Sheet2!$C$3'),
      ),
    );

    expect(result.definedNames).toEqual([
      {
        comment: 'Comment',
        customMenu: 'Menu',
        description: 'Description',
        expression: 'SUM(Sheet1!$A$1:$A$2)',
        function: true,
        functionGroupId: 4_294_967_295,
        help: 'Help',
        hidden: true,
        name: 'Sales.Total_2026',
        publishToServer: false,
        shortcutKey: 'S',
        statusBar: 'Ready',
        vbProcedure: false,
        workbookParameter: true,
        xlm: false,
      },
      {
        expression: 'Sheet1!$A$1',
        hidden: false,
        name: 'Sales.Total_2026',
        sheetIndex: 0,
      },
      {
        expression: 'Sheet1!$A$1:$D$20',
        hidden: false,
        name: '_xlnm.Print_Area',
      },
      {
        expression: 'Sheet1!$B$2',
        hidden: false,
        name: '売上_2026',
      },
      {
        expression: 'Sheet2!$C$3',
        hidden: false,
        name: '\\LocalName',
      },
    ]);
    expect(result.formulaCharacters).toBe(
      result.definedNames.reduce(
        (total, value) => total + value.expression.length,
        0,
      ),
    );
    expect(result.textCharacters).toBe(
      xlsxDefinedNameTextCharacters(result.definedNames),
    );
    expect(xlsxDefinedNameFormulaCharacters(result.definedNames)).toBe(
      result.formulaCharacters,
    );
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });

  it('preserves explicit false and true defined-name flags', () => {
    const result = parse(
      collection(
        definedName({
          function: '0',
          hidden: 'false',
          name: 'Flags',
          publishToServer: '1',
          vbProcedure: 'true',
          workbookParameter: '0',
          xlm: '1',
        }),
      ),
    );
    expect(result.definedNames[0]).toMatchObject({
      function: false,
      hidden: false,
      publishToServer: true,
      vbProcedure: true,
      workbookParameter: false,
      xlm: true,
    });
  });

  it('decodes predefined, decimal, and hexadecimal XML references once', () => {
    expect(
      parse(
        collection(
          definedName(
            {
              comment: '&amp;&apos;&gt;&lt;&quot;&#32;&#x42;',
              name: 'N&#97;me',
            },
            'A&amp;B',
          ),
        ),
      ).definedNames,
    ).toEqual([
      {
        comment: `&'><" B`,
        expression: 'A&B',
        hidden: false,
        name: 'Name',
      },
    ]);
  });

  it.each([
    [null, 'Workbook defined-names collection is invalid'],
    ['names', 'Workbook defined-names collection is invalid'],
    [{}, 'Workbook defined-names collection is empty'],
    [{ definedName: [] }, 'Workbook defined-names collection is empty'],
    [{ definedName: 'text' }, 'Workbook defined-names collection is empty'],
  ])('rejects invalid defined-name collection %#', (value, message) => {
    expect(capture(value).diagnostic).toMatchObject({
      code: 'invalid-document-structure',
      message,
      part: PART,
    });
  });

  it.each([
    undefined,
    '',
    'A1',
    'R',
    'c',
    'R1C1',
    '1Name',
    '.Name',
    'Has Space',
    'Has-Dash',
    'a'.repeat(256),
  ])('rejects invalid defined name %#', (name) => {
    expect(capture(collection(definedName({ name }))).diagnostic.message).toBe(
      'Workbook defined name is invalid',
    );
  });

  it('accepts the exact name-length boundary and validates names directly', () => {
    const name = `A${'a'.repeat(254)}`;
    expect(isValidXlsxDefinedName(name)).toBe(true);
    expect(parse(collection(definedName({ name }))).definedNames[0]?.name).toBe(
      name,
    );
    expect(isValidXlsxDefinedName(undefined)).toBe(false);
    expect(isValidXlsxDefinedName(1)).toBe(false);
    expect(isValidXlsxDefinedName('R12C34')).toBe(false);
    expect(isValidXlsxDefinedName('NameR1C1')).toBe(true);
    expect(isValidXlsxDefinedName('R1C1Name')).toBe(true);
  });

  it.each([
    { attrs: { name: 'Name' } },
    definedName({ name: 'Name' }, ''),
    definedName({ name: 'Name' }, { nested: 'formula' }),
  ])('rejects invalid defined-name expression %#', (node) => {
    expect(capture(collection(node)).diagnostic.message).toBe(
      'Workbook defined-name expression is missing',
    );
  });

  it('allows equal names in distinct scopes and rejects same-scope collisions', () => {
    expect(
      parse(
        collection(
          definedName({ localSheetId: '0', name: 'Local' }),
          definedName({ localSheetId: '1', name: 'local' }),
        ),
      ).definedNames,
    ).toHaveLength(2);
    expect(
      parse(
        collection(
          definedName({ name: 'GlobalAndLocal' }),
          definedName({ localSheetId: '1', name: 'GlobalAndLocal' }),
        ),
      ).definedNames,
    ).toHaveLength(2);
    expect(
      capture(
        collection(
          definedName({ localSheetId: '0', name: 'Local' }),
          definedName({ localSheetId: '0', name: 'local' }),
        ),
      ).diagnostic.message,
    ).toBe('Workbook contains duplicate defined names in one scope');
    expect(
      capture(
        collection(
          definedName({ name: 'Global' }),
          definedName({ name: 'global' }),
        ),
      ).diagnostic.message,
    ).toBe('Workbook contains duplicate defined names in one scope');
    expect(
      capture(
        collection(definedName({ name: 'σ' }), definedName({ name: 'ς' })),
      ).diagnostic.message,
    ).toBe('Workbook contains duplicate defined names in one scope');
  });

  it('rejects non-string numeric and metadata attributes', () => {
    expect(
      capture(collection(definedName({ localSheetId: 0, name: 'Scope' })))
        .diagnostic.message,
    ).toBe('Workbook defined-name sheet scope is invalid');
    expect(
      capture(collection(definedName({ functionGroupId: 1, name: 'Function' })))
        .diagnostic.message,
    ).toBe('Workbook defined-name function group is invalid');
    expect(
      capture(collection(definedName({ comment: 1, name: 'Metadata' })))
        .diagnostic.message,
    ).toBe('Workbook defined-name metadata is invalid');
  });

  it.each(['-1', '01', '2', '4294967296'])(
    'rejects sheet scope %s',
    (scope) => {
      expect(
        capture(
          collection(definedName({ localSheetId: scope, name: 'Scoped' })),
        ).diagnostic.message,
      ).toBe('Workbook defined-name sheet scope is invalid');
    },
  );

  it.each(['yes', 'TRUE', '2'])(
    'rejects invalid defined-name boolean %s',
    (flag) => {
      expect(
        capture(collection(definedName({ hidden: flag, name: 'Flag' })))
          .diagnostic.message,
      ).toBe('Workbook defined-name hidden flag is invalid');
      expect(
        capture(collection(definedName({ function: flag, name: 'Flag' })))
          .diagnostic.message,
      ).toBe('Workbook defined-name flag is invalid');
    },
  );

  it.each(['-1', '01', '4294967296'])(
    'rejects invalid function group %s',
    (functionGroupId) => {
      expect(
        capture(collection(definedName({ functionGroupId, name: 'Function' })))
          .diagnostic.message,
      ).toBe('Workbook defined-name function group is invalid');
    },
  );

  it('enforces defined-name count exactly', () => {
    const value = collection(
      definedName({ name: 'One' }),
      definedName({ name: 'Two' }),
    );
    expect(parse(value, { maxDefinedNames: 2 }).definedNames).toHaveLength(2);
    expect(
      captureResource(() => parse(value, { maxDefinedNames: 1 })),
    ).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxDefinedNames',
      name: 'XlsxResourceLimitError',
      part: PART,
    });
  });

  it('enforces per-formula and aggregate formula boundaries', () => {
    const value = collection(
      definedName({ name: 'One' }, '123'),
      definedName({ name: 'Two' }, '456'),
    );
    expect(
      parse(value, {
        maxFormulaCharacters: 3,
        maxTotalFormulaCharacters: 6,
      }).formulaCharacters,
    ).toBe(6);
    expect(
      captureResource(() =>
        parse(value, {
          maxFormulaCharacters: 2,
          maxTotalFormulaCharacters: 6,
        }),
      ),
    ).toMatchObject({
      actual: 3,
      limit: 2,
      limitName: 'maxFormulaCharacters',
    });
    expect(
      captureResource(() =>
        parse(value, {
          maxFormulaCharacters: 3,
          maxTotalFormulaCharacters: 5,
        }),
      ),
    ).toMatchObject({
      actual: 6,
      limit: 5,
      limitName: 'maxTotalFormulaCharacters',
    });
  });

  it('enforces normalized name and metadata text boundaries', () => {
    const value = collection(definedName({ comment: 'xy', name: 'Abc' }, '1'));
    expect(parse(value, { maxTextCharacters: 5 }).textCharacters).toBe(5);
    expect(
      captureResource(() => parse(value, { maxTextCharacters: 4 })),
    ).toMatchObject({
      actual: 5,
      limit: 4,
      limitName: 'maxTextCharacters',
    });
  });
});
