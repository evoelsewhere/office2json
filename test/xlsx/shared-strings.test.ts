import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import type { XlsxDiagnostic } from '../../src/formats/xlsx/types';
import { XlsxPartReader } from '../../src/formats/xlsx/internal/part-reader';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from '../../src/formats/xlsx/internal/resource-limits';
import { parseXlsxSharedStringPart } from '../../src/formats/xlsx/internal/shared-strings';
import {
  XLSX_SPREADSHEET_NAMESPACES,
  type XlsxWorkbookDiscovery,
} from '../../src/formats/xlsx/internal/workbook-discovery';

const PART = 'custom/tables/strings.xml';

function sharedStrings(
  body: string,
  attributes = '',
  namespace = XLSX_SPREADSHEET_NAMESPACES.transitional,
): string {
  return `<sst xmlns="${namespace}"${attributes}>${body}</sst>`;
}

async function parse(
  xml: string,
  limitOverrides: Partial<ResolvedXlsxResourceLimits> = {},
  dialect: XlsxWorkbookDiscovery['dialect'] = 'transitional',
) {
  const zip = new JSZip();
  zip.file(PART, xml);
  const limits = { ...defaultXlsxResourceLimits(), ...limitOverrides };
  const reader = new XlsxPartReader(zip, [], limits);
  return parseXlsxSharedStringPart(PART, dialect, reader, limits);
}

async function captureParseError(xml: string): Promise<XlsxParseError> {
  try {
    await parse(xml);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected shared-string parsing to fail');
}

describe('XLSX shared strings', () => {
  it('preserves plain, rich, empty, and phonetic strings immutably', async () => {
    const result = await parse(
      sharedStrings(
        `<si><t xml:space="preserve"> A &amp; <![CDATA[B]]> </t></si>
         <si>
           <r><rPr><b/><color rgb="FF000000"/></rPr><t>Rich</t></r>
           <r><t xml:space="preserve"> text</t></r>
         </si>
         <si>
           <t>東京</t>
           <rPh sb="0" eb="2"><t>とうきょう</t></rPh>
           <phoneticPr fontId="2" type="Hiragana" alignment="center"/>
         </si>
         <si/>`,
        ' count="5" uniqueCount="4"',
      ),
    );

    expect(result).toEqual({
      part: PART,
      values: [
        { text: ' A & B ' },
        {
          runs: [{ text: 'Rich' }, { text: ' text' }],
          text: 'Rich text',
        },
        {
          phoneticProperties: {
            alignment: 'center',
            fontId: 2,
            type: 'hiragana',
          },
          phoneticRuns: [{ end: 2, start: 0, text: 'とうきょう' }],
          text: '東京',
        },
        { text: '' },
      ],
    });
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.values)).toBe(true);
    expect(Object.isFrozen(result.values[1])).toBe(true);
    expect(Object.isFrozen(result.values[1]?.runs)).toBe(true);
    expect(Object.isFrozen(result.values[2]?.phoneticRuns)).toBe(true);
    expect(Object.isFrozen(result.values[2]?.phoneticProperties)).toBe(true);
  });

  it('parses prefixed Strict shared strings and ignores formatting payload text', async () => {
    const namespace = XLSX_SPREADSHEET_NAMESPACES.strict;
    const result = await parse(
      `<s:sst xmlns:s="${namespace}">
        <s:si><s:r><s:rPr><s:rFont val="Ignored"/></s:rPr><s:t>Strict</s:t></s:r></s:si>
      </s:sst>`,
      {},
      'strict',
    );

    expect(result.values).toEqual([
      { runs: [{ text: 'Strict' }], text: 'Strict' },
    ]);
  });

  it('accepts every supported rich-text run property and default xml:space', async () => {
    const properties = [
      'b',
      'charset',
      'color',
      'condense',
      'extend',
      'family',
      'i',
      'outline',
      'rFont',
      'scheme',
      'shadow',
      'strike',
      'sz',
      'u',
      'vertAlign',
    ];
    const result = await parse(
      sharedStrings(
        `<si><r><rPr>${properties.map((name) => `<${name}/>`).join('')}</rPr><t xml:space="default">A</t></r></si>`,
      ),
    );

    expect(result.values).toEqual([{ runs: [{ text: 'A' }], text: 'A' }]);
  });

  it.each([
    ['noControl', 'no-control'],
    ['left', 'left'],
    ['center', 'center'],
    ['distributed', 'distributed'],
  ] as const)('normalizes phonetic alignment %s', async (source, expected) => {
    const result = await parse(
      sharedStrings(`<si><t>A</t><phoneticPr alignment="${source}"/></si>`),
    );
    expect(result.values[0]?.phoneticProperties?.alignment).toBe(expected);
  });

  it.each([
    ['fullwidthKatakana', 'full-width-katakana'],
    ['halfwidthKatakana', 'half-width-katakana'],
    ['Hiragana', 'hiragana'],
    ['noConversion', 'no-conversion'],
  ] as const)('normalizes phonetic type %s', async (source, expected) => {
    const result = await parse(
      sharedStrings(`<si><t>A</t><phoneticPr type="${source}"/></si>`),
    );
    expect(result.values[0]?.phoneticProperties?.type).toBe(expected);
  });

  it('accepts empty tables and unsigned metadata boundaries', async () => {
    await expect(
      parse(sharedStrings('', ' count="0" uniqueCount="0"')),
    ).resolves.toEqual({ part: PART, values: [] });
    await expect(
      parse(
        sharedStrings(
          '<si><t>A</t><phoneticPr fontId="4294967295"/></si>',
          ' count="4294967295" uniqueCount="1"',
        ),
      ),
    ).resolves.toMatchObject({
      values: [{ phoneticProperties: { fontId: 0xffff_ffff } }],
    });
  });

  it.each([
    `<root xmlns="${XLSX_SPREADSHEET_NAMESPACES.transitional}"/>`,
    '<sst xmlns="urn:wrong"><si><t>A</t></si></sst>',
    sharedStrings('<unknown/>'),
    sharedStrings('<si xmlns:x="urn:x"><x:t>A</x:t></si>'),
    sharedStrings('<si>raw</si>'),
    sharedStrings('<si><t><b/></t></si>'),
    sharedStrings('<si><t>A</t><t>B</t></si>'),
    sharedStrings('<si><t>A</t><r><t>B</t></r></si>'),
    sharedStrings('<si><r><t>A</t></r><t>B</t></si>'),
    sharedStrings('<si><r/></si>'),
    sharedStrings('<si><r><t>A</t><t>B</t></r></si>'),
    sharedStrings('<si><r><r/></r></si>'),
    sharedStrings('<si><r><si/></r></si>'),
    sharedStrings('<si><r><rPr><si/></rPr><t>A</t></r></si>'),
    sharedStrings('<si><unknown/></si>'),
    sharedStrings('<si><r><unknown/></r></si>'),
    sharedStrings('<si><r><rPr><unknown/></rPr><t>A</t></r></si>'),
    sharedStrings('<si><r><rPr><b>text</b></rPr><t>A</t></r></si>'),
    sharedStrings('<si><r><rPr><b><i/></b></rPr><t>A</t></r></si>'),
    sharedStrings('<si><t>A</t><rPh sb="0" eb="1"/></si>'),
    sharedStrings('<si><t>A</t><rPh sb="0" eb="1"><unknown/></rPh></si>'),
    sharedStrings(
      '<si><t>A</t><phoneticPr/><rPh sb="0" eb="1"><t>a</t></rPh></si>',
    ),
    sharedStrings('<si><phoneticPr><t>A</t></phoneticPr></si>'),
    sharedStrings('<si><t>A</t><phoneticPr/><phoneticPr/></si>'),
    sharedStrings(
      '<si><t>A</t><rPh sb="0" eb="1"><t>a</t></rPh><r><t>B</t></r></si>',
    ),
    sharedStrings('<si><rPh sb="0" eb="1"><t>a</t></rPh><r><t>B</t></r></si>'),
    sharedStrings('<si><phoneticPr/><r><t>B</t></r></si>'),
  ])('rejects malformed shared-string structure %#', async (xml) => {
    const error = await captureParseError(xml);
    expect(error.diagnostic).toMatchObject({
      code: 'invalid-document-structure',
      part: PART,
      severity: 'error',
    });
  });

  it.each([
    sharedStrings('', ' count=""'),
    sharedStrings('', ' count="-1"'),
    sharedStrings('', ' count="01"'),
    sharedStrings('', ' count="1.0"'),
    sharedStrings('', ' count="4294967296"'),
    sharedStrings('', ' count="9007199254740992"'),
    sharedStrings('<si/>', ' uniqueCount="0"'),
    sharedStrings('<si/>', ' count="0" uniqueCount="1"'),
    sharedStrings('<si><t xml:space="invalid">A</t></si>'),
    sharedStrings('<si><t>A</t><rPh eb="1"><t>a</t></rPh></si>'),
    sharedStrings('<si><t>A</t><rPh sb="0"><t>a</t></rPh></si>'),
    sharedStrings('<si><t>A</t><rPh sb="00" eb="1"><t>a</t></rPh></si>'),
    sharedStrings('<si><t>A</t><rPh sb="0" eb="0"><t>a</t></rPh></si>'),
    sharedStrings('<si><t>A</t><rPh sb="1" eb="1"><t>a</t></rPh></si>'),
    sharedStrings('<si><t>A</t><rPh sb="1" eb="2"><t>a</t></rPh></si>'),
    sharedStrings('<si><t>A</t><phoneticPr fontId="-1"/></si>'),
    sharedStrings('<si><t>A</t><phoneticPr alignment="right"/></si>'),
    sharedStrings('<si><t>A</t><phoneticPr type="unknown"/></si>'),
  ])('rejects malformed shared-string values %#', async (xml) => {
    const error = await captureParseError(xml);
    expect(error.diagnostic).toMatchObject({
      code: 'invalid-document-value',
      part: PART,
      severity: 'error',
    });
  });

  it.each([
    [
      '<sst xmlns="urn:wrong"/>',
      'Shared-string element has an unsupported namespace',
    ],
    [
      `<root xmlns="${XLSX_SPREADSHEET_NAMESPACES.transitional}"/>`,
      'Shared-string root is missing',
    ],
    [sharedStrings('', ' count="bad"'), 'Shared-string count is invalid'],
    [
      sharedStrings('', ' uniqueCount="bad"'),
      'Shared-string unique count is invalid',
    ],
    [
      sharedStrings('<si>raw</si>'),
      'Shared-string text must be contained by a text element',
    ],
    [
      sharedStrings('<si/>', ' uniqueCount="0"'),
      'Shared-string unique count does not match entries',
    ],
    [
      sharedStrings('<si/>', ' count="0" uniqueCount="1"'),
      'Shared-string count is smaller than its table',
    ],
    [sharedStrings('<unknown/>'), 'Shared-string element nesting is invalid'],
    [
      sharedStrings('<si><t>A</t><t>B</t></si>'),
      'Shared-string plain text is out of order',
    ],
    [
      sharedStrings('<si><t>A</t><r><t>B</t></r></si>'),
      'Shared-string rich text is out of order',
    ],
    [
      sharedStrings(
        '<si><t>A</t><phoneticPr/><rPh sb="0" eb="1"><t>a</t></rPh></si>',
      ),
      'Shared-string phonetic run is out of order',
    ],
    [
      sharedStrings('<si><t>A</t><rPh sb="0"><t>a</t></rPh></si>'),
      'Shared-string phonetic end index is invalid',
    ],
    [
      sharedStrings('<si><t>A</t><rPh eb="1"><t>a</t></rPh></si>'),
      'Shared-string phonetic start index is invalid',
    ],
    [
      sharedStrings('<si><t>A</t><phoneticPr/><phoneticPr/></si>'),
      'Shared string has duplicate phonetic properties',
    ],
    [
      sharedStrings('<si><t>A</t><phoneticPr fontId="bad"/></si>'),
      'Shared-string phonetic font ID is invalid',
    ],
    [
      sharedStrings('<si><t>A</t><phoneticPr alignment="right"/></si>'),
      'Shared-string phonetic alignment is invalid',
    ],
    [
      sharedStrings('<si><t>A</t><phoneticPr type="unknown"/></si>'),
      'Shared-string phonetic type is invalid',
    ],
    [
      sharedStrings('<si><r><t>A</t><t>B</t></r></si>'),
      'Shared-string run text is invalid',
    ],
    [sharedStrings('<si><r/></si>'), 'Shared-string rich run has no text'],
    [
      sharedStrings('<si><t>A</t><rPh sb="0" eb="1"/></si>'),
      'Shared-string phonetic run has no text',
    ],
    [
      sharedStrings('<si><t>A</t><rPh sb="0" eb="2"><t>a</t></rPh></si>'),
      'Shared-string phonetic range is invalid',
    ],
    [
      sharedStrings('<si><t xml:space="invalid">A</t></si>'),
      'Shared-string xml:space value is invalid',
    ],
  ] as const)('reports stable shared-string error %#', async (xml, message) => {
    const error = await captureParseError(xml);
    expect(error.message).toBe(message);
    expect(error.diagnostic.message).toBe(message);
  });

  it('accepts exact shared-string, run, and text budgets', async () => {
    const xml = sharedStrings(
      '<si><r><t>AB</t></r><rPh sb="0" eb="2"><t>cde</t></rPh></si>',
    );
    await expect(
      parse(xml, {
        maxRichTextRuns: 2,
        maxSharedStrings: 1,
        maxTextCharacters: 5,
      }),
    ).resolves.toMatchObject({ values: [{ text: 'AB' }] });
  });

  it.each([
    [
      sharedStrings('<si/><si/>'),
      { maxSharedStrings: 1 },
      'maxSharedStrings',
      2,
      1,
    ],
    [
      sharedStrings('<si><r><t>A</t></r><r><t>B</t></r></si>'),
      { maxRichTextRuns: 1 },
      'maxRichTextRuns',
      2,
      1,
    ],
    [
      sharedStrings('<si><t>AB</t><rPh sb="0" eb="2"><t>cde</t></rPh></si>'),
      { maxTextCharacters: 4 },
      'maxTextCharacters',
      5,
      4,
    ],
  ] as const)(
    'enforces shared-string resource limit %#',
    async (xml, overrides, limitName, actual, limit) => {
      let thrown: unknown;
      try {
        await parse(xml, overrides);
      } catch (error) {
        thrown = error;
      }
      expect(thrown).toBeInstanceOf(XlsxResourceLimitError);
      expect(thrown).toMatchObject({
        actual,
        limit,
        limitName,
        part: PART,
      });
    },
  );

  it('requires the resolved shared-string part to exist', async () => {
    const limits = defaultXlsxResourceLimits();
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(new JSZip(), diagnostics, limits);

    await expect(
      parseXlsxSharedStringPart(PART, 'transitional', reader, limits),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'missing-required-part',
        part: PART,
      },
    });
    expect(diagnostics).toHaveLength(1);
  });
});
