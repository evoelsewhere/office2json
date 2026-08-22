import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { parseXlsxXfFormatting } from '../../src/formats/xlsx/internal/style-formatting';

const PART = 'xl/styles.xml';

function capture(value: unknown, prefix = ''): XlsxParseError {
  try {
    parseXlsxXfFormatting(value, prefix, PART);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected XF formatting parsing to fail');
}

describe('XLSX XF formatting', () => {
  it('normalizes alignment and protection into frozen portable JSON', () => {
    const result = parseXlsxXfFormatting(
      {
        alignment: {
          attrs: {
            horizontal: 'distributed',
            indent: '250',
            justifyLastLine: 'true',
            readingOrder: '2',
            relativeIndent: '-15',
            shrinkToFit: '1',
            textRotation: '180',
            vertical: 'top',
            wrapText: 'true',
          },
        },
        protection: { attrs: { hidden: '1', locked: 'false' } },
      },
      '',
      PART,
    );

    expect(result).toEqual({
      alignment: {
        horizontal: 'distributed',
        indent: 250,
        justifyLastLine: true,
        readingOrder: 'right-to-left',
        relativeIndent: -15,
        shrinkToFit: true,
        textRotation: 180,
        vertical: 'top',
        wrapText: true,
      },
      protection: { hidden: true, locked: false },
    });
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.alignment)).toBe(true);
    expect(Object.isFrozen(result.protection)).toBe(true);
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });

  it('parses prefixed one-item arrays and omits semantic defaults', () => {
    expect(
      parseXlsxXfFormatting(
        {
          's:alignment': [
            {
              attrs: {
                horizontal: 'general',
                indent: '0',
                justifyLastLine: '0',
                readingOrder: '0',
                relativeIndent: '+0',
                shrinkToFit: 'false',
                textRotation: '0',
                vertical: 'bottom',
                wrapText: 'false',
              },
            },
          ],
          's:protection': [{ attrs: { hidden: 'false', locked: 'true' } }],
        },
        's',
        PART,
      ),
    ).toEqual({});
  });

  it('parses non-default prefixed formatting and empty attribute sets', () => {
    expect(
      parseXlsxXfFormatting(
        {
          's:alignment': { attrs: { horizontal: 'left' } },
          's:protection': { attrs: { hidden: 'true' } },
        },
        's',
        PART,
      ),
    ).toEqual({
      alignment: { horizontal: 'left' },
      protection: { hidden: true },
    });
    expect(
      parseXlsxXfFormatting({ alignment: {}, protection: {} }, '', PART),
    ).toEqual({});
  });

  it.each([
    ['1', 1],
    ['15', 15],
  ] as const)('accepts relativeIndent boundary %s', (source, expected) => {
    expect(
      parseXlsxXfFormatting(
        { alignment: { attrs: { relativeIndent: source } } },
        '',
        PART,
      ).alignment?.relativeIndent,
    ).toBe(expected);
  });

  it.each([
    ['left', 'left'],
    ['center', 'center'],
    ['right', 'right'],
    ['fill', 'fill'],
    ['justify', 'justify'],
    ['centerContinuous', 'centerContinuous'],
    ['distributed', 'distributed'],
  ] as const)('preserves horizontal alignment %s', (source, expected) => {
    expect(
      parseXlsxXfFormatting(
        { alignment: { attrs: { horizontal: source } } },
        '',
        PART,
      ).alignment?.horizontal,
    ).toBe(expected);
  });

  it.each([
    ['top', 'top'],
    ['center', 'center'],
    ['justify', 'justify'],
    ['distributed', 'distributed'],
  ] as const)('preserves vertical alignment %s', (source, expected) => {
    expect(
      parseXlsxXfFormatting(
        { alignment: { attrs: { vertical: source } } },
        '',
        PART,
      ).alignment?.vertical,
    ).toBe(expected);
  });

  it.each([
    ['1', 'left-to-right'],
    ['2', 'right-to-left'],
  ] as const)('normalizes reading order %s', (source, expected) => {
    expect(
      parseXlsxXfFormatting(
        { alignment: { attrs: { readingOrder: source } } },
        '',
        PART,
      ).alignment?.readingOrder,
    ).toBe(expected);
  });

  it.each(['1', 'true'] as const)(
    'normalizes true boolean value %s',
    (source) => {
      const result = parseXlsxXfFormatting(
        {
          alignment: {
            attrs: {
              justifyLastLine: source,
              shrinkToFit: source,
              wrapText: source,
            },
          },
        },
        '',
        PART,
      );
      expect(result.alignment).toEqual({
        justifyLastLine: true,
        shrinkToFit: true,
        wrapText: true,
      });
    },
  );

  it.each([
    ['1', 1],
    ['180', 180],
    ['255', 255],
  ] as const)('accepts text rotation %s', (source, expected) => {
    expect(
      parseXlsxXfFormatting(
        { alignment: { attrs: { textRotation: source } } },
        '',
        PART,
      ).alignment?.textRotation,
    ).toBe(expected);
  });

  it.each([
    [null, 'invalid-document-structure', 'XF element is invalid'],
    [
      { alignment: [] },
      'invalid-document-structure',
      'XF alignment element is duplicated',
    ],
    [
      { alignment: [{}, {}] },
      'invalid-document-structure',
      'XF alignment element is duplicated',
    ],
    [
      { alignment: ['bad'] },
      'invalid-document-structure',
      'XF alignment element is invalid',
    ],
    [
      { alignment: 'bad' },
      'invalid-document-structure',
      'XF alignment element is invalid',
    ],
    [
      { alignment: { attrs: 'bad' } },
      'invalid-document-structure',
      'Alignment attributes are invalid',
    ],
    [
      { protection: { attrs: 'bad' } },
      'invalid-document-structure',
      'Protection attributes are invalid',
    ],
    [
      { alignment: { attrs: { horizontal: 'middle' } } },
      'invalid-document-value',
      'Alignment horizontal value is invalid',
    ],
    [
      { alignment: { attrs: { horizontal: 1 } } },
      'invalid-document-value',
      'Alignment horizontal value is invalid',
    ],
    [
      { alignment: { attrs: { vertical: 'baseline' } } },
      'invalid-document-value',
      'Alignment vertical value is invalid',
    ],
    [
      { alignment: { attrs: { vertical: 1 } } },
      'invalid-document-value',
      'Alignment vertical value is invalid',
    ],
    [
      { alignment: { attrs: { wrapText: 'yes' } } },
      'invalid-document-value',
      'Alignment wrapText value is invalid',
    ],
    [
      { alignment: { attrs: { shrinkToFit: 'yes' } } },
      'invalid-document-value',
      'Alignment shrinkToFit value is invalid',
    ],
    [
      { alignment: { attrs: { justifyLastLine: 'yes' } } },
      'invalid-document-value',
      'Alignment justifyLastLine value is invalid',
    ],
    [
      { alignment: { attrs: { indent: '-1' } } },
      'invalid-document-value',
      'Alignment indent value is invalid',
    ],
    [
      { alignment: { attrs: { indent: '01' } } },
      'invalid-document-value',
      'Alignment indent value is invalid',
    ],
    [
      { alignment: { attrs: { indent: '251' } } },
      'invalid-document-value',
      'Alignment indent value is invalid',
    ],
    [
      { alignment: { attrs: { relativeIndent: '-16' } } },
      'invalid-document-value',
      'Alignment relativeIndent value is invalid',
    ],
    [
      { alignment: { attrs: { relativeIndent: '16' } } },
      'invalid-document-value',
      'Alignment relativeIndent value is invalid',
    ],
    [
      { alignment: { attrs: { relativeIndent: '01' } } },
      'invalid-document-value',
      'Alignment relativeIndent value is invalid',
    ],
    [
      { alignment: { attrs: { readingOrder: '3' } } },
      'invalid-document-value',
      'Alignment readingOrder value is invalid',
    ],
    [
      { alignment: { attrs: { textRotation: '181' } } },
      'invalid-document-value',
      'Alignment textRotation value is invalid',
    ],
    [
      { alignment: { attrs: { textRotation: '254' } } },
      'invalid-document-value',
      'Alignment textRotation value is invalid',
    ],
    [
      { alignment: { attrs: { textRotation: '256' } } },
      'invalid-document-value',
      'Alignment textRotation value is invalid',
    ],
    [
      { protection: { attrs: { hidden: 'yes' } } },
      'invalid-document-value',
      'Protection hidden value is invalid',
    ],
    [
      { protection: { attrs: { locked: 'yes' } } },
      'invalid-document-value',
      'Protection locked value is invalid',
    ],
    [
      { protection: [] },
      'invalid-document-structure',
      'XF protection element is duplicated',
    ],
    [
      { protection: 'bad' },
      'invalid-document-structure',
      'XF protection element is invalid',
    ],
  ] as const)('rejects invalid XF formatting %#', (source, code, message) => {
    expect(capture(source).diagnostic).toEqual({
      code,
      message,
      part: PART,
      severity: 'error',
    });
  });
});
