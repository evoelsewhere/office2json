import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { parseXlsxDifferentialStyle } from '../../src/formats/xlsx/internal/style-differential';

const PART = 'xl/styles.xml';

function capture(value: unknown, prefix = ''): XlsxParseError {
  try {
    parseXlsxDifferentialStyle(value, prefix, PART);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected differential style parsing to fail');
}

describe('XLSX differential styles', () => {
  it('normalizes every inline differential component into portable JSON', () => {
    const result = parseXlsxDifferentialStyle(
      {
        alignment: { attrs: { horizontal: 'center', wrapText: '1' } },
        border: {
          bottom: {
            attrs: { style: 'thin' },
            color: { attrs: { rgb: 'FFFF0000' } },
          },
        },
        fill: {
          patternFill: {
            attrs: { patternType: 'solid' },
            fgColor: { attrs: { theme: '4', tint: '.25' } },
          },
        },
        font: { b: {}, color: { attrs: { rgb: 'FF010203' } } },
        numFmt: { attrs: { formatCode: '0.000', numFmtId: '164' } },
        protection: { attrs: { hidden: 'true', locked: 'false' } },
      },
      '',
      PART,
    );

    expect(result).toEqual({
      alignment: { horizontal: 'center', wrapText: true },
      border: {
        bottom: {
          color: { argb: 'FFFF0000', kind: 'rgb' },
          style: 'thin',
        },
      },
      fill: {
        foregroundColor: { index: 4, kind: 'theme', tint: 0.25 },
        kind: 'pattern',
        pattern: 'solid',
      },
      font: {
        bold: true,
        color: { argb: 'FF010203', kind: 'rgb' },
      },
      numberFormat: '0.000',
      protection: { hidden: true, locked: false },
    });
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.alignment)).toBe(true);
    expect(Object.isFrozen(result.border)).toBe(true);
    expect(Object.isFrozen(result.fill)).toBe(true);
    expect(Object.isFrozen(result.font)).toBe(true);
    expect(Object.isFrozen(result.protection)).toBe(true);
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });

  it('parses prefixed one-item components and omits semantic defaults', () => {
    expect(
      parseXlsxDifferentialStyle(
        {
          's:border': [{}],
          's:fill': [{ 's:patternFill': {} }],
          's:font': [{}],
          's:numFmt': [{ attrs: { formatCode: 'General', numFmtId: '0' } }],
          's:alignment': {},
          's:protection': {},
        },
        's',
        PART,
      ),
    ).toEqual({});
  });

  it('parses non-default prefixed components', () => {
    expect(
      parseXlsxDifferentialStyle(
        {
          's:font': { 's:b': {} },
          's:numFmt': {
            attrs: { formatCode: '0.00', numFmtId: '4294967295' },
          },
        },
        's',
        PART,
      ),
    ).toEqual({ font: { bold: true }, numberFormat: '0.00' });
  });

  it.each([
    [
      { patternFill: { attrs: { patternType: 'gray125' } } },
      { kind: 'pattern', pattern: 'gray125' },
    ],
    [
      {
        patternFill: {
          attrs: { patternType: 'none' },
          fgColor: { attrs: { rgb: 'FFFF0000' } },
        },
      },
      {
        foregroundColor: { argb: 'FFFF0000', kind: 'rgb' },
        kind: 'pattern',
        pattern: 'none',
      },
    ],
    [
      {
        patternFill: {
          attrs: { patternType: 'none' },
          bgColor: { attrs: { indexed: '2' } },
        },
      },
      {
        backgroundColor: { index: 2, kind: 'indexed' },
        kind: 'pattern',
        pattern: 'none',
      },
    ],
  ] as const)(
    'preserves non-default differential fill %#',
    (fill, expected) => {
      expect(parseXlsxDifferentialStyle({ fill }, '', PART).fill).toEqual(
        expected,
      );
    },
  );

  it.each([
    [
      null,
      'invalid-document-structure',
      'Differential style element is invalid',
    ],
    [
      { font: [] },
      'invalid-document-structure',
      'Differential style font element is duplicated',
    ],
    [
      { font: [{}, {}] },
      'invalid-document-structure',
      'Differential style font element is duplicated',
    ],
    [
      { font: ['bad'] },
      'invalid-document-structure',
      'Differential style font element is invalid',
    ],
    [
      { fill: 'bad' },
      'invalid-document-structure',
      'Differential style fill element is invalid',
    ],
    [
      { border: 'bad' },
      'invalid-document-structure',
      'Differential style border element is invalid',
    ],
    [
      { numFmt: 'bad' },
      'invalid-document-structure',
      'Differential style numFmt element is invalid',
    ],
    [
      { numFmt: { attrs: 'bad' } },
      'invalid-document-structure',
      'Differential number format attributes are invalid',
    ],
    [
      { numFmt: {} },
      'invalid-document-value',
      'Differential number-format ID is invalid',
    ],
    [
      { numFmt: { attrs: { formatCode: '0' } } },
      'invalid-document-value',
      'Differential number-format ID is invalid',
    ],
    [
      { numFmt: { attrs: { formatCode: '0', numFmtId: '-1' } } },
      'invalid-document-value',
      'Differential number-format ID is invalid',
    ],
    [
      { numFmt: { attrs: { formatCode: '0', numFmtId: '01' } } },
      'invalid-document-value',
      'Differential number-format ID is invalid',
    ],
    [
      { numFmt: { attrs: { formatCode: '0', numFmtId: '4294967296' } } },
      'invalid-document-value',
      'Differential number-format ID is invalid',
    ],
    [
      { numFmt: { attrs: { numFmtId: '1' } } },
      'invalid-document-value',
      'Differential number-format code is invalid',
    ],
    [
      { numFmt: { attrs: { formatCode: '', numFmtId: '1' } } },
      'invalid-document-value',
      'Differential number-format code is invalid',
    ],
    [
      { font: { color: { attrs: { theme: '12' } } } },
      'invalid-document-value',
      'Font theme-color index is invalid',
    ],
    [
      { alignment: { attrs: { horizontal: 'middle' } } },
      'invalid-document-value',
      'Alignment horizontal value is invalid',
    ],
  ] as const)(
    'rejects invalid differential style %#',
    (source, code, message) => {
      expect(capture(source).diagnostic).toEqual({
        code,
        message,
        part: PART,
        severity: 'error',
      });
    },
  );
});
