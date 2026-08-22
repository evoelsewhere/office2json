import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { parseXlsxStyleFont } from '../../src/formats/xlsx/internal/style-font';

const PART = 'xl/styles.xml';

function element(val?: unknown): unknown {
  return val === undefined ? {} : { attrs: { val } };
}

function capture(value: unknown, prefix = ''): XlsxParseError {
  try {
    parseXlsxStyleFont(value, prefix, PART);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected style font parsing to fail');
}

describe('XLSX style fonts', () => {
  it('normalizes every supported font property into frozen portable JSON', () => {
    const result = parseXlsxStyleFont(
      {
        b: element(),
        charset: element('255'),
        color: { attrs: { theme: '4', tint: '.25' } },
        condense: element('true'),
        extend: element('1'),
        family: element('5'),
        i: element('false'),
        name: element('Aptos Display'),
        outline: element('1'),
        scheme: element('major'),
        shadow: element('true'),
        strike: element('1'),
        sz: element('11.5'),
        u: element('doubleAccounting'),
        vertAlign: element('superscript'),
      },
      '',
      PART,
    );

    expect(result).toEqual({
      bold: true,
      charset: 255,
      color: { index: 4, kind: 'theme', tint: 0.25 },
      condense: true,
      extend: true,
      family: 5,
      name: 'Aptos Display',
      outline: true,
      scheme: 'major',
      shadow: true,
      size: 11.5,
      strike: true,
      underline: 'double-accounting',
      verticalAlignment: 'superscript',
    });
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.color)).toBe(true);
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });

  it('parses prefixed font children and semantic defaults', () => {
    expect(
      parseXlsxStyleFont(
        {
          's:b': element('0'),
          's:charset': element('0'),
          's:color': {},
          's:family': element('0'),
          's:i': element('true'),
          's:scheme': element('none'),
          's:sz': element('.5'),
          's:u': element('none'),
          's:vertAlign': element('baseline'),
        },
        's',
        PART,
      ),
    ).toEqual({ charset: 0, family: 0, italic: true, size: 0.5 });
  });

  it('accepts a normalized parser array containing exactly one font child', () => {
    expect(parseXlsxStyleFont({ b: [element()] }, '', PART)).toEqual({
      bold: true,
    });
  });

  it.each([
    ['11', 11],
    ['1.', 1],
    ['1.25', 1.25],
    ['.25', 0.25],
    ['409', 409],
  ] as const)('accepts exact font-size boundary %s', (source, expected) => {
    expect(parseXlsxStyleFont({ sz: element(source) }, '', PART).size).toBe(
      expected,
    );
  });

  it.each([
    ['b', 'bold', undefined, true],
    ['b', 'bold', '1', true],
    ['b', 'bold', 'true', true],
    ['b', 'bold', '0', false],
    ['b', 'bold', 'false', false],
    ['i', 'italic', undefined, true],
    ['strike', 'strike', undefined, true],
    ['outline', 'outline', undefined, true],
    ['shadow', 'shadow', undefined, true],
    ['condense', 'condense', undefined, true],
    ['extend', 'extend', undefined, true],
  ] as const)(
    'normalizes boolean %s=%s',
    (child, property, source, expected) => {
      const result = parseXlsxStyleFont({ [child]: element(source) }, '', PART);
      expect(property in result).toBe(expected);
    },
  );

  it.each([
    [undefined, 'single'],
    ['single', 'single'],
    ['double', 'double'],
    ['singleAccounting', 'single-accounting'],
    ['doubleAccounting', 'double-accounting'],
  ] as const)('normalizes underline %s', (source, expected) => {
    expect(parseXlsxStyleFont({ u: element(source) }, '', PART).underline).toBe(
      expected,
    );
  });

  it.each([
    ['major', 'major'],
    ['minor', 'minor'],
  ] as const)('preserves font scheme %s', (source, expected) => {
    expect(
      parseXlsxStyleFont({ scheme: element(source) }, '', PART).scheme,
    ).toBe(expected);
  });

  it.each([
    ['superscript', 'superscript'],
    ['subscript', 'subscript'],
  ] as const)('preserves vertical alignment %s', (source, expected) => {
    expect(
      parseXlsxStyleFont({ vertAlign: element(source) }, '', PART)
        .verticalAlignment,
    ).toBe(expected);
  });

  it.each([
    [null, 'invalid-document-structure', 'Font element is invalid'],
    [{ b: [] }, 'invalid-document-structure', 'Font b element is duplicated'],
    [
      { b: [element(), element()] },
      'invalid-document-structure',
      'Font b element is duplicated',
    ],
    [{ b: ['bad'] }, 'invalid-document-structure', 'Font b element is invalid'],
    [{ b: 'bad' }, 'invalid-document-structure', 'Font b element is invalid'],
    [
      { b: { attrs: 'bad' } },
      'invalid-document-structure',
      'Font b attributes are invalid',
    ],
    [
      { b: element('yes') },
      'invalid-document-value',
      'Font b value is invalid',
    ],
    [
      { name: element() },
      'invalid-document-value',
      'Font name value is invalid',
    ],
    [
      { name: element('') },
      'invalid-document-value',
      'Font name value is invalid',
    ],
    [
      { name: element(1) },
      'invalid-document-value',
      'Font name value is invalid',
    ],
    [
      { sz: element('0') },
      'invalid-document-value',
      'Font sz value is invalid',
    ],
    [
      { sz: element('410') },
      'invalid-document-value',
      'Font sz value is invalid',
    ],
    [
      { sz: element('1x') },
      'invalid-document-value',
      'Font sz value is invalid',
    ],
    [
      { sz: element(' 1') },
      'invalid-document-value',
      'Font sz value is invalid',
    ],
    [
      { sz: element('1 ') },
      'invalid-document-value',
      'Font sz value is invalid',
    ],
    [{ sz: element(1) }, 'invalid-document-value', 'Font sz value is invalid'],
    [
      { family: element('-1') },
      'invalid-document-value',
      'Font family value is invalid',
    ],
    [
      { family: element('01') },
      'invalid-document-value',
      'Font family value is invalid',
    ],
    [
      { family: element('6') },
      'invalid-document-value',
      'Font family value is invalid',
    ],
    [
      { charset: element('256') },
      'invalid-document-value',
      'Font charset value is invalid',
    ],
    [
      { scheme: element() },
      'invalid-document-value',
      'Font scheme value is invalid',
    ],
    [
      { scheme: element('Major') },
      'invalid-document-value',
      'Font scheme value is invalid',
    ],
    [
      { u: element('triple') },
      'invalid-document-value',
      'Font u value is invalid',
    ],
    [
      { u: { attrs: 'bad' } },
      'invalid-document-structure',
      'Font u attributes are invalid',
    ],
    [
      { vertAlign: element() },
      'invalid-document-value',
      'Font vertAlign value is invalid',
    ],
    [
      { vertAlign: element('top') },
      'invalid-document-value',
      'Font vertAlign value is invalid',
    ],
    [
      { color: { attrs: { theme: '12' } } },
      'invalid-document-value',
      'Font theme-color index is invalid',
    ],
  ] as const)('rejects invalid font %#', (source, code, message) => {
    expect(capture(source).diagnostic).toEqual({
      code,
      message,
      part: PART,
      severity: 'error',
    });
  });
});
