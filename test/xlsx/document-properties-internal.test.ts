import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common/xml/tree';
import {
  parseXlsxApplicationDocumentProperties,
  parseXlsxCoreDocumentProperties,
  parseXlsxCustomDocumentProperties,
  parseXlsxDocumentBoolean,
  parseXlsxDocumentUnsignedInteger,
} from '../../src/formats/xlsx/internal/document-properties';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import { createXlsxWorksheetBudget } from '../../src/formats/xlsx/internal/worksheet';

const CORE_NS =
  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const DC_NS = 'http://purl.org/dc/elements/1.1/';
const APP_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const CUSTOM_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties';
const VT_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
const FORMAT_ID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';

function tree(value: object): XmlLookupValue {
  return value as XmlLookupValue;
}

function budget() {
  return createXlsxWorksheetBudget({ part: null, values: [] });
}

const limits = defaultXlsxResourceLimits();

describe('XLSX normalized document properties', () => {
  it.each([
    ['0', 0],
    ['1', 1],
    ['4294967295', 4294967295],
  ] as const)('parses document-property UInt32 %s', (value, expected) => {
    expect(
      parseXlsxDocumentUnsignedInteger(value, 'bad uint', 'part.xml'),
    ).toBe(expected);
  });

  it.each(['', '-1', '01', '1.0', '1x', 'x1', '4294967296'])(
    'rejects document-property UInt32 %s',
    (value) => {
      expect(() =>
        parseXlsxDocumentUnsignedInteger(value, 'bad uint', 'part.xml'),
      ).toThrow('bad uint');
    },
  );

  it('rejects a non-string document-property UInt32', () => {
    expect(() =>
      parseXlsxDocumentUnsignedInteger(undefined, 'bad uint', 'part.xml'),
    ).toThrow('bad uint');
  });

  it.each([
    ['0', false],
    ['false', false],
    ['1', true],
    ['true', true],
  ] as const)('parses document-property boolean %s', (value, expected) => {
    expect(parseXlsxDocumentBoolean(value, 'bad bool', 'part.xml')).toBe(
      expected,
    );
  });

  it('rejects an invalid document-property boolean with the supplied context', () => {
    expect(() =>
      parseXlsxDocumentBoolean('False', 'bad bool', 'part.xml'),
    ).toThrow('bad bool');
  });

  it('parses default and normalized reserved-prefix roots', () => {
    expect(
      parseXlsxCoreDocumentProperties(
        tree({
          coreProperties: {
            attrs: { xmlns: CORE_NS },
            category: 'default',
          },
        }),
        budget(),
        limits,
        'core.xml',
      ),
    ).toStrictEqual({ category: 'default' });
    expect(
      parseXlsxCoreDocumentProperties(
        tree({
          'ns_cp:coreProperties': {
            attrs: { 'xmlns:cp': CORE_NS, 'xmlns:dc': DC_NS },
            'dc:title': 'normalized',
          },
        }),
        budget(),
        limits,
        'core.xml',
      ),
    ).toStrictEqual({ title: 'normalized' });
  });

  it.each([
    [tree({ 'cp:coreProperties': 'text' })],
    [
      tree({
        'cp:coreProperties': { attrs: { 'xmlns:cp': CORE_NS } },
        'x:coreProperties': { attrs: { 'xmlns:x': CORE_NS } },
      }),
    ],
  ] as const)('rejects invalid normalized core root %#', (value) => {
    expect(() =>
      parseXlsxCoreDocumentProperties(value, budget(), limits, 'core.xml'),
    ).toThrow(/properties root is missing or duplicated|root is invalid/u);
  });

  it('ignores unrelated normalized entries and rejects an unbound root prefix', () => {
    expect(
      parseXlsxCoreDocumentProperties(
        tree({
          ignored: {},
          coreProperties: {
            attrs: { xmlns: CORE_NS },
            category: 'kept',
          },
        }),
        budget(),
        limits,
        'core.xml',
      ),
    ).toStrictEqual({ category: 'kept' });
    expect(() =>
      parseXlsxCoreDocumentProperties(
        tree({
          'x:coreProperties': { attrs: { 'xmlns:cp': CORE_NS } },
        }),
        budget(),
        limits,
        'core.xml',
      ),
    ).toThrow('Document properties root has the wrong namespace');
  });

  it('rejects non-string normalized scalar values and nested values', () => {
    for (const value of [{ attrs: {}, value: 1 }, { nested: {} }]) {
      expect(() =>
        parseXlsxCoreDocumentProperties(
          tree({
            'cp:coreProperties': {
              attrs: { 'xmlns:cp': CORE_NS, 'xmlns:dc': DC_NS },
              'dc:title': value,
            },
          }),
          budget(),
          limits,
          'core.xml',
        ),
      ).toThrow('Core document property is invalid');
    }
  });

  it('accepts an explicit undefined normalized leaf as empty text', () => {
    expect(
      parseXlsxCoreDocumentProperties(
        tree({
          'cp:coreProperties': {
            attrs: { 'xmlns:cp': CORE_NS, 'xmlns:dc': DC_NS },
            'dc:title': undefined,
          },
        }),
        budget(),
        limits,
        'core.xml',
      ),
    ).toStrictEqual({ title: '' });
  });

  it('parses normalized application and custom empty leaves', () => {
    expect(
      parseXlsxApplicationDocumentProperties(
        tree({
          Properties: {
            Application: undefined,
            attrs: { xmlns: APP_NS },
          },
        }),
        budget(),
        limits,
        'app.xml',
      ),
    ).toStrictEqual({ application: '' });
    expect(
      parseXlsxCustomDocumentProperties(
        tree({
          Properties: {
            attrs: { 'xmlns:vt': VT_NS, xmlns: CUSTOM_NS },
            property: {
              attrs: { fmtid: FORMAT_ID, name: 'EmptyString', pid: '2' },
              'vt:lpstr': undefined,
            },
          },
        }),
        budget(),
        limits,
        'custom.xml',
      )[0]?.value,
    ).toStrictEqual({ kind: 'string', value: '' });
  });

  it('rejects a normalized non-object custom property', () => {
    expect(() =>
      parseXlsxCustomDocumentProperties(
        tree({
          Properties: {
            attrs: { 'xmlns:vt': VT_NS, xmlns: CUSTOM_NS },
            property: 'not-an-object',
          },
        }),
        budget(),
        limits,
        'custom.xml',
      ),
    ).toThrow('Custom document property is invalid');
  });
});
