import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common/xml/tree';
import type { XlsxRichValueField } from '../../src/formats/xlsx';
import {
  cloneXlsxRichValueForOutput,
  parseXlsxRichValueParts,
} from '../../src/formats/xlsx/internal/rich-value';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';

const NS = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';
const limits = defaultXlsxResourceLimits();

function tree(value: object): XmlLookupValue {
  return value as XmlLookupValue;
}

interface KeyDefinition {
  name: string;
  type?: XlsxRichValueField['type'];
}

function structures(
  keys: readonly KeyDefinition[],
  overrides: Record<string, unknown> = {},
): XmlLookupValue {
  return tree({
    rvStructures: {
      attrs: { count: '1', xmlns: NS },
      s: {
        attrs: { t: 'entity' },
        k: keys.map((key) => ({
          attrs: {
            n: key.name,
            ...(key.type === undefined ? {} : { t: key.type }),
          },
        })),
      },
      ...overrides,
    },
  });
}

function data(
  values: readonly unknown[],
  fallback?: unknown,
  overrides: Record<string, unknown> = {},
): XmlLookupValue {
  return tree({
    rvData: {
      attrs: { count: '1', xmlns: NS },
      rv: {
        attrs: { s: '0' },
        ...(fallback === undefined ? {} : { fb: fallback }),
        v: [...values],
      },
      ...overrides,
    },
  });
}

function parse(
  keyDefinitions: readonly KeyDefinition[],
  values: readonly unknown[],
  fallback?: unknown,
) {
  return parseXlsxRichValueParts(
    structures(keyDefinitions),
    data(values, fallback),
    'xl/richData/structure.xml',
    'xl/richData/data.xml',
    limits,
    { records: 0 },
  );
}

describe('XLSX normalized rich values', () => {
  it('parses prefixed normalized roots, children, and scalar nodes', () => {
    const registry = parseXlsxRichValueParts(
      tree({
        'ns_rd:rvStructures': {
          attrs: { count: '1', 'xmlns:rd': NS },
          'ns_rd:s': {
            attrs: { t: 'entity', 'xmlns:rd': NS },
            'ns_rd:k': { attrs: { n: 'Text', 'xmlns:rd': NS } },
          },
        },
      }),
      tree({
        'ns_rd:rvData': {
          attrs: { count: '1', 'xmlns:rd': NS },
          'ns_rd:rv': {
            attrs: { s: '0', 'xmlns:rd': NS },
            'ns_rd:v': {
              attrs: { 'xmlns:rd': NS },
              value: 'hello',
            },
          },
        },
      }),
      'structure.xml',
      'data.xml',
      limits,
      { records: 0 },
    );
    expect(registry).toStrictEqual({
      part: 'data.xml',
      values: [
        {
          fields: [
            {
              name: 'Text',
              type: 's',
              value: { kind: 'text', value: 'hello' },
            },
          ],
          sourceDataOmitted: false,
          type: 'entity',
        },
      ],
    });
  });

  it.each([
    ['0', false],
    ['false', false],
    ['1', true],
    ['true', true],
  ] as const)('parses rich boolean %s', (raw, expected) => {
    expect(parse([{ name: 'Flag', type: 'b' }], [raw]).values[0]).toMatchObject(
      {
        fields: [{ value: { kind: 'boolean', value: expected } }],
      },
    );
  });

  it.each([
    ['+1', 1],
    ['-1.5', -1.5],
    ['.5', 0.5],
    ['.55', 0.55],
    ['1.', 1],
    ['1e3', 1000],
    ['1e30', 1e30],
    ['-2.5E-2', -0.025],
    ['-0', 0],
  ] as const)('parses rich finite number %s', (raw, expected) => {
    const value = parse([{ name: 'Value', type: 'd' }], [raw]).values[0];
    expect(value?.fields[0]?.value).toStrictEqual({
      kind: 'number',
      value: expected,
    });
    if (raw === '-0') expect(Object.is(expected, -0)).toBe(false);
  });

  it.each([' 1', '1 ', '1e', '.', 'x1', '1x', '1e999'])(
    'rejects rich finite number %s',
    (raw) => {
      expect(() => parse([{ name: 'Value', type: 'd' }], [raw])).toThrow(
        'Rich-value number is invalid',
      );
    },
  );

  it.each(['-1', '01', '1.0', ' 1', '1 ', '4294967296'])(
    'rejects rich UInt32 %s',
    (raw) => {
      expect(() => parse([{ name: 'Count', type: 'i' }], [raw])).toThrow(
        'Rich-value integer is invalid',
      );
    },
  );

  it('accepts the complete UInt32 range and every index kind', () => {
    const registry = parse(
      [
        { name: 'Count', type: 'i' },
        { name: 'Array', type: 'a' },
        { name: 'Bag', type: 'spb' },
      ],
      ['4294967295', '0', '4294967295'],
    );
    expect(registry.values[0]?.fields).toStrictEqual([
      {
        name: 'Count',
        type: 'i',
        value: { kind: 'integer', value: 4294967295 },
      },
      { name: 'Array', type: 'a', value: { kind: 'array-index', value: 0 } },
      { name: 'Bag', type: 'spb', value: { kind: 'omitted' } },
    ]);
  });

  it.each([
    ['a', 'bad', 'Rich-value array index is invalid'],
    ['r', 'bad', 'Rich-value index is invalid'],
    ['spb', 'bad', 'Rich-value property-bag index is invalid'],
  ] as const)('rejects invalid %s scalar', (type, raw, message) => {
    expect(() => parse([{ name: 'Value', type }], [raw])).toThrow(message);
  });

  it('parses every fallback type and omitted fallback', () => {
    expect(parse([{ name: 'Text' }], ['x']).values[0]).not.toHaveProperty(
      'fallback',
    );
    expect(
      parse([{ name: 'Text' }], ['x'], 'plain').values[0]?.fallback,
    ).toStrictEqual({
      kind: 'text',
      value: 'plain',
    });
    expect(
      parse([{ name: 'Text' }], ['x'], { attrs: {}, value: 'plain' }).values[0]
        ?.fallback,
    ).toStrictEqual({ kind: 'text', value: 'plain' });
    expect(
      parse([{ name: 'Text' }], ['x'], { attrs: { t: 'b' }, value: 'false' })
        .values[0]?.fallback,
    ).toStrictEqual({ kind: 'boolean', value: false });
    expect(
      parse([{ name: 'Text' }], ['x'], { attrs: { t: 'n' }, value: '.5' })
        .values[0]?.fallback,
    ).toStrictEqual({ kind: 'number', value: 0.5 });
    expect(
      parse([{ name: 'Text' }], ['x'], { attrs: { t: 'e' }, value: '#N/A' })
        .values[0]?.fallback,
    ).toStrictEqual({ code: '#N/A', kind: 'error' });
  });

  it.each([
    [{ attrs: { t: 'b' }, value: 'bad' }, 'Rich-value fallback is invalid'],
    [{ attrs: { t: 'n' }, value: 'bad' }, 'Rich-value fallback is invalid'],
    [
      { attrs: { t: 'e' }, value: 'bad' },
      'Rich-value fallback type is invalid',
    ],
    [
      { attrs: { t: 'x' }, value: '#N/A' },
      'Rich-value fallback type is invalid',
    ],
    [
      { attrs: { t: 'e' }, value: 'x#N/A' },
      'Rich-value fallback type is invalid',
    ],
    [
      { attrs: { t: 'e' }, value: '#N/Ax' },
      'Rich-value fallback type is invalid',
    ],
    [{ child: 'nested' }, 'Rich-value fallback is invalid'],
  ] as const)('rejects invalid fallback %#', (fallback, message) => {
    expect(() => parse([{ name: 'Text' }], ['x'], fallback)).toThrow(message);
  });

  it('redacts every sensitive key class while preserving explicit display keys', () => {
    const names = [
      '_DisplayString',
      'Text',
      'Name',
      '%EntityId',
      'streetAddress',
      'crid',
      'identifier',
      'license',
      'provider',
      'dataSource',
      'imageUrl',
      'entityId',
      'recordId',
      'serviceId',
      '_id',
      'id',
      'abc%Entity',
      'idSuffix',
    ];
    const value = parse(
      names.map((name) => ({ name })),
      names.map(() => 'secret'),
    ).values[0]!;
    expect(
      value.fields.slice(0, 3).map((field) => field.value.kind),
    ).toStrictEqual(['text', 'text', 'text']);
    expect(
      value.fields
        .slice(3, 16)
        .every((field) => field.value.kind === 'omitted'),
    ).toBe(true);
    expect(
      value.fields.slice(16).map((field) => field.value.kind),
    ).toStrictEqual(['text', 'text']);
    expect(value.sourceDataOmitted).toBe(true);
  });

  it.each([
    [
      tree({
        unrelated: {},
        rvStructures: {
          attrs: { count: '1', xmlns: NS },
          s: {
            attrs: { t: 'entity' },
            k: { attrs: { n: 'Text' } },
          },
        },
      }),
      data(['x']),
      undefined,
    ],
    [
      tree({}),
      data(['x']),
      'Rich-value rvStructures root is missing or duplicated',
    ],
    [
      tree({
        rvStructures: { attrs: { xmlns: NS } },
        'x:rvStructures': { attrs: { 'xmlns:x': NS } },
      }),
      data(['x']),
      'Rich-value rvStructures root is missing or duplicated',
    ],
    [
      tree({ rvStructures: 'text' }),
      data(['x']),
      'Rich-value rvStructures root has the wrong namespace',
    ],
    [
      tree({ rvStructures: { attrs: { count: '0', xmlns: NS }, s: 'text' } }),
      data(['x']),
      'Rich-value collection is invalid',
    ],
    [
      tree({
        rvStructures: {
          attrs: { count: '1', xmlns: NS },
          'x:s': { attrs: { t: 'entity', 'xmlns:x': 'urn:wrong' } },
        },
      }),
      data(['x']),
      'Rich-value element has the wrong namespace',
    ],
    [
      tree({
        rvStructures: {
          attrs: { count: '1', xmlns: NS },
          s: { attrs: { t: 'entity' } },
        },
      }),
      data(['x']),
      'Rich-value structure has no keys',
    ],
    [
      structures([{ name: 'Text' }]),
      tree({
        rvData: {
          attrs: { count: '1', xmlns: NS },
          rv: {
            attrs: { s: '0' },
            'x:v': { attrs: { 'xmlns:x': 'urn:wrong' }, value: 'x' },
          },
        },
      }),
      'Rich-value element has the wrong namespace',
    ],
    [
      structures([{ name: 'Text' }]),
      tree({
        rvData: {
          attrs: { count: '1', xmlns: NS },
          rv: { attrs: { s: '0' }, v: { child: 'nested' } },
        },
      }),
      'Rich-value field is invalid',
    ],
    [
      structures([{ name: 'Text' }]),
      data(['x'], undefined, { attrs: { count: '2', xmlns: NS } }),
      'Rich-value data count does not match',
    ],
    [
      structures([{ name: 'Text' }]),
      tree({
        rvData: {
          attrs: { count: '1', xmlns: NS },
          rv: { attrs: { s: '0' }, fb: ['a', 'b'], v: ['x'] },
        },
      }),
      'Rich-value fallback is duplicated',
    ],
  ] as const)(
    'rejects invalid normalized rich structure %#',
    (structure, values, message) => {
      const action = () =>
        parseXlsxRichValueParts(
          structure,
          values,
          'structure.xml',
          'data.xml',
          limits,
          { records: 0 },
        );
      if (message === undefined) expect(action()).toBeDefined();
      else expect(action).toThrow(message);
    },
  );

  it('accepts explicit empty normalized collections and empty scalar text', () => {
    expect(
      parseXlsxRichValueParts(
        tree({
          rvStructures: { attrs: { count: '0', xmlns: NS }, s: undefined },
        }),
        tree({ rvData: { attrs: { count: '0', xmlns: NS }, rv: undefined } }),
        'structure.xml',
        'data.xml',
        limits,
        { records: 0 },
      ),
    ).toStrictEqual({ part: 'data.xml', values: [] });
    expect(
      parse([{ name: 'Text' }], [{ attrs: {} }]).values[0]?.fields[0]?.value,
    ).toStrictEqual({ kind: 'text', value: '' });
  });

  it.each(['x#N/A', '#N/Ax'])('rejects unanchored rich errors %s', (raw) => {
    expect(() => parse([{ name: 'Failure', type: 'e' }], [raw])).toThrow(
      'Rich-value error is invalid',
    );
  });

  it('rejects missing counts and a lexical structure reference', () => {
    let missingCountError: unknown;
    try {
      parseXlsxRichValueParts(
        tree({
          rvStructures: {
            attrs: { xmlns: NS },
            s: { attrs: { t: 'entity' }, k: { attrs: { n: 'Text' } } },
          },
        }),
        data(['x']),
        'structure.xml',
        'data.xml',
        limits,
        { records: 0 },
      );
    } catch (error) {
      missingCountError = error;
    }
    expect(missingCountError).toMatchObject({
      diagnostic: {
        code: 'invalid-document-structure',
        message: 'Rich-value structure count does not match',
      },
    });
    expect(() =>
      parseXlsxRichValueParts(
        structures([{ name: 'Text' }]),
        tree({
          rvData: {
            attrs: { count: '1', xmlns: NS },
            rv: { attrs: { s: 'bad' }, v: ['x'] },
          },
        }),
        'structure.xml',
        'data.xml',
        limits,
        { records: 0 },
      ),
    ).toThrow('Rich-value structure reference is invalid');
  });

  it('rejects a rich-value reference exactly at the data count boundary', () => {
    expect(() => parse([{ name: 'Related', type: 'r' }], ['1'])).toThrow(
      'Rich-value reference is invalid',
    );
  });

  it('enforces unsafe aggregate record and text accounting', () => {
    expect(() =>
      parseXlsxRichValueParts(
        structures([{ name: 'Text' }]),
        data(['x']),
        'structure.xml',
        'data.xml',
        { ...limits, maxMetadataRecords: Number.MAX_SAFE_INTEGER },
        { records: Number.MAX_SAFE_INTEGER },
      ),
    ).toThrow('maxMetadataRecords');
    expect(() =>
      cloneXlsxRichValueForOutput(
        {
          fields: [],
          sourceDataOmitted: false,
          type: 'x',
        },
        { records: 0 },
        { textCharacters: Number.MAX_SAFE_INTEGER },
        { ...limits, maxTextCharacters: Number.MAX_SAFE_INTEGER },
        'sheet.xml',
      ),
    ).toThrow('maxTextCharacters');
  });
});
