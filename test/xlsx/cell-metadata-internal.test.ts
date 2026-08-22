import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common/xml/tree';
import {
  parseXlsxCellMetadataPart,
  parseXlsxCellMetadataUnsignedInteger,
  resolveXlsxCellMetadata,
  type XlsxCellMetadataBudget,
  type XlsxCellMetadataRegistry,
} from '../../src/formats/xlsx/internal/cell-metadata';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

function tree(value: object): XmlLookupValue {
  return value as XmlLookupValue;
}

function budget(): XlsxCellMetadataBudget {
  return { records: 0 };
}

const limits = defaultXlsxResourceLimits();

describe('XLSX normalized cell metadata', () => {
  it.each([
    ['0', 0],
    ['1', 1],
    ['4294967295', 4294967295],
  ] as const)('parses metadata UInt32 %s', (value, expected) => {
    expect(
      parseXlsxCellMetadataUnsignedInteger(value, 'bad uint', 'metadata.xml'),
    ).toBe(expected);
  });

  it.each([
    undefined,
    '',
    '-1',
    '01',
    'x1',
    '1x',
    '4294967296',
    '9007199254740992',
  ])('rejects metadata UInt32 %s', (value) => {
    expect(() =>
      parseXlsxCellMetadataUnsignedInteger(value, 'bad uint', 'metadata.xml'),
    ).toThrow('bad uint');
  });

  it('parses default and normalized reserved-prefix roots while ignoring unrelated entries', () => {
    expect(
      parseXlsxCellMetadataPart(
        tree({ ignored: {}, metadata: { attrs: { xmlns: NS } } }),
        'transitional',
        'metadata.xml',
        limits,
        budget(),
      ),
    ).toStrictEqual({ cellBlocks: [], part: 'metadata.xml', valueBlocks: [] });
    expect(
      parseXlsxCellMetadataPart(
        tree({
          'ns_a:metadata': { attrs: { 'xmlns:a': NS } },
        }),
        'transitional',
        'metadata.xml',
        limits,
        budget(),
      ),
    ).toStrictEqual({ cellBlocks: [], part: 'metadata.xml', valueBlocks: [] });
    expect(
      parseXlsxCellMetadataPart(
        tree({ metadata: { attrs: { xmlns: NS }, metadataTypes: undefined } }),
        'transitional',
        'metadata.xml',
        limits,
        budget(),
      ),
    ).toStrictEqual({ cellBlocks: [], part: 'metadata.xml', valueBlocks: [] });
  });

  it.each([
    [tree({ metadata: 'text' }), 'Cell metadata root has the wrong namespace'],
    [
      tree({
        metadata: { attrs: { xmlns: NS } },
        'x:metadata': { attrs: { 'xmlns:x': NS } },
      }),
      'Cell metadata root is missing or duplicated',
    ],
    [
      tree({
        metadata: {
          attrs: { xmlns: NS },
          metadataTypes: 'text',
        },
      }),
      'Cell metadata collection is invalid',
    ],
    [
      tree({
        metadata: {
          'x:metadataTypes': { attrs: { 'xmlns:x': 'urn:wrong' } },
          attrs: { xmlns: NS },
        },
      }),
      'Cell metadata element has the wrong namespace',
    ],
    [
      tree({
        metadata: {
          attrs: { xmlns: NS },
          metadataTypes: [{}, {}],
        },
      }),
      'Cell metadata metadataTypes is duplicated',
    ],
  ] as const)('rejects invalid normalized structure %#', (value, message) => {
    expect(() =>
      parseXlsxCellMetadataPart(
        value,
        'transitional',
        'metadata.xml',
        limits,
        budget(),
      ),
    ).toThrow(message);
  });

  it('resolves cell and value registries independently and copies entries', () => {
    const entry = {
      collapsed: false,
      dynamic: true,
      kind: 'dynamic-array' as const,
    };
    const registry: XlsxCellMetadataRegistry = {
      cellBlocks: [[entry]],
      part: 'metadata.xml',
      valueBlocks: [[{ kind: 'rich-value', valueIndex: 4 }]],
    };
    const state = budget();
    expect(
      resolveXlsxCellMetadata(
        registry,
        'cell',
        undefined,
        state,
        { textCharacters: 0 },
        limits,
        'sheet.xml',
        'A1',
      ),
    ).toBeUndefined();
    const resolved = resolveXlsxCellMetadata(
      registry,
      'cell',
      1,
      state,
      { textCharacters: 0 },
      limits,
      'sheet.xml',
      'A1',
    );
    expect(resolved).toStrictEqual({ entries: [entry], unsupported: false });
    expect(resolved?.entries[0]).not.toBe(entry);
    expect(
      resolveXlsxCellMetadata(
        registry,
        'value',
        1,
        state,
        { textCharacters: 0 },
        limits,
        'sheet.xml',
        'A1',
      )?.entries,
    ).toStrictEqual([{ kind: 'rich-value', valueIndex: 4 }]);
  });

  it('rejects unsafe-integer aggregate metadata accounting', () => {
    const registry: XlsxCellMetadataRegistry = {
      cellBlocks: [[{ kind: 'rich-value', valueIndex: 0 }]],
      part: 'metadata.xml',
      valueBlocks: [],
    };
    expect(() =>
      resolveXlsxCellMetadata(
        registry,
        'cell',
        1,
        { records: Number.MAX_SAFE_INTEGER },
        { textCharacters: 0 },
        { ...limits, maxMetadataRecords: Number.MAX_SAFE_INTEGER },
        'sheet.xml',
        'A1',
      ),
    ).toThrow('maxMetadataRecords');
  });
});
