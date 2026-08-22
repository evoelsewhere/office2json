import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { parseXlsxContentTypes } from '../../src/formats/xlsx/internal/content-types';
import { parseXlsxRelationships } from '../../src/formats/xlsx/internal/relationships';
import { XlsxResourceLimitError } from '../../src/formats/xlsx/internal/resource-limits';

const CONTENT_TYPES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';

function node(attrs: Record<string, string>): {
  attrs: Record<string, string>;
} {
  return { attrs };
}

function contentTypes(
  defaults: Array<Record<string, string>> = [],
  overrides: Array<Record<string, string>> = [],
): unknown {
  return {
    Types: {
      attrs: { xmlns: CONTENT_TYPES_NAMESPACE },
      Default: defaults.map(node),
      Override: overrides.map(node),
    },
  };
}

function relationships(entries: Array<Record<string, string>>): unknown {
  return {
    Relationships: {
      attrs: { xmlns: RELATIONSHIPS_NAMESPACE },
      Relationship: entries.map(node),
    },
  };
}

function captureParseError(action: () => unknown): XlsxParseError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected XLSX OPC validation to fail');
}

describe('XLSX content types', () => {
  it('resolves canonical overrides before case-insensitive defaults', () => {
    const table = parseXlsxContentTypes(
      contentTypes(
        [
          { Extension: 'XML', ContentType: 'application/xml' },
          { Extension: 'bin', ContentType: 'application/octet-stream' },
        ],
        [
          {
            PartName: '/xl/workbook.xml',
            ContentType:
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
          },
          {
            PartName: '/xl/shared%53trings.xml',
            ContentType:
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
          },
        ],
      ),
    );

    expect(table.contentTypeFor('xl/workbook.xml')).toBe(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    );
    expect(table.contentTypeFor('xl/sharedStrings.xml')).toBe(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
    );
    expect(table.contentTypeFor('custom/item.XML')).toBe('application/xml');
    expect(table.contentTypeFor('custom/blob.bin')).toBe(
      'application/octet-stream',
    );
    expect(table.contentTypeFor('custom/no-extension')).toBeUndefined();
    expect(table.contentTypeFor('custom/.xml')).toBeUndefined();
    expect(table.contentTypeFor('custom/file.')).toBeUndefined();
  });

  it('accepts absent collections and one non-array entry of each kind', () => {
    const empty = parseXlsxContentTypes({
      Types: { attrs: { xmlns: CONTENT_TYPES_NAMESPACE } },
    });
    expect(empty.defaults.size).toBe(0);
    expect(empty.overrides.size).toBe(0);

    const single = parseXlsxContentTypes({
      Types: {
        attrs: { xmlns: CONTENT_TYPES_NAMESPACE },
        Default: node({
          Extension: '09AZaz_-',
          ContentType: 'a/b',
        }),
        Override: node({
          PartName: '/single.item',
          ContentType: 'c/d',
        }),
      },
    });
    expect(single.contentTypeFor('other.09azAZ_-')).toBe('a/b');
    expect(single.contentTypeFor('single.item')).toBe('c/d');
  });

  it.each([
    [{}, 'Content types root is missing or has the wrong namespace'],
    [
      { Types: { attrs: { xmlns: 'urn:wrong' } } },
      'Content types root is missing or has the wrong namespace',
    ],
    [null, 'Content types root is missing or has the wrong namespace'],
    [[], 'Content types root is missing or has the wrong namespace'],
    [
      { Types: { attrs: [] } },
      'Content types root is missing or has the wrong namespace',
    ],
    [
      {
        Types: {
          attrs: { xmlns: CONTENT_TYPES_NAMESPACE },
          Default: 'invalid',
        },
      },
      'Content types contain an invalid entry collection',
    ],
    [
      {
        Types: {
          attrs: { xmlns: CONTENT_TYPES_NAMESPACE },
          Override: [null],
        },
      },
      'Content types contain an invalid entry collection',
    ],
    [
      contentTypes([{ Extension: '', ContentType: 'application/xml' }]),
      'Content type default has an invalid extension',
    ],
    [
      contentTypes([{ Extension: 'xml', ContentType: '' }]),
      'Content type default has an invalid MIME type',
    ],
    [
      contentTypes([], [{ PartName: '', ContentType: 'application/xml' }]),
      'Content type override has an invalid part name',
    ],
    [
      contentTypes(
        [],
        [{ PartName: 1 as unknown as string, ContentType: 'application/xml' }],
      ),
      'Content type override has an invalid part name',
    ],
    [
      contentTypes([], [{ PartName: '/xl/a.xml', ContentType: 'text plain' }]),
      'Content type override has an invalid MIME type',
    ],
  ] as const)('rejects malformed content types %#', (value, message) => {
    const error = captureParseError(() => parseXlsxContentTypes(value));
    expect(error.diagnostic).toEqual({
      code: 'invalid-document-structure',
      message,
      part: '[Content_Types].xml',
      severity: 'error',
    });
    expect(error.cause).toBeUndefined();
  });

  it.each(['.', '/', ':', '@', '[', '`', '{', 'é'])(
    'rejects invalid default extension character %s',
    (extension) => {
      const error = captureParseError(() =>
        parseXlsxContentTypes(
          contentTypes([{ Extension: extension, ContentType: 'a/b' }]),
        ),
      );
      expect(error.diagnostic.message).toBe(
        'Content type default has an invalid extension',
      );
    },
  );

  it.each([
    undefined,
    '',
    'plain',
    '/plain',
    'text/',
    'text plain/value',
    `text/${String.fromCodePoint(0x1f)}value`,
    `text/${String.fromCodePoint(0x7f)}value`,
  ])('rejects invalid MIME type %s', (mimeType) => {
    const error = captureParseError(() =>
      parseXlsxContentTypes(
        contentTypes([
          {
            Extension: 'xml',
            ...(mimeType === undefined ? {} : { ContentType: mimeType }),
          },
        ]),
      ),
    );
    expect(error.diagnostic.message).toBe(
      'Content type default has an invalid MIME type',
    );
  });

  it('rejects duplicate default extensions case-insensitively', () => {
    const error = captureParseError(() =>
      parseXlsxContentTypes(
        contentTypes([
          { Extension: 'XML', ContentType: 'application/xml' },
          { Extension: 'xml', ContentType: 'application/other+xml' },
        ]),
      ),
    );
    expect(error.diagnostic.message).toBe(
      'Content types contain a duplicate default extension',
    );
  });

  it('rejects duplicate canonical override names', () => {
    const error = captureParseError(() =>
      parseXlsxContentTypes(
        contentTypes(
          [],
          [
            {
              PartName: '/xl/sharedStrings.xml',
              ContentType: 'application/xml',
            },
            {
              PartName: '/xl/shared%53trings.xml',
              ContentType: 'application/other+xml',
            },
          ],
        ),
      ),
    );
    expect(error.diagnostic.message).toBe(
      'Content types contain a duplicate canonical part name',
    );
  });

  it('preserves the canonical part validation cause', () => {
    const error = captureParseError(() =>
      parseXlsxContentTypes(
        contentTypes(
          [],
          [{ PartName: '/xl/%GG.xml', ContentType: 'application/xml' }],
        ),
      ),
    );
    expect(error.diagnostic.message).toBe(
      'Content type override has an invalid part name',
    );
    expect(error.cause).toBeInstanceOf(TypeError);
  });
});

describe('XLSX owner-scoped relationships', () => {
  it('resolves package-root relationships and reports their relationship part', () => {
    const table = parseXlsxRelationships(
      relationships([
        {
          Id: 'rIdWorkbook',
          Type: 'urn:type:office-document',
          Target: '/custom/books/workbook.xml',
        },
      ]),
      null,
      1,
    );

    expect(table.get('rIdWorkbook')).toEqual({
      id: 'rIdWorkbook',
      mode: 'internal',
      target: 'custom/books/workbook.xml',
      type: 'urn:type:office-document',
    });

    const error = captureParseError(() =>
      parseXlsxRelationships(
        relationships([
          { Id: 'rIdWorkbook', Type: 'urn:type', Target: '../book.xml' },
        ]),
        null,
        1,
      ),
    );
    expect(error.diagnostic).toMatchObject({
      code: 'invalid-relationship-target',
      part: '_rels/.rels',
      severity: 'error',
    });
  });

  it('resolves internal targets and preserves explicit external targets', () => {
    const table = parseXlsxRelationships(
      relationships([
        {
          Id: 'rIdSheet',
          Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
          Target: 'worksheets/sheet1.xml',
        },
        {
          Id: 'rIdExternal',
          Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
          Target: 'https://example.com/book.xlsx',
          TargetMode: 'External',
        },
      ]),
      'xl/workbook.xml',
      2,
    );

    expect(table.get('rIdSheet')).toEqual({
      id: 'rIdSheet',
      mode: 'internal',
      target: 'xl/worksheets/sheet1.xml',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
    });
    expect(table.get('rIdExternal')).toEqual({
      id: 'rIdExternal',
      mode: 'external',
      target: 'https://example.com/book.xlsx',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
    });
  });

  it('accepts exactly the relationship limit and rejects one over', () => {
    const value = relationships([
      { Id: 'rId1', Type: 'urn:type:one', Target: 'one.xml' },
      { Id: 'rId2', Type: 'urn:type:two', Target: 'two.xml' },
    ]);

    expect(parseXlsxRelationships(value, 'xl/workbook.xml', 2).size).toBe(2);
    expect(() => parseXlsxRelationships(value, 'xl/workbook.xml', 1)).toThrow(
      XlsxResourceLimitError,
    );
    try {
      parseXlsxRelationships(value, 'xl/workbook.xml', 1);
    } catch (error) {
      expect(error).toMatchObject({
        actual: 2,
        limit: 1,
        limitName: 'maxRelationships',
        part: 'xl/workbook.xml',
      });
    }
  });

  it('accepts an absent collection and a single non-array relationship', () => {
    expect(
      parseXlsxRelationships(
        { Relationships: { attrs: { xmlns: RELATIONSHIPS_NAMESPACE } } },
        'xl/workbook.xml',
        1,
      ).size,
    ).toBe(0);
    expect(
      parseXlsxRelationships(
        {
          Relationships: {
            attrs: { xmlns: RELATIONSHIPS_NAMESPACE },
            Relationship: node({
              Id: 'rId1',
              Type: 'urn:type',
              Target: 'sheet.xml',
            }),
          },
        },
        'xl/workbook.xml',
        1,
      ).get('rId1'),
    ).toMatchObject({ target: 'xl/sheet.xml' });
  });

  it.each([
    [
      {},
      'Relationship root is missing or has the wrong namespace',
      'invalid-document-structure',
    ],
    [
      { Relationships: { attrs: { xmlns: 'urn:wrong' } } },
      'Relationship root is missing or has the wrong namespace',
      'invalid-document-structure',
    ],
    [
      null,
      'Relationship root is missing or has the wrong namespace',
      'invalid-document-structure',
    ],
    [
      [],
      'Relationship root is missing or has the wrong namespace',
      'invalid-document-structure',
    ],
    [
      { Relationships: { attrs: [] } },
      'Relationship root is missing or has the wrong namespace',
      'invalid-document-structure',
    ],
    [
      {
        Relationships: {
          attrs: { xmlns: RELATIONSHIPS_NAMESPACE },
          Relationship: 'invalid',
        },
      },
      'Relationships contain an invalid entry collection',
      'invalid-document-structure',
    ],
    [
      relationships([{ Id: '', Type: 'urn:type', Target: 'sheet.xml' }]),
      'Relationship has an invalid ID',
      'invalid-document-structure',
    ],
    [
      relationships([{ Id: ' rId1', Type: 'urn:type', Target: 'sheet.xml' }]),
      'Relationship has an invalid ID',
      'invalid-document-structure',
    ],
    [
      relationships([{ Id: 'rId1', Type: ' urn:type', Target: 'sheet.xml' }]),
      'Relationship has an invalid type',
      'invalid-document-structure',
    ],
    [
      relationships([{ Id: 'rId1', Type: '', Target: 'sheet.xml' }]),
      'Relationship has an invalid type',
      'invalid-document-structure',
    ],
    [
      relationships([{ Id: 'rId1', Type: 'urn:type', Target: '' }]),
      'Relationship has an invalid internal target',
      'invalid-relationship-target',
    ],
    [
      relationships([
        {
          Id: 'rId1',
          Type: 'urn:type',
          Target: 'sheet.xml',
          TargetMode: 'Internal',
        },
      ]),
      'Relationship has an invalid TargetMode',
      'invalid-document-structure',
    ],
    [
      relationships([
        {
          Id: 'rId1',
          Type: 'urn:type',
          Target: ' https://example.com',
          TargetMode: 'External',
        },
      ]),
      'Relationship has an invalid external target',
      'invalid-relationship-target',
    ],
  ] as const)('rejects malformed relationships %#', (value, message, code) => {
    const error = captureParseError(() =>
      parseXlsxRelationships(value, 'xl/workbook.xml', 10),
    );
    expect(error.diagnostic).toEqual({
      code,
      message,
      part: 'xl/workbook.xml',
      severity: 'error',
    });
    if (message === 'Relationship has an invalid internal target') {
      expect(error.cause).toBeInstanceOf(TypeError);
    } else {
      expect(error.cause).toBeUndefined();
    }
  });

  it('rejects a non-string internal target without an underlying path error', () => {
    const error = captureParseError(() =>
      parseXlsxRelationships(
        relationships([
          {
            Id: 'rId1',
            Type: 'urn:type',
            Target: 1 as unknown as string,
          },
        ]),
        'xl/workbook.xml',
        1,
      ),
    );
    expect(error.diagnostic).toEqual({
      code: 'invalid-relationship-target',
      message: 'Relationship has an invalid internal target',
      part: 'xl/workbook.xml',
      severity: 'error',
    });
    expect(error.cause).toBeUndefined();
  });

  it('rejects duplicate relationship IDs', () => {
    const error = captureParseError(() =>
      parseXlsxRelationships(
        relationships([
          { Id: 'rId1', Type: 'urn:type:one', Target: 'one.xml' },
          { Id: 'rId1', Type: 'urn:type:two', Target: 'two.xml' },
        ]),
        'xl/workbook.xml',
        10,
      ),
    );
    expect(error.diagnostic.message).toBe(
      'Relationships contain a duplicate ID',
    );
  });
});
