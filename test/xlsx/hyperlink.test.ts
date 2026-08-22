import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { parseXlsxHyperlink } from '../../src/formats/xlsx/internal/hyperlink';
import type { XlsxRelationship } from '../../src/formats/xlsx/internal/relationships';
import type { XlsxXmlElement } from '../../src/formats/xlsx/internal/streaming-xml';

const PART = 'xl/worksheets/sheet1.xml';
const REL_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STRICT_REL_NS = 'http://purl.oclc.org/ooxml/officeDocument/relationships';

function element(
  attributes: Record<string, string>,
  id?: string,
  strict = false,
): XlsxXmlElement {
  return {
    attributes: new Map([
      ...Object.entries(attributes).map(
        ([name, value]) => [`{}${name}`, value] as const,
      ),
      ...(id === undefined
        ? []
        : [[`{${strict ? STRICT_REL_NS : REL_NS}}id`, id] as const]),
    ]),
    localName: 'hyperlink',
    namespace: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  };
}

function relationship(
  overrides: Partial<XlsxRelationship> = {},
  strict = false,
): XlsxRelationship {
  return {
    id: 'rId1',
    mode: 'external',
    target: 'https://example.com/',
    type: `${strict ? STRICT_REL_NS : REL_NS}/hyperlink`,
    ...overrides,
  };
}

function capture(action: () => unknown): XlsxParseError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected hyperlink parsing to fail');
}

describe('XLSX hyperlinks', () => {
  it('normalizes internal hyperlink locations and presentation text', () => {
    const result = parseXlsxHyperlink(
      element({
        display: 'Open',
        location: "'Other sheet'!A1",
        ref: 'a1:b2',
        tooltip: 'Jump',
      }),
      'transitional',
      new Map(),
      PART,
    );

    expect(result).toEqual({
      display: 'Open',
      range: {
        end: { column: 2, row: 2 },
        reference: 'A1:B2',
        start: { column: 1, row: 1 },
      },
      target: { kind: 'internal', location: "'Other sheet'!A1" },
      textCharacters: 24,
      tooltip: 'Jump',
    });
  });

  it('allowlists, decodes, canonicalizes, and redacts external URLs', () => {
    const relationships = new Map([
      [
        'rId1',
        relationship({
          target: 'https://user:password@example.com/path?first=1&amp;second=2',
        }),
      ],
    ]);
    const result = parseXlsxHyperlink(
      element({ location: 'Fragment', ref: 'C3' }, 'rId1'),
      'transitional',
      relationships,
      PART,
    );

    expect(result.target).toEqual({
      kind: 'external',
      location: 'Fragment',
      url: 'https://example.com/path?first=1&second=2',
    });
    expect(result.textCharacters).toBe(
      'https://example.com/path?first=1&second=2'.length + 'Fragment'.length,
    );
    expect(JSON.stringify(result)).not.toContain('password');
    expect(JSON.stringify(result)).not.toContain('rId1');
  });

  it.each([
    ['http://example.com/a', 'http://example.com/a'],
    ['https://example.com/a', 'https://example.com/a'],
    ['mailto:agent@example.com', 'mailto:agent@example.com'],
  ])('allows external target %s', (target, expected) => {
    const result = parseXlsxHyperlink(
      element({ ref: 'A1' }, 'rId1'),
      'transitional',
      new Map([['rId1', relationship({ target })]]),
      PART,
    );
    expect(result.target).toMatchObject({ kind: 'external', url: expected });
  });

  it('resolves Strict hyperlink relationship namespaces and types', () => {
    const result = parseXlsxHyperlink(
      element({ ref: 'A1' }, 'strictLink', true),
      'strict',
      new Map([
        [
          'strictLink',
          relationship(
            { id: 'strictLink', target: 'https://strict.example/' },
            true,
          ),
        ],
      ]),
      PART,
    );
    expect(result.target).toEqual({
      kind: 'external',
      url: 'https://strict.example/',
    });
  });

  it.each([undefined, '', '$A$1', 'B2:A1', 'XFE1'])(
    'rejects hyperlink range %#',
    (ref) => {
      expect(
        capture(() =>
          parseXlsxHyperlink(
            element(ref === undefined ? {} : { ref, location: 'A1' }),
            'transitional',
            new Map(),
            PART,
          ),
        ).diagnostic.message,
      ).toBe('Worksheet hyperlink range is invalid');
    },
  );

  it('requires either an internal location or external relationship', () => {
    expect(
      capture(() =>
        parseXlsxHyperlink(
          element({ ref: 'A1' }),
          'transitional',
          new Map(),
          PART,
        ),
      ).diagnostic.message,
    ).toBe('Worksheet hyperlink target is missing');
    expect(
      capture(() =>
        parseXlsxHyperlink(
          element({ location: '', ref: 'A1' }),
          'transitional',
          new Map(),
          PART,
        ),
      ).diagnostic.message,
    ).toBe('Worksheet hyperlink target is missing');
  });

  it.each([
    undefined,
    relationship({ mode: 'internal', target: 'xl/workbook.xml' }),
    relationship({ type: `${REL_NS}/image` }),
  ])('rejects invalid hyperlink relationship %#', (candidate) => {
    const relationships =
      candidate === undefined
        ? new Map<string, XlsxRelationship>()
        : new Map([['rId1', candidate]]);
    const error = capture(() =>
      parseXlsxHyperlink(
        element({ ref: 'A1' }, 'rId1'),
        'transitional',
        relationships,
        PART,
      ),
    );
    expect(error.diagnostic).toMatchObject({
      code: 'invalid-relationship-target',
      message: 'Worksheet hyperlink relationship is invalid',
      part: PART,
    });
  });

  it.each([
    'javascript:alert(1)',
    'data:text/html,unsafe',
    'file:///etc/passwd',
    '/relative/path',
    'ftp://example.com/file',
  ])('rejects unsafe external protocol %#', (target) => {
    const error = capture(() =>
      parseXlsxHyperlink(
        element({ ref: 'A1' }, 'rId1'),
        'transitional',
        new Map([['rId1', relationship({ target })]]),
        PART,
      ),
    );
    expect(error.diagnostic).toEqual({
      code: 'security-rejected-content',
      message: 'Worksheet hyperlink protocol is not allowed',
      part: PART,
      relationshipType: 'hyperlink',
      severity: 'error',
    });
  });
});
