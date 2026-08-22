import { describe, expect, it } from 'vitest';

import { ZipEntrySizeLimitError } from '../../src/common/archive/read-entry';
import {
  XmlComplexityLimitError,
  XmlStructureError,
} from '../../src/common/xml/types';
import {
  XlsxStreamingXmlParser,
  type XlsxStreamingXmlLimits,
  type XlsxXmlElement,
} from '../../src/formats/xlsx/internal/streaming-xml';

const UNBOUNDED: XlsxStreamingXmlLimits = {
  maxBytes: Number.MAX_SAFE_INTEGER,
  maxDepth: Number.MAX_SAFE_INTEGER,
  maxNodes: Number.MAX_SAFE_INTEGER,
};

function utf16(value: string, littleEndian: boolean, bom = true): Uint8Array {
  const output = new Uint8Array(value.length * 2 + (bom ? 2 : 0));
  let offset = 0;
  if (bom) {
    output[0] = littleEndian ? 0xff : 0xfe;
    output[1] = littleEndian ? 0xfe : 0xff;
    offset = 2;
  }
  for (const character of value) {
    const code = character.charCodeAt(0);
    output[offset] = littleEndian ? code & 0xff : code >> 8;
    output[offset + 1] = littleEndian ? code >> 8 : code & 0xff;
    offset += 2;
  }
  return output;
}

function parseChunks(
  bytes: Uint8Array,
  chunkSizes: readonly number[],
  limits: XlsxStreamingXmlLimits = UNBOUNDED,
): { closed: string[]; opened: XlsxXmlElement[]; text: string } {
  const opened: XlsxXmlElement[] = [];
  const closed: string[] = [];
  let text = '';
  const parser = new XlsxStreamingXmlParser(limits, {
    closeElement: (node) => closed.push(`${node.namespace}|${node.localName}`),
    openElement: (node) => opened.push(node),
    text: (value) => {
      text += value;
    },
  });
  let offset = 0;
  for (const size of chunkSizes) {
    parser.write(bytes.subarray(offset, offset + size));
    offset += size;
  }
  if (offset < bytes.byteLength) parser.write(bytes.subarray(offset));
  parser.close();
  return { closed, opened, text };
}

function captureFailure(action: () => void): Error {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(Error);
    return error as Error;
  }
  throw new Error('Expected action to fail');
}

describe('XLSX streaming XML parser', () => {
  it('preserves namespaces, expanded attributes, entities, CDATA, and text', () => {
    const source = new TextEncoder().encode(
      '<?xml version="1.0" encoding="UTF-8"?><x:root xmlns:x="urn:x" xmlns:y="urn:y" y:id="7"><x:child>A&amp;😀<![CDATA[<B>]]></x:child></x:root>',
    );
    const parsed = parseChunks(source, [1, 1, 1, 2, 3, 5, 8, 13]);

    expect(
      parsed.opened.map(({ localName, namespace }) => ({
        localName,
        namespace,
      })),
    ).toEqual([
      { localName: 'root', namespace: 'urn:x' },
      { localName: 'child', namespace: 'urn:x' },
    ]);
    expect(parsed.opened[0]?.attributes.get('{urn:y}id')).toBe('7');
    expect(parsed.closed).toEqual(['urn:x|child', 'urn:x|root']);
    expect(parsed.text).toBe('A&😀<B>');
  });

  it.each([
    [
      'UTF-8 BOM',
      new Uint8Array([
        0xef,
        0xbb,
        0xbf,
        ...new TextEncoder().encode('<r>é</r>'),
      ]),
    ],
    ['UTF-16LE BOM', utf16('<r>é</r>', true)],
    ['UTF-16BE BOM', utf16('<r>é</r>', false)],
    ['UTF-16LE inferred', utf16('<r>é</r>', true, false)],
    ['UTF-16BE inferred', utf16('<r>é</r>', false, false)],
  ] as const)('decodes %s across every initial split', (_name, bytes) => {
    for (let split = 0; split <= Math.min(bytes.byteLength, 8); split += 1) {
      const parsed = parseChunks(bytes, [split, 1, 1]);
      expect(parsed.text).toBe('é');
      expect(parsed.closed).toEqual(['|r']);
    }
  });

  it.each([
    [
      'UTF-8 without encoding',
      '<?xml version="1.0"?><r/>',
      (value: string) => new TextEncoder().encode(value),
    ],
    [
      'UTF-16LE generic',
      '<?xml version="1.0" encoding="UTF-16"?><r/>',
      (value: string) => utf16(value, true),
    ],
    [
      'UTF-16LE exact',
      '<?xml version="1.0" encoding="UTF-16LE"?><r/>',
      (value: string) => utf16(value, true),
    ],
    [
      'UTF-16BE generic',
      '<?xml version="1.0" encoding="utf-16"?><r/>',
      (value: string) => utf16(value, false),
    ],
    [
      'UTF-16BE exact',
      '<?xml version="1.0" encoding="utf-16be"?><r/>',
      (value: string) => utf16(value, false),
    ],
  ] as const)('accepts a matching %s declaration', (_name, source, encode) => {
    expect(parseChunks(encode(source), [3]).closed).toEqual(['|r']);
  });

  it.each([
    [
      '<?xml version="1.0" encoding="UTF-16"?><r/>',
      (value: string) => new TextEncoder().encode(value),
    ],
    [
      '<?xml version="1.0" encoding="UTF-8"?><r/>',
      (value: string) => utf16(value, true),
    ],
    [
      '<?xml version="1.0" encoding="ISO-8859-1"?><r/>',
      (value: string) => new TextEncoder().encode(value),
    ],
  ] as const)('rejects mismatched declaration %#', (source, encode) => {
    expect(() => parseChunks(encode(source), [1, 2, 3])).toThrow(
      'XML declaration encoding does not match source bytes',
    );
  });

  it('rejects XML versions other than 1.0', () => {
    expect(() =>
      parseChunks(new TextEncoder().encode('<?xml version="1.1"?><r/>'), [3]),
    ).toThrow('Only XML 1.0 documents are supported');
  });

  it.each([
    '<!DOCTYPE r><r/>',
    '<!DOCTYPE r [<!ENTITY x "value">]><r>&x;</r>',
    '<r>&unknown;</r>',
    '<r><x></r>',
    '<r xmlns:a="urn:x" xmlns:b="urn:x" a:id="1" b:id="2"/>',
    '<r/><r/>',
    '',
  ])('rejects unsafe or malformed XML %#', (source) => {
    expect(() =>
      parseChunks(new TextEncoder().encode(source), [0, 1, 2]),
    ).toThrow(XmlStructureError);
  });

  it('reports doctype and parser failures with stable messages and causes', () => {
    const doctype = captureFailure(() =>
      parseChunks(new TextEncoder().encode('<!DOCTYPE r><r/>'), [3]),
    );
    expect(doctype.message).toBe(
      'XML document type declarations are not allowed',
    );
    expect(doctype.name).toBe('XmlStructureError');

    const malformed = captureFailure(() =>
      parseChunks(new TextEncoder().encode('<r><x></r>'), [3]),
    );
    expect(malformed.message).toBe('Invalid XML structure');
    expect(malformed.name).toBe('XmlStructureError');
    expect(malformed.cause).toBeInstanceOf(Error);
  });

  it('does not infer UTF-16 through leading whitespace or partial signatures', () => {
    expect(() => parseChunks(utf16(' <r/>', true, false), [3])).toThrow(
      XmlStructureError,
    );
    expect(() => parseChunks(utf16(' <r/>', false, false), [3])).toThrow(
      XmlStructureError,
    );
    expect(parseChunks(new TextEncoder().encode(' <r/>'), [3]).closed).toEqual([
      '|r',
    ]);
  });

  it.each([
    new Uint8Array([0x3c, 0x72, 0x3e, 0xff, 0x3c, 0x2f, 0x72, 0x3e]),
    new Uint8Array([0xff, 0xfe, 0x3c]),
    new Uint8Array([0xfe, 0xff, 0x00]),
  ])('rejects truncated or invalid encoded bytes %#', (bytes) => {
    const error = captureFailure(() => parseChunks(bytes, [1, 1]));
    expect(error.message).toMatch(/^Invalid UTF-(?:8|16LE|16BE) XML$/);
    expect(error.name).toBe('XmlStructureError');
    expect(error.cause).toBeInstanceOf(TypeError);
  });

  it.each([
    new Uint8Array([0x20, 0xfe, 0x3c, 0x72, 0x2f, 0x3e]),
    new Uint8Array([0xff, 0x20, 0x3c, 0x72, 0x2f, 0x3e]),
    new Uint8Array([0x20, 0xff, 0x3c, 0x72, 0x2f, 0x3e]),
    new Uint8Array([0xfe, 0x20, 0x3c, 0x72, 0x2f, 0x3e]),
  ])('does not accept a partial UTF-16 BOM %#', (bytes) => {
    const error = captureFailure(() => parseChunks(bytes, [3]));
    expect(error.message).toBe('Invalid UTF-8 XML');
  });

  it('begins parsing as soon as three signature bytes are available', () => {
    const opened: string[] = [];
    const parser = new XlsxStreamingXmlParser(UNBOUNDED, {
      openElement: ({ localName }) => opened.push(localName),
    });
    parser.write(new TextEncoder().encode('<r>'));
    expect(opened).toEqual(['r']);
    parser.write(new TextEncoder().encode('</r>'));
    parser.close();
  });

  it('accepts exact byte, depth, and node limits and reports accounting', () => {
    const bytes = new TextEncoder().encode('<r><a/><b/></r>');
    let consumedBytes = 0;
    let consumedNodes = 0;
    expect(() =>
      parseChunks(bytes, [1, 2, 3], {
        consumeBytes: (value) => {
          consumedBytes += value;
        },
        consumeNodes: (value) => {
          consumedNodes += value;
        },
        maxBytes: bytes.byteLength,
        maxDepth: 2,
        maxNodes: 3,
      }),
    ).not.toThrow();
    expect(consumedBytes).toBe(bytes.byteLength);
    expect(consumedNodes).toBe(3);
  });

  it.each([
    [
      { maxBytes: 14, maxDepth: 2, maxNodes: 3 },
      ZipEntrySizeLimitError,
      { actual: 15, limit: 14 },
    ],
    [
      { maxBytes: 15, maxDepth: 1, maxNodes: 3 },
      XmlComplexityLimitError,
      { actual: 2, limit: 1, limitName: 'maxXmlDepth' },
    ],
    [
      { maxBytes: 15, maxDepth: 2, maxNodes: 2 },
      XmlComplexityLimitError,
      { actual: 3, limit: 2, limitName: 'maxXmlNodes' },
    ],
  ] as const)('rejects one-over stream limit %#', (limits, type, details) => {
    const bytes = new TextEncoder().encode('<r><a/><b/></r>');
    let thrown: unknown;
    try {
      parseChunks(bytes, [bytes.byteLength], limits);
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(type);
    expect(thrown).toMatchObject(details);
  });

  it('propagates aggregate budget and sink failures without wrapping', () => {
    const budgetError = new Error('aggregate budget');
    const sinkError = new Error('sink failure');
    expect(() =>
      parseChunks(new TextEncoder().encode('<r/>'), [1], {
        ...UNBOUNDED,
        consumeBytes: () => {
          throw budgetError;
        },
      }),
    ).toThrow(budgetError);
    const parser = new XlsxStreamingXmlParser(UNBOUNDED, {
      openElement: () => {
        throw sinkError;
      },
    });
    expect(() => parser.write(new TextEncoder().encode('<r/>'))).toThrow(
      sinkError,
    );
  });

  it('rejects writes and duplicate close after completion', () => {
    const parser = new XlsxStreamingXmlParser(UNBOUNDED, {});
    parser.write(new TextEncoder().encode('<r/>'));
    parser.close();

    expect(() => parser.write(new Uint8Array())).toThrow(
      'XLSX streaming XML parser is closed',
    );
    expect(() => parser.close()).toThrow('XLSX streaming XML parser is closed');
  });
});
