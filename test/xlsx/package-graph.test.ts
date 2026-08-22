import { describe, expect, it } from 'vitest';

import {
  defaultXlsxResourceLimits,
  XlsxResourceLimitError,
} from '../../src/formats/xlsx/internal/resource-limits';
import {
  assertXlsxGraphRelationshipTargets,
  consumeXlsxGraphExpandedBytes,
  inspectXlsxPackageGraph,
  xlsxActiveContent,
  xlsxLexicalCompare,
  xlsxOpaqueContent,
  xlsxPackagePartIsXml,
  xlsxRelationshipOwner,
} from '../../src/formats/xlsx/roundtrip/internal/package-graph';
import { createIndependentXlsx } from '../black-box/xlsx-package';

function captureResource(action: () => unknown): XlsxResourceLimitError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxResourceLimitError);
    return error as XlsxResourceLimitError;
  }
  throw new Error('Expected graph resource limit failure');
}

describe('XLSX canonical package graph', () => {
  it('inventories canonical parts, content types, hashes, and relationships', async () => {
    const bytes = await createIndependentXlsx();
    const graph = await inspectXlsxPackageGraph(
      bytes,
      defaultXlsxResourceLimits(),
    );

    expect(graph.conformance).toBe('transitional');
    expect(graph.parts.map((part) => part.name)).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/_rels/workbook.xml.rels',
      'xl/sharedStrings.xml',
      'xl/styles.xml',
      'xl/workbook.xml',
      'xl/worksheets/sheet1.xml',
    ]);
    expect(graph.parts[0]).toMatchObject({
      contentType: 'application/vnd.openxmlformats-package.content-types+xml',
      relationshipPart: false,
    });
    expect(graph.parts[1]).toMatchObject({
      contentType: 'application/vnd.openxmlformats-package.relationships+xml',
      relationshipPart: true,
    });
    expect(
      graph.parts.every(
        (part) => part.byteLength > 0 && /^[0-9a-f]{64}$/u.test(part.sha256),
      ),
    ).toBe(true);
    expect(
      graph.relationships.map(({ id, mode, owner, target }) => ({
        id,
        mode,
        owner,
        target,
      })),
    ).toEqual([
      {
        id: 'rIdWorkbook',
        mode: 'internal',
        owner: null,
        target: 'xl/workbook.xml',
      },
      {
        id: 'rIdSharedStrings',
        mode: 'internal',
        owner: 'xl/workbook.xml',
        target: 'xl/sharedStrings.xml',
      },
      {
        id: 'rIdSheet1',
        mode: 'internal',
        owner: 'xl/workbook.xml',
        target: 'xl/worksheets/sheet1.xml',
      },
      {
        id: 'rIdStyles',
        mode: 'internal',
        owner: 'xl/workbook.xml',
        target: 'xl/styles.xml',
      },
    ]);
    expect(graph).toMatchObject({
      containsActiveContent: false,
      containsDigitalSignatures: false,
      containsExternalRelationships: false,
      containsOpaqueContent: false,
    });
    expect(graph.manifestHash).toMatch(/^[0-9a-f]{64}$/u);
  });

  it('changes the source manifest hash when part bytes change', async () => {
    const first = await inspectXlsxPackageGraph(
      await createIndependentXlsx(),
      defaultXlsxResourceLimits(),
    );
    const second = await inspectXlsxPackageGraph(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
      }),
      defaultXlsxResourceLimits(),
    );
    expect(second.manifestHash).not.toBe(first.manifestHash);
  });

  it('accepts the exact expanded-byte total and rejects one under', async () => {
    const bytes = await createIndependentXlsx();
    const graph = await inspectXlsxPackageGraph(
      bytes,
      defaultXlsxResourceLimits(),
    );
    const total = graph.parts.reduce((sum, part) => sum + part.byteLength, 0);
    await expect(
      inspectXlsxPackageGraph(bytes, {
        ...defaultXlsxResourceLimits(),
        maxTotalUncompressedBytes: total,
      }),
    ).resolves.toMatchObject({ parts: graph.parts });
    await expect(
      inspectXlsxPackageGraph(bytes, {
        ...defaultXlsxResourceLimits(),
        maxTotalUncompressedBytes: total - 1,
      }),
    ).rejects.toMatchObject({
      actual: total,
      limit: total - 1,
      limitName: 'maxTotalUncompressedBytes',
      name: 'XlsxResourceLimitError',
    });
  });

  it('rejects an undeclared package part with a structured graph error', async () => {
    const bytes = await createIndependentXlsx({
      'xl/unknown.bin': new Uint8Array([1]),
    });
    await expect(
      inspectXlsxPackageGraph(bytes, defaultXlsxResourceLimits()),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'relationship-graph-invalid',
        message: 'XLSX package part has no declared content type',
        part: 'xl/unknown.bin',
      },
      name: 'XlsxWriteError',
    });
  });

  it('rejects missing internal relationship targets and owners', async () => {
    await expect(
      inspectXlsxPackageGraph(
        await createIndependentXlsx({
          'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/missing.xml"/></Relationships>`,
        }),
        defaultXlsxResourceLimits(),
      ),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'relationship-graph-invalid',
        message: 'XLSX internal relationship target part is missing',
        part: 'xl/workbook.xml',
      },
    });
    try {
      assertXlsxGraphRelationshipTargets(
        [{ name: 'target.xml' }],
        [
          {
            id: 'one',
            mode: 'external',
            owner: 'missing.xml',
            target: 'https://example.invalid',
            type: 'hyperlink',
          },
        ],
      );
      throw new Error('Expected owner validation to fail');
    } catch (error) {
      expect(error).toMatchObject({
        diagnostic: {
          message: 'XLSX relationship owner part is missing',
          part: 'missing.xml',
        },
      });
    }
    expect(() =>
      assertXlsxGraphRelationshipTargets(
        [{ name: 'other.xml' }],
        [
          {
            id: 'root',
            mode: 'internal',
            owner: null,
            target: 'missing.xml',
            type: 'officeDocument',
          },
        ],
      ),
    ).toThrow('XLSX internal relationship target part is missing');
    try {
      assertXlsxGraphRelationshipTargets(
        [{ name: 'other.xml' }],
        [
          {
            id: 'root',
            mode: 'internal',
            owner: null,
            target: 'missing.xml',
            type: 'officeDocument',
          },
        ],
      );
    } catch (error) {
      expect(error).toMatchObject({ diagnostic: { part: '_rels/.rels' } });
    }
  });

  it('detects digital signature part identities', async () => {
    const graph = await inspectXlsxPackageGraph(
      await createIndependentXlsx({
        '_xmlsignatures/sig1.xml': '<Signature/>',
      }),
      defaultXlsxResourceLimits(),
    );
    expect(graph.containsDigitalSignatures).toBe(true);
  });

  it('stream-validates unknown XML parts without retaining their trees', async () => {
    await expect(
      inspectXlsxPackageGraph(
        await createIndependentXlsx({
          'xl/custom.xml': '<custom><unclosed></custom>',
        }),
        defaultXlsxResourceLimits(),
      ),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'xml-parse-failed',
        part: 'xl/custom.xml',
      },
      name: 'XlsxParseError',
    });
  });

  it.each([
    ['_rels/.rels', null],
    ['xl/_rels/workbook.xml.rels', 'xl/workbook.xml'],
    ['xl/worksheets/_rels/sheet1.xml.rels', 'xl/worksheets/sheet1.xml'],
    ['xl/workbook.xml', undefined],
    ['xl/_rels/workbook.xml', undefined],
    ['_rels/other.rels', undefined],
    ['/_rels/other.rels', undefined],
  ] as const)('resolves relationship owner %s', (part, expected) => {
    expect(xlsxRelationshipOwner(part)).toBe(expected);
  });

  it.each([
    ['xl/vbaProject.bin', 'application/octet-stream'],
    ['xl/activeX/activeX1.bin', 'application/octet-stream'],
    ['xl/embeddings/object1.bin', 'application/octet-stream'],
    ['xl/custom.bin', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'],
    ['xl/custom.bin', 'application/vnd.ms-office.activeX'],
    [
      'xl/custom.bin',
      'application/vnd.openxmlformats-officedocument.oleObject',
    ],
  ])('detects active content %s / %s', (name, contentType) => {
    expect(xlsxActiveContent(name, contentType)).toBe(true);
    expect(
      xlsxActiveContent(name.toUpperCase(), contentType.toUpperCase()),
    ).toBe(true);
  });

  it('does not classify ordinary XML as active content', () => {
    expect(xlsxActiveContent('xl/workbook.xml', 'application/xml')).toBe(false);
  });

  it.each([
    ['[Content_Types].xml', 'application/xml', false],
    ['_rels/.rels', 'application/octet-stream', false],
    ['xl/workbook.xml', 'application/octet-stream', false],
    ['xl/media/image1.png', 'image/png', false],
    ['xl/custom.dat', 'application/xml', false],
    ['xl/custom.bin', 'application/octet-stream', true],
  ] as const)('classifies opaque content %s as %s', (name, type, expected) => {
    expect(xlsxOpaqueContent(name, type)).toBe(expected);
  });

  it.each([
    ['a', 'b', -1],
    ['b', 'a', 1],
    ['a', 'a', 0],
    ['A', 'a', -1],
  ] as const)('compares %s and %s lexically', (left, right, expected) => {
    expect(xlsxLexicalCompare(left, right)).toBe(expected);
  });

  it.each([
    ['xl/custom.xml', 'application/octet-stream', true],
    ['xl/custom.bin', 'application/xml', true],
    ['xl/custom.bin', 'application/example+xml', true],
    ['xl/custom.bin', 'application/octet-stream', false],
  ] as const)('classifies XML package part %s / %s', (name, type, expected) => {
    expect(xlsxPackagePartIsXml(name, type)).toBe(expected);
  });

  it('accounts actual expanded bytes with safe exact boundaries', () => {
    const limits = {
      ...defaultXlsxResourceLimits(),
      maxTotalUncompressedBytes: 3,
    };
    expect(consumeXlsxGraphExpandedBytes(1, 2, limits, 'part.bin')).toBe(3);
    expect(
      captureResource(() =>
        consumeXlsxGraphExpandedBytes(1, 3, limits, 'part.bin'),
      ),
    ).toMatchObject({
      actual: 4,
      limit: 3,
      limitName: 'maxTotalUncompressedBytes',
      part: 'part.bin',
    });
    expect(
      captureResource(() =>
        consumeXlsxGraphExpandedBytes(
          Number.MAX_SAFE_INTEGER,
          1,
          {
            ...limits,
            maxTotalUncompressedBytes: Number.MAX_SAFE_INTEGER,
          },
          'part.bin',
        ),
      ),
    ).toMatchObject({ actual: 2 ** 53 });
  });
});
