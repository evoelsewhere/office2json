import { describe, expect, it } from 'vitest';

import {
  assertXlsxSafeCellEditSource,
  xlsxCellEditFormulaIsUnsafe,
  xlsxCellEditPartIsSafe,
  xlsxCellEditRelationshipIsSafe,
  xlsxCellEditRelationshipKind,
  xlsxPlannedCell,
} from '../../src/formats/xlsx/roundtrip/cell-edit-policy';
import {
  generateBoundedXlsxZip,
  verifyXlsxCellEditR1Parts,
  xlsxCellEditPartTopologyEqual,
} from '../../src/formats/xlsx/roundtrip/cell-edit-verification';
import { xlsxStructuralRelationshipTargets } from '../../src/formats/xlsx/roundtrip/cell-edit-package';
import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import type { XlsxPackageGraph } from '../../src/formats/xlsx/roundtrip/internal/package-graph';
import { readXlsxRoundTrip } from '../../src/formats/xlsx/roundtrip/read-snapshot';
import type { XlsxRoundTripDocument } from '../../src/formats/xlsx/roundtrip/types';
import { verifyXlsxEditedSemantics } from '../../src/formats/xlsx/roundtrip/write';
import { assertXlsxWriteOptions } from '../../src/formats/xlsx/roundtrip/write-options';
import { createIndependentXlsx } from '../black-box/xlsx-package';

function graph(overrides: Partial<XlsxPackageGraph> = {}): XlsxPackageGraph {
  return {
    conformance: 'transitional',
    containsActiveContent: false,
    containsDigitalSignatures: false,
    containsExternalRelationships: false,
    containsOpaqueContent: false,
    manifestHash: 'a'.repeat(64),
    parts: [
      {
        byteLength: 1,
        contentType: 'application/xml',
        name: 'one.xml',
        relationshipPart: false,
        sha256: '1'.repeat(64),
      },
      {
        byteLength: 2,
        contentType: 'application/xml',
        name: 'two.xml',
        relationshipPart: false,
        sha256: '2'.repeat(64),
      },
    ],
    relationships: [],
    ...overrides,
  };
}

function capture(action: () => unknown): XlsxWriteError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected XLSX cell package verification to fail');
}

class FakeZipStream {
  private data?: (chunk: Uint8Array) => void;
  private end?: () => void;
  private error?: (error: unknown) => void;
  paused = false;

  constructor(
    private readonly chunks: Uint8Array[],
    private readonly failure = false,
  ) {}

  on(event: string, listener: (value?: unknown) => void): FakeZipStream {
    if (event === 'data') {
      this.data = listener;
    } else if (event === 'end') {
      this.end = listener;
    } else {
      this.error = listener;
    }
    return this;
  }

  pause(): FakeZipStream {
    this.paused = true;
    return this;
  }

  resume(): FakeZipStream {
    if (this.failure) {
      this.error?.(new Error('private failure'));
      return this;
    }
    for (const chunk of this.chunks) {
      this.data?.(chunk);
      if (this.paused) return this;
    }
    this.end?.();
    return this;
  }
}

function fakeArchive(
  stream: FakeZipStream,
): Parameters<typeof generateBoundedXlsxZip>[0] {
  return {
    generateInternalStream: () => stream,
  } as unknown as Parameters<typeof generateBoundedXlsxZip>[0];
}

describe('XLSX cell-edit package verification', () => {
  it('selects exact internal structural relationship targets once in authored order', () => {
    const relationshipType =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
    const relationships: XlsxPackageGraph['relationships'] = [
      {
        id: 'exact-a',
        mode: 'internal',
        owner: 'xl/drawings/drawing1.xml',
        target: 'xl/charts/chart1.xml',
        type: relationshipType,
      },
      {
        id: 'duplicate-a',
        mode: 'internal',
        owner: 'xl/drawings/drawing1.xml',
        target: 'xl/charts/chart1.xml',
        type: relationshipType,
      },
      {
        id: 'exact-b',
        mode: 'internal',
        owner: 'xl/drawings/drawing1.xml',
        target: 'xl/charts/chart2.xml',
        type: relationshipType,
      },
      {
        id: 'foreign-owner',
        mode: 'internal',
        owner: 'xl/drawings/drawing2.xml',
        target: 'xl/charts/foreign.xml',
        type: relationshipType,
      },
      {
        id: 'external',
        mode: 'external',
        owner: 'xl/drawings/drawing1.xml',
        target: 'https://example.invalid/chart.xml',
        type: relationshipType,
      },
      {
        id: 'wrong-type',
        mode: 'internal',
        owner: 'xl/drawings/drawing1.xml',
        target: 'xl/charts/wrong.xml',
        type: 'http://example.invalid/relationships/chart',
      },
    ];
    expect(
      xlsxStructuralRelationshipTargets(
        { relationships },
        'xl/drawings/drawing1.xml',
        relationshipType,
      ),
    ).toEqual(['xl/charts/chart1.xml', 'xl/charts/chart2.xml']);
  });
  it.each([
    'application/vnd.openxmlformats-officedocument.custom-properties+xml',
    'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
    'application/vnd.openxmlformats-officedocument.theme+xml',
    'application/vnd.openxmlformats-package.core-properties+xml',
    'application/vnd.openxmlformats-package.content-types+xml',
    'application/vnd.openxmlformats-package.relationships+xml',
  ])('accepts safe part content type %s', (contentType) => {
    expect(xlsxCellEditPartIsSafe(contentType)).toBe(true);
  });

  it('rejects unlisted part content types', () => {
    expect(xlsxCellEditPartIsSafe('application/xml')).toBe(false);
    expect(xlsxCellEditPartIsSafe('')).toBe(false);
  });

  it.each([
    'core-properties',
    'custom-properties',
    'extended-properties',
    'hyperlink',
    'officeDocument',
    'sharedStrings',
    'styles',
    'theme',
    'worksheet',
  ])('accepts safe relationship kind %s', (kind) => {
    const type = `http://example.invalid/relationships/${kind}`;
    expect(xlsxCellEditRelationshipKind(type)).toBe(kind);
    expect(xlsxCellEditRelationshipIsSafe(type)).toBe(true);
  });

  it('rejects unlisted relationship kinds', () => {
    expect(xlsxCellEditRelationshipIsSafe('customXml')).toBe(false);
    expect(xlsxCellEditRelationshipKind('plain')).toBe('plain');
  });

  it('reports every unsafe source policy with bounded provenance', () => {
    const safePart = {
      byteLength: 1,
      contentType:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      name: 'xl/worksheets/sheet1.xml',
      relationshipPart: false,
      sha256: '1'.repeat(64),
    };
    const safe = graph({
      parts: [safePart],
      relationships: [
        {
          id: 'one',
          mode: 'internal',
          owner: null,
          target: safePart.name,
          type: 'http://example.invalid/relationships/worksheet',
        },
      ],
    });
    expect(() => assertXlsxSafeCellEditSource(safe, {})).not.toThrow();
    const cases: Array<
      readonly [
        XlsxPackageGraph,
        Parameters<typeof assertXlsxSafeCellEditSource>[1],
        Record<string, unknown>,
      ]
    > = [
      [
        { ...safe, containsDigitalSignatures: true },
        {},
        {
          code: 'signed-package-conflict',
          featureClass: 'digital-signature',
          message:
            'Signed XLSX packages cannot be edited without invalidating signatures',
        },
      ],
      [
        { ...safe, containsActiveContent: true },
        {},
        {
          code: 'preservation-conflict',
          featureClass: 'active-content',
          message:
            'Active XLSX package content cannot enter the cell-edit closure',
        },
      ],
      [
        { ...safe, containsOpaqueContent: true },
        {},
        {
          code: 'opaque-content-conflict',
          featureClass: 'opaque-content',
          message:
            'Opaque XLSX content requires acknowledgement and a proven independent closure',
        },
      ],
      [
        { ...safe, containsOpaqueContent: true },
        { acknowledgeOpaqueContent: true },
        {
          code: 'opaque-content-conflict',
          featureClass: 'opaque-content',
          message:
            'Opaque XLSX content has no proven independent cell-edit closure',
        },
      ],
      [
        {
          ...safe,
          parts: [{ ...safePart, contentType: 'application/xml' }],
        },
        {},
        {
          code: 'opaque-content-conflict',
          featureClass: 'unsupported-part',
          message: 'XLSX cell editing encountered an unsupported package part',
          part: safePart.name,
        },
      ],
      [
        {
          ...safe,
          relationships: [
            {
              ...safe.relationships[0]!,
              owner: 'xl/workbook.xml',
              type: 'http://example.invalid/relationships/customXml',
            },
          ],
        },
        {},
        {
          code: 'opaque-content-conflict',
          featureClass: 'unsupported-relationship',
          message:
            'XLSX cell editing encountered an unsupported relationship dependency',
          part: 'xl/workbook.xml',
        },
      ],
      [
        {
          ...safe,
          relationships: [
            {
              ...safe.relationships[0]!,
              type: 'http://example.invalid/relationships/customXml',
            },
          ],
        },
        {},
        {
          code: 'opaque-content-conflict',
          featureClass: 'unsupported-relationship',
          message:
            'XLSX cell editing encountered an unsupported relationship dependency',
        },
      ],
    ];
    for (const [source, options, expected] of cases) {
      expect(
        capture(() => assertXlsxSafeCellEditSource(source, options)).diagnostic,
      ).toMatchObject(expected);
    }
  });

  it('allows table dependencies only for a proven structural closure', () => {
    const worksheetPart = {
      byteLength: 1,
      contentType:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      name: 'xl/worksheets/sheet1.xml',
      relationshipPart: false,
      sha256: '1'.repeat(64),
    };
    const tablePart = {
      byteLength: 1,
      contentType:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
      name: 'xl/tables/table1.xml',
      relationshipPart: false,
      sha256: '2'.repeat(64),
    };
    const source = graph({
      parts: [worksheetPart, tablePart],
      relationships: [
        {
          id: 'table',
          mode: 'internal',
          owner: worksheetPart.name,
          target: tablePart.name,
          type: 'http://example.invalid/relationships/table',
        },
      ],
    });
    expect(
      capture(() => assertXlsxSafeCellEditSource(source, {})).diagnostic,
    ).toMatchObject({
      featureClass: 'unsupported-part',
      part: tablePart.name,
    });
    expect(() => assertXlsxSafeCellEditSource(source, {}, true)).not.toThrow();
    const customPart = {
      ...tablePart,
      contentType: 'application/example+xml',
      name: 'xl/custom/example.xml',
    };
    expect(
      capture(() =>
        assertXlsxSafeCellEditSource(
          { ...source, parts: [worksheetPart, customPart] },
          {},
          true,
        ),
      ).diagnostic,
    ).toMatchObject({
      featureClass: 'unsupported-part',
      part: customPart.name,
    });
    expect(
      capture(() =>
        assertXlsxSafeCellEditSource(
          {
            ...source,
            parts: [worksheetPart],
            relationships: [
              {
                ...source.relationships[0]!,
                target: worksheetPart.name,
                type: 'http://example.invalid/relationships/customXml',
              },
            ],
          },
          {},
          true,
        ),
      ).diagnostic,
    ).toMatchObject({ featureClass: 'unsupported-relationship' });
  });

  it('allows comment dependencies only for a proven structural closure', () => {
    const contentTypes = [
      'application/vnd.ms-excel.person+xml',
      'application/vnd.ms-excel.threadedcomments+xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
      'application/vnd.openxmlformats-officedocument.vmlDrawing',
    ];
    const relationshipKinds = [
      'comments',
      'person',
      'threadedComment',
      'vmlDrawing',
    ];
    for (const [index, contentType] of contentTypes.entries()) {
      const part = {
        byteLength: 1,
        contentType,
        name: `xl/comment-part-${index}.xml`,
        relationshipPart: false,
        sha256: String(index + 1).repeat(64),
      };
      const source = graph({
        containsOpaqueContent: index === 3,
        parts: [part],
        relationships: [
          {
            id: `comment-${index}`,
            mode: 'internal',
            owner: null,
            target: part.name,
            type: `http://example.invalid/relationships/${relationshipKinds[index]}`,
          },
        ],
      });
      expect(
        capture(() => assertXlsxSafeCellEditSource(source, {})).diagnostic,
      ).toMatchObject({
        featureClass: index === 3 ? 'opaque-content' : 'unsupported-part',
      });
      expect(() =>
        assertXlsxSafeCellEditSource(source, {}, false, true),
      ).not.toThrow();
    }
    const worksheetPart = {
      byteLength: 1,
      contentType:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      name: 'xl/worksheets/sheet1.xml',
      relationshipPart: false,
      sha256: 'a'.repeat(64),
    };
    const vmlPart = {
      byteLength: 1,
      contentType: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
      name: 'xl/drawings/comments.vml',
      relationshipPart: false,
      sha256: 'b'.repeat(64),
    };
    const vmlGraph = graph({
      containsOpaqueContent: true,
      parts: [worksheetPart, vmlPart],
      relationships: [],
    });
    expect(() =>
      assertXlsxSafeCellEditSource(vmlGraph, {}, false, true),
    ).not.toThrow();
    expect(
      capture(() =>
        assertXlsxSafeCellEditSource(
          {
            ...vmlGraph,
            parts: [worksheetPart],
          },
          {},
          false,
          true,
        ),
      ).diagnostic,
    ).toMatchObject({ featureClass: 'opaque-content' });
    const customPart = {
      ...worksheetPart,
      contentType: 'application/example',
      name: 'xl/opaque.bin',
    };
    expect(
      capture(() =>
        assertXlsxSafeCellEditSource(
          { ...vmlGraph, parts: [vmlPart, customPart] },
          {},
          false,
          true,
        ),
      ).diagnostic,
    ).toMatchObject({ featureClass: 'opaque-content' });
    expect(
      capture(() =>
        assertXlsxSafeCellEditSource(
          {
            ...vmlGraph,
            relationships: [
              {
                id: 'custom',
                mode: 'internal',
                owner: worksheetPart.name,
                target: vmlPart.name,
                type: 'http://example.invalid/relationships/customXml',
              },
            ],
          },
          {},
          false,
          true,
        ),
      ).diagnostic,
    ).toMatchObject({ featureClass: 'unsupported-relationship' });
  });

  it('allows drawing and image dependencies only for structural closure', () => {
    for (const [index, contentType, relationshipKind, name] of [
      [
        0,
        'application/vnd.openxmlformats-officedocument.drawing+xml',
        'drawing',
        'xl/drawings/drawing1.xml',
      ],
      [1, 'image/png', 'image', 'xl/media/image1.png'],
      [
        2,
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
        'chart',
        'xl/charts/chart1.xml',
      ],
    ] as const) {
      const part = {
        byteLength: 1,
        contentType,
        name,
        relationshipPart: false,
        sha256: String(index + 1).repeat(64),
      };
      const source = graph({
        parts: [part],
        relationships: [
          {
            id: `drawing-${index}`,
            mode: 'internal',
            owner: null,
            target: part.name,
            type: `http://example.invalid/relationships/${relationshipKind}`,
          },
        ],
      });
      expect(
        capture(() => assertXlsxSafeCellEditSource(source, {})).diagnostic,
      ).toMatchObject({ featureClass: 'unsupported-part' });
      expect(() =>
        assertXlsxSafeCellEditSource(source, {}, false, false, true),
      ).not.toThrow();
    }
    const drawingPart = {
      byteLength: 1,
      contentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
      name: 'xl/drawings/drawing1.xml',
      relationshipPart: false,
      sha256: 'd'.repeat(64),
    };
    const customPart = {
      ...drawingPart,
      contentType: 'application/example',
      name: 'xl/custom.bin',
    };
    expect(
      capture(() =>
        assertXlsxSafeCellEditSource(
          { ...graph(), parts: [customPart] },
          {},
          false,
          false,
          true,
        ),
      ).diagnostic,
    ).toMatchObject({ featureClass: 'unsupported-part' });
    expect(
      capture(() =>
        assertXlsxSafeCellEditSource(
          {
            ...graph(),
            parts: [drawingPart],
            relationships: [
              {
                id: 'custom',
                mode: 'internal',
                owner: null,
                target: drawingPart.name,
                type: 'http://example.invalid/relationships/customXml',
              },
            ],
          },
          {},
          false,
          false,
          true,
        ),
      ).diagnostic,
    ).toMatchObject({ featureClass: 'unsupported-relationship' });
  });

  it.each([
    '[Book.xlsx]Sheet1!A1',
    'call(A1)',
    'DDE(A1)',
    'Exec (A1)',
    'FILTERXML(A1,A2)',
    'HYPERLINK(A1)',
    'REGISTER.ID(A1)',
    'rtd(A1)',
    'WEBSERVICE(A1)',
  ])('blocks external-capable formula %s', (formula) => {
    expect(xlsxCellEditFormulaIsUnsafe(formula)).toBe(true);
  });

  it.each(['SUM(A1:A2)', 'MY_WEBSERVICE_NAME+A1', '"text"'])(
    'accepts formula without an active external token %s',
    (formula) => {
      expect(xlsxCellEditFormulaIsUnsafe(formula)).toBe(false);
    },
  );

  it('compares part topology independently from bytes', () => {
    const source = graph().parts[0]!;
    expect(xlsxCellEditPartTopologyEqual(source, { ...source })).toBe(true);
    for (const changed of [
      { ...source, contentType: 'other' },
      { ...source, name: 'other.xml' },
      { ...source, relationshipPart: true },
    ]) {
      expect(xlsxCellEditPartTopologyEqual(source, changed)).toBe(false);
    }
    expect(
      xlsxCellEditPartTopologyEqual(source, {
        ...source,
        byteLength: 99,
        sha256: 'f'.repeat(64),
      }),
    ).toBe(true);
  });

  it('creates literal R1 copy and patch evidence', () => {
    const source = graph();
    const output = graph({
      parts: [
        source.parts[0]!,
        {
          ...source.parts[1]!,
          byteLength: 3,
          sha256: '3'.repeat(64),
        },
      ],
    });
    expect(
      verifyXlsxCellEditR1Parts(source, output, new Set(['two.xml'])),
    ).toEqual([
      {
        byteLength: 1,
        disposition: 'copy',
        name: 'one.xml',
        sha256: '1'.repeat(64),
        sourceByteLength: 1,
        sourceSha256: '1'.repeat(64),
      },
      {
        byteLength: 3,
        disposition: 'patch',
        name: 'two.xml',
        sha256: '3'.repeat(64),
        sourceByteLength: 2,
        sourceSha256: '2'.repeat(64),
      },
    ]);
  });

  it('creates literal R1 add evidence for an allowed relationship closure', () => {
    const source = graph({
      relationships: [
        {
          id: 'link',
          mode: 'external',
          owner: 'two.xml',
          target: 'https://old.invalid/',
          type: 'hyperlink',
        },
      ],
    });
    const addedPart = {
      byteLength: 4,
      contentType: 'application/vnd.openxmlformats-package.relationships+xml',
      name: '_rels/two.xml.rels',
      relationshipPart: true,
      sha256: '4'.repeat(64),
    };
    const output = graph({
      parts: [
        source.parts[0]!,
        {
          ...source.parts[1]!,
          byteLength: 3,
          sha256: '3'.repeat(64),
        },
        addedPart,
      ],
      relationships: [
        {
          id: 'link',
          mode: 'external',
          owner: 'two.xml',
          target: 'https://new.invalid/',
          type: 'hyperlink',
        },
      ],
    });
    expect(
      verifyXlsxCellEditR1Parts(
        source,
        output,
        new Set(['two.xml']),
        new Set([addedPart.name]),
        new Set(['two.xml']),
      ),
    ).toEqual([
      {
        byteLength: 1,
        disposition: 'copy',
        name: 'one.xml',
        sha256: '1'.repeat(64),
        sourceByteLength: 1,
        sourceSha256: '1'.repeat(64),
      },
      {
        byteLength: 3,
        disposition: 'patch',
        name: 'two.xml',
        sha256: '3'.repeat(64),
        sourceByteLength: 2,
        sourceSha256: '2'.repeat(64),
      },
      {
        byteLength: 4,
        disposition: 'add',
        name: addedPart.name,
        sha256: '4'.repeat(64),
      },
    ]);
  });

  it('never classifies a source part as added evidence', () => {
    const source = graph();
    const output = graph({
      parts: [...source.parts, { ...source.parts[0]! }],
    });
    expect(
      verifyXlsxCellEditR1Parts(
        source,
        output,
        new Set(),
        new Set(['one.xml']),
      ).map((part) => part.disposition),
    ).toEqual(['copy', 'copy', 'copy']);
  });

  it('rejects an undeclared added R1 part with exact provenance', () => {
    const source = graph();
    const unexpected = {
      byteLength: 3,
      contentType: 'application/xml',
      name: 'three.xml',
      relationshipPart: false,
      sha256: '3'.repeat(64),
    };
    const output = graph({ parts: [...source.parts, unexpected] });
    expect(
      capture(() =>
        verifyXlsxCellEditR1Parts(
          source,
          output,
          new Set(),
          new Set(['different.xml']),
        ),
      ).diagnostic,
    ).toMatchObject({
      message: 'XLSX edited package contains an unexpected part',
      part: unexpected.name,
    });
  });

  it('rejects every R1 topology and copy mismatch', () => {
    const source = graph();
    const mismatches: Array<
      readonly [XlsxPackageGraph, string, string | undefined]
    > = [
      [graph({ conformance: 'strict' }), 'topology differs', undefined],
      [graph({ parts: [source.parts[0]!] }), 'topology differs', undefined],
      [
        graph({
          relationships: [
            {
              id: 'one',
              mode: 'external',
              owner: null,
              target: 'https://example.invalid',
              type: 'hyperlink',
            },
          ],
        }),
        'topology differs',
        undefined,
      ],
      [
        graph({
          parts: [source.parts[0]!, { ...source.parts[1]!, name: 'other.xml' }],
        }),
        'removed a source part',
        'two.xml',
      ],
      [
        graph({
          parts: [
            source.parts[0]!,
            { ...source.parts[1]!, contentType: 'other' },
          ],
        }),
        'unexpected part',
        'two.xml',
      ],
      [
        graph({
          parts: [
            source.parts[0]!,
            { ...source.parts[1]!, sha256: '3'.repeat(64) },
          ],
        }),
        'copied part bytes changed',
        'two.xml',
      ],
      [
        graph({
          parts: [source.parts[0]!, { ...source.parts[1]!, byteLength: 3 }],
        }),
        'copied part bytes changed',
        'two.xml',
      ],
    ];
    for (const [output, message, part] of mismatches) {
      const error = capture(() =>
        verifyXlsxCellEditR1Parts(source, output, new Set()),
      );
      expect(error.diagnostic.message).toContain(message);
      if (part === undefined) {
        expect(error.diagnostic).not.toHaveProperty('part');
      } else {
        expect(error.diagnostic.part).toBe(part);
      }
    }
  });

  it('streams ZIP output at and over its exact byte limit', async () => {
    const exact = new FakeZipStream([
      new Uint8Array([1, 2]),
      new Uint8Array([3]),
    ]);
    await expect(
      generateBoundedXlsxZip(fakeArchive(exact), 3),
    ).resolves.toEqual(new Uint8Array([1, 2, 3]));
    const over = new FakeZipStream([
      new Uint8Array([1, 2]),
      new Uint8Array([3, 4]),
    ]);
    await expect(
      generateBoundedXlsxZip(fakeArchive(over), 3),
    ).rejects.toMatchObject({
      diagnostic: {
        actual: 4,
        code: 'resource-limit-exceeded',
        limit: 3,
        limitName: 'maxOutputBytes',
        message: 'XLSX generated package exceeds its output byte limit',
      },
    });
    expect(over.paused).toBe(true);
  });

  it('redacts ZIP generator failures behind a structured error', async () => {
    await expect(
      generateBoundedXlsxZip(fakeArchive(new FakeZipStream([], true)), 10),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'generated-package-invalid',
        message: 'Failed to generate the XLSX output package',
      },
    });
  });

  it('independently accepts exact output semantics and rejects all mismatches', async () => {
    const bytes = await createIndependentXlsx();
    const snapshot = await readXlsxRoundTrip(bytes);
    await expect(
      verifyXlsxEditedSemantics(bytes, snapshot.document, {}),
    ).resolves.toBeUndefined();
    await expect(
      verifyXlsxEditedSemantics(new Uint8Array([1]), snapshot.document, {}),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'semantic-verification-failed',
        fidelity: 'R2',
        message: 'Strictly reparsing the edited XLSX package failed',
      },
    });
    const noSheets: XlsxRoundTripDocument = {
      ...snapshot.document,
      sheets: [],
    };
    await expect(
      verifyXlsxEditedSemantics(bytes, noSheets, {}),
    ).rejects.toMatchObject({
      diagnostic: {
        fidelity: 'R2',
        message:
          'Edited XLSX worksheet inventory differs from the operation preview',
      },
    });
    const changed: XlsxRoundTripDocument = structuredClone(snapshot.document);
    changed.workbook.dateSystem = '1904';
    await expect(
      verifyXlsxEditedSemantics(bytes, changed, {}),
    ).rejects.toMatchObject({
      diagnostic: {
        fidelity: 'R2',
        message: 'Edited XLSX semantics differ from the operation preview',
      },
    });
    await expect(
      verifyXlsxEditedSemantics(bytes, snapshot.document, {
        readerLimits: { maxReturnedCells: 2, maxScannedCells: 3 },
      }),
    ).rejects.toThrow('Strictly reparsing the edited XLSX package failed');
  });

  it('resolves stable planned cells exactly and rejects invalid targets', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const document = structuredClone(snapshot.document);
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    sheet.rows[0]!.cells.push({
      address: 'B1',
      column: 2,
      content: { kind: 'blank' },
    });
    expect(xlsxPlannedCell(document, sheet.key, 'B1').address).toBe('B1');
    expect(() => xlsxPlannedCell(document, sheet.key, 'Z99')).toThrow(
      'Expected an existing XLSX cell edit target',
    );
    const chartDocument: XlsxRoundTripDocument = {
      ...document,
      sheets: [
        {
          index: 0,
          key: sheet.key,
          kind: 'chart-sheet',
          name: 'Chart',
          payload: 'full-sheet',
          state: 'visible',
        },
      ],
    };
    expect(() => xlsxPlannedCell(chartDocument, sheet.key, 'A1')).toThrow(
      'Expected an XLSX worksheet edit target',
    );
  });

  it('validates every write option branch exactly', () => {
    const valid = [
      {},
      { acknowledgeOpaqueContent: true },
      { acknowledgeOpaqueContent: false },
      { minimumEditedFidelity: 'R1' as const },
      { minimumEditedFidelity: 'R2' as const },
      { minimumEditedFidelity: 'R3' as const },
      { limits: {} },
      { readerLimits: {} },
    ];
    for (const options of valid) {
      expect(() => assertXlsxWriteOptions(options)).not.toThrow();
    }
    const invalid: Array<readonly [unknown, string]> = [
      [null, 'plain object'],
      [[], 'plain object'],
      [Object.create(null), 'plain object'],
      [{ unknown: true }, 'Unknown XLSX write option unknown'],
      [{ acknowledgeOpaqueContent: 'yes' }, 'must be boolean'],
      [{ minimumEditedFidelity: 'R0' }, 'minimumEditedFidelity is invalid'],
    ];
    for (const [options, message] of invalid) {
      expect(() => assertXlsxWriteOptions(options as never)).toThrow(message);
    }
  });
});
