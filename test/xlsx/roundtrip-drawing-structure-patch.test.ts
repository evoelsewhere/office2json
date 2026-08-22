import { describe, expect, it } from 'vitest';

import { patchXlsxDrawingStructure } from '../../src/formats/xlsx/roundtrip/drawing-structure-patch';
import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';

const PART = 'xl/drawings/drawing1.xml';
const NS =
  'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

function bytes(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function capture(action: () => unknown): XlsxWriteError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected drawing structural patch to fail');
}

function marker(name: 'from' | 'to', column: number, row: number): string {
  return `<xdr:${name}><xdr:col>${column}</xdr:col><xdr:colOff>10</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>20</xdr:rowOff></xdr:${name}>`;
}

describe('XLSX drawing structural patching', () => {
  it('transforms one-cell and two-cell anchors by edit mode', () => {
    const source = `<xdr:wsDr xmlns:xdr="${NS}"><xdr:oneCellAnchor>${marker('from', 1, 1)}<xdr:ext cx="1" cy="1"/><xdr:clientData/></xdr:oneCellAnchor><xdr:twoCellAnchor>${marker('from', 0, 0)}<wrapper>${marker('to', 7, 7)}</wrapper><x:to xmlns:x="urn:foreign"><xdr:col>8</xdr:col><xdr:row>8</xdr:row></x:to>${marker('to', 2, 2)}<xdr:clientData/></xdr:twoCellAnchor><xdr:twoCellAnchor foo="x" editAs="oneCell">${marker('from', 1, 1)}${marker('to', 3, 3)}<xdr:clientData/></xdr:twoCellAnchor><xdr:twoCellAnchor editAs="absolute">${marker('from', 4, 4)}${marker('to', 5, 5)}<xdr:clientData/></xdr:twoCellAnchor><xdr:absoluteAnchor><xdr:pos x="1" y="1"/><xdr:ext cx="1" cy="1"/><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`;
    const result = patchXlsxDrawingStructure(
      bytes(source),
      [
        { count: 1, index: 2, kind: 'insert-rows', operationId: 'rows' },
        {
          count: 1,
          index: 2,
          kind: 'insert-columns',
          operationId: 'columns',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const output = new TextDecoder().decode(result.data);
    expect(output).toContain(marker('from', 2, 2));
    expect(output).toContain(marker('from', 0, 0));
    expect(output).toContain(marker('to', 3, 3));
    expect(output).toContain(
      `foo="x" editAs="oneCell">${marker('from', 2, 2)}${marker('to', 4, 4)}`,
    );
    expect(output).toContain(
      `editAs="absolute">${marker('from', 4, 4)}${marker('to', 5, 5)}`,
    );
    expect(output).toContain('<xdr:absoluteAnchor>');
    expect(result.patchCount).toBe(8);
  });

  it('fails closed when a two-cell or one-cell-size anchor leaves the grid', () => {
    const deleted = `<xdr:wsDr xmlns:xdr="${NS}"><xdr:twoCellAnchor>${marker('from', 1, 1)}${marker('to', 1, 1)}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
    expect(
      capture(() =>
        patchXlsxDrawingStructure(
          bytes(deleted),
          [
            {
              count: 1,
              index: 2,
              kind: 'delete-rows',
              operationId: 'delete-drawing',
            },
          ],
          defaultXlsxWriteLimits(),
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({
      featureClass: 'drawing-anchor-deletion',
      message: 'XLSX structural edit would delete a drawing anchor',
      range: '2:2',
    });
    const overflow = `<xdr:wsDr xmlns:xdr="${NS}"><xdr:twoCellAnchor editAs="oneCell">${marker('from', 16383, 0)}${marker('to', 16383, 0)}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
    expect(() =>
      patchXlsxDrawingStructure(
        bytes(overflow),
        [
          {
            count: 1,
            index: 16384,
            kind: 'insert-columns',
            operationId: 'overflow-drawing',
          },
        ],
        defaultXlsxWriteLimits(),
        PART,
      ),
    ).toThrow(XlsxWriteError);
  });

  it('selects only owned prefixed anchors and marker fields', () => {
    const source = `<xdr:wsDr xmlns:xdr="${NS}" xmlns:x="urn:foreign"><wrapper><xdr:oneCellAnchor>${marker('from', 7, 7)}</xdr:oneCellAnchor></wrapper><x:oneCellAnchor>${marker('from', 8, 8)}</x:oneCellAnchor><xdr:oneCellAnchor><wrapper>${marker('from', 6, 6)}</wrapper><x:from><xdr:col>9</xdr:col><xdr:row>9</xdr:row></x:from><xdr:from><wrapper><xdr:col>6</xdr:col><xdr:row>6</xdr:row></wrapper><x:col>9</x:col><x:row>9</x:row><other>8</other><xdr:col> 0 </xdr:col><xdr:colOff>10</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>20</xdr:rowOff></xdr:from></xdr:oneCellAnchor><wrapper><xdr:from><xdr:col>5</xdr:col><xdr:row>5</xdr:row></xdr:from></wrapper></xdr:wsDr>`;
    const result = patchXlsxDrawingStructure(
      bytes(source),
      [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'owned' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    const output = new TextDecoder().decode(result.data);
    expect(output).toContain('<xdr:col> 0 </xdr:col>');
    expect(output).toContain('<xdr:row>1</xdr:row>');
    expect(output).toContain(marker('from', 7, 7));
    expect(output).toContain(marker('from', 8, 8));
    expect(output).toContain(marker('from', 6, 6));
  });

  it('rejects malformed drawing roots, modes, anchors, and markers', () => {
    for (const [source, message] of [
      ['<wrong/>', 'XLSX drawing root cannot patch structure'],
      ['<!--empty-->', 'XLSX drawing root cannot patch structure'],
      [
        `<xdr:wsDr xmlns:xdr="${NS}"><xdr:oneCellAnchor/></xdr:wsDr>`,
        'XLSX structural drawing anchor is invalid',
      ],
      [
        `<xdr:wsDr xmlns:xdr="${NS}"><xdr:oneCellAnchor><xdr:from><xdr:row>0</xdr:row></xdr:from></xdr:oneCellAnchor></xdr:wsDr>`,
        'XLSX structural drawing marker is invalid',
      ],
      [
        `<xdr:wsDr xmlns:xdr="${NS}"><xdr:twoCellAnchor>${marker('from', 0, 0)}</xdr:twoCellAnchor></xdr:wsDr>`,
        'XLSX structural drawing anchor is invalid',
      ],
      [
        `<xdr:wsDr xmlns:xdr="${NS}"><xdr:twoCellAnchor editAs="bad">${marker('from', 0, 0)}${marker('to', 1, 1)}</xdr:twoCellAnchor></xdr:wsDr>`,
        'XLSX structural drawing edit mode is invalid',
      ],
      [
        `<xdr:wsDr xmlns:xdr="${NS}"><xdr:oneCellAnchor><xdr:from><xdr:col>bad</xdr:col><xdr:row>0</xdr:row></xdr:from></xdr:oneCellAnchor></xdr:wsDr>`,
        'XLSX structural drawing marker is invalid',
      ],
      [
        `<xdr:wsDr xmlns:xdr="${NS}"><xdr:oneCellAnchor><xdr:from><xdr:col>x1</xdr:col><xdr:row>0</xdr:row></xdr:from></xdr:oneCellAnchor></xdr:wsDr>`,
        'XLSX structural drawing marker is invalid',
      ],
      [
        `<xdr:wsDr xmlns:xdr="${NS}"><xdr:oneCellAnchor><xdr:from><xdr:col>1x</xdr:col><xdr:row>0</xdr:row></xdr:from></xdr:oneCellAnchor></xdr:wsDr>`,
        'XLSX structural drawing marker is invalid',
      ],
      [
        `<xdr:wsDr xmlns:xdr="${NS}"><xdr:oneCellAnchor><xdr:from><xdr:col>1.0</xdr:col><xdr:row>0</xdr:row></xdr:from></xdr:oneCellAnchor></xdr:wsDr>`,
        'XLSX structural drawing marker is invalid',
      ],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxDrawingStructure(
            bytes(source),
            [
              {
                count: 1,
                index: 1,
                kind: 'insert-rows',
                operationId: 'bad-drawing',
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ featureClass: 'drawing-structure-xml', message });
    }
  });

  it('preserves no-op bytes and enforces exact resource limits', () => {
    const source = `<xdr:wsDr xmlns:xdr="${NS}"><xdr:oneCellAnchor>${marker('from', 0, 0)}</xdr:oneCellAnchor></xdr:wsDr>`;
    const noOp = patchXlsxDrawingStructure(
      bytes(source),
      [{ count: 1, index: 3, kind: 'insert-rows', operationId: 'no-op' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(noOp.data)).toBe(source);
    expect(noOp.patchCount).toBe(0);
    const input = bytes(source);
    const empty = patchXlsxDrawingStructure(
      input,
      [],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(empty.data).toEqual(input);
    expect(empty.data).not.toBe(input);
    const variable = `<xdr:wsDr xmlns:xdr="${NS}"><xdr:oneCellAnchor>${marker('from', 0, 9)}</xdr:oneCellAnchor><xdr:oneCellAnchor>${marker('from', 0, 99)}</xdr:oneCellAnchor></xdr:wsDr>`;
    expect(
      new TextDecoder().decode(
        patchXlsxDrawingStructure(
          bytes(variable),
          [
            {
              count: 1,
              index: 1,
              kind: 'insert-rows',
              operationId: 'variable-length',
            },
          ],
          defaultXlsxWriteLimits(),
          PART,
        ).data,
      ),
    ).toBe(
      `<xdr:wsDr xmlns:xdr="${NS}"><xdr:oneCellAnchor>${marker('from', 0, 10)}</xdr:oneCellAnchor><xdr:oneCellAnchor>${marker('from', 0, 100)}</xdr:oneCellAnchor></xdr:wsDr>`,
    );
    const request = {
      count: 1,
      index: 1,
      kind: 'insert-rows' as const,
      operationId: 'limits',
    };
    const changed = patchXlsxDrawingStructure(
      bytes(source),
      [request],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(changed.patchCount).toBe(1);
    expect(changed.patchBytes).toBe(new TextEncoder().encode('1').byteLength);
    expect(() =>
      patchXlsxDrawingStructure(
        bytes(source),
        [request],
        {
          ...defaultXlsxWriteLimits(),
          maxGeneratedXmlBytes: changed.data.byteLength,
          maxPatchBytes: changed.patchBytes,
          maxPatchCount: changed.patchCount,
        },
        PART,
      ),
    ).not.toThrow();
    for (const [limitName, limit] of [
      ['maxGeneratedXmlBytes', changed.data.byteLength - 1],
      ['maxPatchBytes', 0],
      ['maxPatchCount', 0],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxDrawingStructure(
            bytes(source),
            [request],
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ limitName });
    }
  });
});
