import { afterEach, describe, expect, it, vi } from 'vitest';

import {
  parseXlsx,
  parseXlsxWithDiagnostics,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
  XlsxParseError,
} from '../../src/formats/xlsx';
import {
  parseXlsxDrawingInteger,
  type XlsxMediaReader,
  XlsxMediaSession,
} from '../../src/formats/xlsx/internal/drawing';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import {
  createIndependentXlsx,
  type XlsxBlackBoxOverrides,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const DRAWING_NS =
  'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const DRAWING_MAIN_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const DRAWING_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}drawing`;
const IMAGE_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}image`;
const IMAGE_BYTES = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

const CONTENT_TYPES = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`;

const WORKSHEET = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_TYPE.slice(0, -1)}"><sheetData/><drawing r:id="rIdDrawing"/></worksheet>`;
const WORKSHEET_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="rIdDrawing" Type="${DRAWING_RELATIONSHIP}" Target="../drawings/drawing1.xml"/></Relationships>`;
const DRAWING_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="image" Type="${IMAGE_RELATIONSHIP}" Target="../media/image1.png"/></Relationships>`;

function marker(
  column: number,
  row: number,
  columnOffset = 25_400,
  rowOffset = 12_700,
): string {
  return `<xdr:col>${column}</xdr:col><xdr:colOff>${columnOffset}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOffset}</xdr:rowOff>`;
}

function picture(id: number, relationshipId = 'image'): string {
  return `<xdr:pic>
    <xdr:nvPicPr><xdr:cNvPr id="${id}" name="Image ${id}" descr="Description ${id}" hidden="${id === 2 ? 1 : 0}"/><xdr:cNvPicPr/></xdr:nvPicPr>
    <xdr:blipFill><a:blip r:embed="${relationshipId}"/><a:srcRect l="1000" t="2000" r="3000" b="4000"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
    <xdr:spPr><a:xfrm rot="60000" flipH="1" flipV="0"><a:off x="0" y="0"/><a:ext cx="127000" cy="254000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
  </xdr:pic>`;
}

function drawingXml(secondRelationship = 'image'): string {
  return `<xdr:wsDr xmlns:xdr="${DRAWING_NS}" xmlns:a="${DRAWING_MAIN_NS}" xmlns:r="${XLSX_OFFICE_REL_TYPE.slice(0, -1)}">
    <xdr:twoCellAnchor editAs="oneCell"><xdr:from>${marker(0, 0)}</xdr:from><xdr:to>${marker(2, 2)}</xdr:to>${picture(1)}<xdr:clientData/></xdr:twoCellAnchor>
    <xdr:oneCellAnchor><xdr:from>${marker(3, 3)}</xdr:from><xdr:ext cx="381000" cy="508000"/>${picture(2, secondRelationship)}<xdr:clientData/></xdr:oneCellAnchor>
    <xdr:absoluteAnchor><xdr:pos x="-12700" y="-25400"/><xdr:ext cx="635000" cy="762000"/>${picture(3)}<xdr:clientData/></xdr:absoluteAnchor>
  </xdr:wsDr>`;
}

function drawingDocument(anchors: string): string {
  return `<xdr:wsDr xmlns:xdr="${DRAWING_NS}" xmlns:a="${DRAWING_MAIN_NS}" xmlns:r="${XLSX_OFFICE_REL_TYPE.slice(0, -1)}">${anchors}</xdr:wsDr>`;
}

function oneCellAnchor(id: number, column: number, row: number): string {
  return `<xdr:oneCellAnchor><xdr:from>${marker(column, row)}</xdr:from><xdr:ext cx="381000" cy="508000"/>${picture(id)}<xdr:clientData/></xdr:oneCellAnchor>`;
}

function twoCellAnchor(
  id: number,
  fromColumn: number,
  fromRow: number,
  toColumn: number,
  toRow: number,
): string {
  return twoCellMarkers(
    id,
    marker(fromColumn, fromRow),
    marker(toColumn, toRow),
  );
}

function twoCellMarkers(id: number, from: string, to: string): string {
  return `<xdr:twoCellAnchor><xdr:from>${from}</xdr:from><xdr:to>${to}</xdr:to>${picture(id)}<xdr:clientData/></xdr:twoCellAnchor>`;
}

function shape(id: number, text = 'Shape &amp; text'): string {
  return `<xdr:sp>
    <xdr:nvSpPr><xdr:cNvPr id="${id}" name="Shape ${id}" descr="Shape description"/><xdr:cNvSpPr/></xdr:nvSpPr>
    <xdr:spPr><a:xfrm rot="-60000" flipH="false" flipV="true"><a:off x="12700" y="25400"/><a:ext cx="127000" cy="254000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="aabbcc"/></a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:prstDash val="dash"/></a:ln></xdr:spPr>
    <xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p><a:p><a:r><a:t>Second</a:t></a:r></a:p></xdr:txBody>
  </xdr:sp>`;
}

function connector(id: number): string {
  return `<xdr:cxnSp>
    <xdr:nvCxnSpPr><xdr:cNvPr id="${id}" name="Connector ${id}"/><xdr:cNvCxnSpPr><a:stCxn id="1" idx="2"/><a:endCxn id="3" idx="4"/></xdr:cNvCxnSpPr></xdr:nvCxnSpPr>
    <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12700" cy="12700"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>
  </xdr:cxnSp>`;
}

function group(id: number): string {
  return `<xdr:grpSp>
    <xdr:nvGrpSpPr><xdr:cNvPr id="${id}" name="Group ${id}"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>
    <xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="254000" cy="254000"/><a:chOff x="12700" y="25400"/><a:chExt cx="127000" cy="127000"/></a:xfrm></xdr:grpSpPr>
    ${shape(id + 1, 'Nested')}${connector(id + 2)}
  </xdr:grpSp>`;
}

function oneCellObject(object: string): string {
  return `<xdr:oneCellAnchor><xdr:from>${marker(0, 0)}</xdr:from><xdr:ext cx="381000" cy="508000"/>${object}<xdr:clientData/></xdr:oneCellAnchor>`;
}

function parts(overrides: XlsxBlackBoxOverrides = {}): XlsxBlackBoxOverrides {
  return {
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/drawings/_rels/drawing1.xml.rels': DRAWING_RELS,
    'xl/drawings/drawing1.xml': drawingXml(),
    'xl/media/image1.png': IMAGE_BYTES,
    'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS,
    'xl/worksheets/sheet1.xml': WORKSHEET,
    ...overrides,
  };
}

async function bytes(
  overrides: XlsxBlackBoxOverrides = {},
): Promise<Uint8Array> {
  return createIndependentXlsx(parts(overrides));
}

async function capture(
  overrides: XlsxBlackBoxOverrides,
  options: Parameters<typeof parseXlsx>[1] = { errorMode: 'strict' },
): Promise<XlsxParseError> {
  try {
    await parseXlsx(await bytes(overrides), options);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected XLSX drawing parsing to fail');
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe('XLSX drawing anchors and images', () => {
  it('parses ordered anchors and safe image metadata without binary output by default', async () => {
    const document = await parseXlsx(await bytes(), { errorMode: 'strict' });
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    expect(sheet.kind === 'worksheet' ? sheet.drawings : []).toMatchObject([
      {
        editAs: 'one-cell',
        extent: { height: 20, width: 10 },
        from: { column: 1, columnOffset: 2, row: 1, rowOffset: 1 },
        kind: 'two-cell',
        object: {
          contentType: 'image/png',
          crop: { bottom: 4, left: 1, right: 3, top: 2 },
          description: 'Description 1',
          hidden: false,
          id: 1,
          kind: 'image',
          name: 'Image 1',
          part: 'xl/media/image1.png',
          transform: {
            flipHorizontal: true,
            flipVertical: false,
            rotation: 1,
          },
        },
        selectionRelation: 'full-sheet',
        to: { column: 3, columnOffset: 2, row: 3, rowOffset: 1 },
      },
      {
        extent: { height: 40, width: 30 },
        from: { column: 4, columnOffset: 2, row: 4, rowOffset: 1 },
        kind: 'one-cell',
        object: { hidden: true, id: 2 },
        selectionRelation: 'full-sheet',
      },
      {
        extent: { height: 60, width: 50 },
        kind: 'absolute',
        object: { id: 3 },
        position: { x: -1, y: -2 },
        selectionRelation: 'full-sheet',
      },
    ]);
    expect(JSON.parse(JSON.stringify(document))).toEqual(document);
  });

  it('returns canonical Base64 and charges repeated returned media bytes', async () => {
    const document = await parseXlsx(await bytes(), {
      errorMode: 'strict',
      imageMode: 'base64',
      limits: { maxMediaBytes: 24 },
    });
    const sheet = document.sheets[0]!;
    const images =
      sheet.kind === 'worksheet'
        ? sheet.drawings.map((item) => item.object)
        : [];
    expect(images).toHaveLength(3);
    expect(images[0]).toMatchObject({
      base64: 'data:image/png;base64,iVBORw0KGgo=',
      byteLength: 8,
    });
    expect(
      (
        await capture(
          {},
          {
            errorMode: 'strict',
            imageMode: 'base64',
            limits: { maxMediaBytes: 23 },
          },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 24,
      code: 'resource-limit-exceeded',
      limit: 23,
      limitName: 'maxMediaBytes',
    });
  });

  it('deduplicates blob URLs, transfers success ownership, and revokes failed parse URLs', async () => {
    const create = vi
      .spyOn(URL, 'createObjectURL')
      .mockReturnValue('blob:xlsx-image');
    const revoke = vi
      .spyOn(URL, 'revokeObjectURL')
      .mockImplementation(() => {});
    const document = await parseXlsx(await bytes(), {
      errorMode: 'strict',
      imageMode: 'both',
      limits: { maxMediaBytes: 24 },
    });
    const sheet = document.sheets[0]!;
    expect(create).toHaveBeenCalledTimes(1);
    expect(revoke).not.toHaveBeenCalled();
    expect(
      sheet.kind === 'worksheet'
        ? sheet.drawings.map((item) =>
            item.object.kind === 'image' ? item.object.blobUrl : undefined,
          )
        : [],
    ).toEqual(['blob:xlsx-image', 'blob:xlsx-image', 'blob:xlsx-image']);

    create.mockClear();
    revoke.mockClear();
    const result = await parseXlsxWithDiagnostics(
      await bytes({
        'xl/drawings/drawing1.xml': drawingXml('missing'),
      }),
      { imageMode: 'blob', limits: { maxMediaBytes: 24 } },
    );
    expect(create).toHaveBeenCalledTimes(1);
    expect(revoke).toHaveBeenCalledExactlyOnceWith('blob:xlsx-image');
    const failedSheet = result.document.sheets[0]!;
    expect(
      failedSheet.kind === 'worksheet' ? failedSheet.drawings : [],
    ).toEqual([]);
    expect(result.diagnostics).toMatchObject([
      { message: 'Image relationship is invalid', severity: 'warning' },
    ]);
  });

  it('filters cell anchors and preserves absolute worksheet-global drawings', async () => {
    const document = await parseXlsx(await bytes(), {
      errorMode: 'strict',
      selection: { ranges: { Sheet1: ['B2'] } },
    });
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.drawings.map((drawing) => [
            drawing.kind,
            drawing.selectionRelation,
          ])
        : [],
    ).toEqual([
      ['two-cell', 'intersects-selection'],
      ['absolute', 'worksheet-global'],
    ]);
  });

  it('does not load image payloads for excluded cell anchors', async () => {
    const document = await parseXlsx(await bytes(), {
      errorMode: 'strict',
      imageMode: 'base64',
      limits: { maxMediaBytes: 16 },
      selection: { ranges: { Sheet1: ['B2'] } },
    });
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.drawings.map((drawing) =>
            drawing.object.kind === 'image'
              ? drawing.object.byteLength
              : undefined,
          )
        : [],
    ).toEqual([8, 8]);
  });

  it('preserves authored order across interleaved anchor kinds', async () => {
    const interleaved = drawingDocument(
      `${oneCellAnchor(1, 0, 0)}${twoCellAnchor(2, 1, 1, 2, 2)}${oneCellAnchor(3, 3, 3)}`,
    );
    const document = await parseXlsx(
      await bytes({ 'xl/drawings/drawing1.xml': interleaved }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.drawings.map((drawing) => drawing.object.id)
        : [],
    ).toEqual([1, 2, 3]);
  });

  it('parses shape geometry, text, paint, line, and object transform', async () => {
    const document = await parseXlsx(
      await bytes({
        'xl/drawings/drawing1.xml': drawingDocument(oneCellObject(shape(10))),
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined,
    ).toStrictEqual({
      description: 'Shape description',
      fill: { color: { kind: 'rgb', value: 'AABBCC' }, kind: 'solid' },
      geometry: { kind: 'preset', preset: 'roundRect' },
      hidden: false,
      id: 10,
      kind: 'shape',
      line: {
        dash: 'dash',
        fill: { color: { kind: 'scheme', value: 'accent1' }, kind: 'solid' },
        width: 1,
      },
      name: 'Shape 10',
      text: 'Shape & text\nSecond',
      transform: {
        flipHorizontal: false,
        flipVertical: true,
        height: 20,
        rotation: -1,
        width: 10,
        x: 1,
        y: 2,
      },
    });
  });

  it('parses connector endpoints and no-fill semantics', async () => {
    const document = await parseXlsx(
      await bytes({
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(connector(20)),
        ),
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined,
    ).toMatchObject({
      endConnection: { shapeId: 3, site: 4 },
      fill: { kind: 'none' },
      geometry: { kind: 'preset', preset: 'line' },
      id: 20,
      kind: 'connector',
      line: { fill: { kind: 'none' } },
      startConnection: { shapeId: 1, site: 2 },
    });
  });

  it('parses nested groups in authored child order', async () => {
    const document = await parseXlsx(
      await bytes({
        'xl/drawings/drawing1.xml': drawingDocument(oneCellObject(group(30))),
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const object =
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
    expect(object).toMatchObject({
      children: [
        { id: 31, kind: 'shape' },
        { id: 32, kind: 'connector' },
      ],
      id: 30,
      kind: 'group',
      transform: {
        childHeight: 10,
        childWidth: 10,
        childX: 1,
        childY: 2,
        height: 20,
        width: 20,
      },
    });
  });

  it('preserves interleaved nested group order and optional descriptions', async () => {
    const interleaved = group(170).replace(
      `${shape(171, 'Nested')}${connector(172)}`,
      `${shape(171, 'One')}${connector(172)}${shape(173, 'Two').replace(' descr="Shape description"', '')}`,
    );
    const document = await parseXlsx(
      await bytes({
        'xl/drawings/drawing1.xml': drawingDocument(oneCellObject(interleaved)),
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const object =
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
    expect(object?.kind).toBe('group');
    if (object?.kind !== 'group') throw new Error('Expected drawing group');
    expect(object.children.map((child) => child.id)).toEqual([171, 172, 173]);
    expect(object.children[2]).not.toHaveProperty('description');
  });

  it('parses nested picture and group child kinds', async () => {
    const nested = group(180).replace(
      `${shape(181, 'Nested')}${connector(182)}`,
      `${picture(181)}${group(182)}`,
    );
    const document = await parseXlsx(
      await bytes({
        'xl/drawings/drawing1.xml': drawingDocument(oneCellObject(nested)),
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const object =
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
    expect(object?.kind).toBe('group');
    if (object?.kind !== 'group') throw new Error('Expected drawing group');
    expect(object.children.map((child) => child.kind)).toEqual([
      'image',
      'group',
    ]);
  });

  it('parses custom geometry and system color metadata', async () => {
    const xml = drawingDocument(
      oneCellObject(
        shape(40, 'Custom')
          .replace(
            '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>',
            '<a:custGeom><a:avLst/><a:pathLst/></a:custGeom>',
          )
          .replace(
            '<a:srgbClr val="aabbcc"/>',
            '<a:sysClr val="windowText" lastClr="112233"/>',
          ),
      ),
    );
    const document = await parseXlsx(
      await bytes({ 'xl/drawings/drawing1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined,
    ).toMatchObject({
      fill: {
        color: {
          kind: 'system',
          lastColor: '112233',
          value: 'windowText',
        },
      },
      geometry: { kind: 'custom' },
    });
  });

  it('preserves mixed-case system fallback colors and multiple runs per paragraph', async () => {
    const xml = drawingDocument(
      oneCellObject(
        shape(45, 'First')
          .replace(
            '<a:srgbClr val="aabbcc"/>',
            '<a:sysClr val="window" lastClr="aaBBcc"/>',
          )
          .replace(
            '<a:r><a:t>First</a:t></a:r>',
            '<a:r><a:t>First</a:t></a:r><a:r><a:t> run</a:t></a:r>',
          ),
      ),
    );
    const document = await parseXlsx(
      await bytes({ 'xl/drawings/drawing1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined,
    ).toMatchObject({
      fill: { color: { lastColor: 'AABBCC' } },
      text: 'First run\nSecond',
    });
  });

  it('preserves authored shape text order, fields, and line breaks', async () => {
    const xml = drawingDocument(
      oneCellObject(
        shape(46, 'First').replace(
          '<a:r><a:t>First</a:t></a:r>',
          '<a:r><a:t>First</a:t></a:r><a:br/><a:fld id="field-1" type="text"><a:t>Field</a:t></a:fld><a:r><a:t> last</a:t></a:r>',
        ),
      ),
    );
    const document = await parseXlsx(
      await bytes({ 'xl/drawings/drawing1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined,
    ).toMatchObject({ text: 'First\nField last\nSecond' });
  });

  it('distinguishes an omitted shape fill from a malformed solid fill', async () => {
    const xml = drawingDocument(
      oneCellObject(
        shape(47).replace(
          '<a:solidFill><a:srgbClr val="aabbcc"/></a:solidFill>',
          '',
        ),
      ),
    );
    const document = await parseXlsx(
      await bytes({ 'xl/drawings/drawing1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined,
    ).not.toHaveProperty('fill');
  });

  it('preserves signed shape and group child positions', async () => {
    const xml = drawingDocument(
      oneCellObject(
        group(190)
          .replace('<a:off x="0" y="0"/>', '<a:off x="-12700" y="-25400"/>')
          .replace(
            '<a:chOff x="12700" y="25400"/>',
            '<a:chOff x="-12700" y="-25400"/>',
          ),
      ),
    );
    const document = await parseXlsx(
      await bytes({ 'xl/drawings/drawing1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const object =
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
    expect(object?.kind).toBe('group');
    if (object?.kind !== 'group') throw new Error('Expected drawing group');
    expect(object.transform).toMatchObject({
      childX: -1,
      childY: -2,
      x: -1,
      y: -2,
    });
  });

  it('rejects a chart graphic frame without typed chart data', async () => {
    const frame = `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="200" name="Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="12700" cy="12700"/></xdr:xfrm><a:graphic><a:graphicData uri="chart"/></a:graphic></xdr:graphicFrame>`;
    expect(
      (
        await capture({
          'xl/drawings/drawing1.xml': drawingDocument(oneCellObject(frame)),
        })
      ).diagnostic,
    ).toMatchObject({ message: 'Chart graphic data URI is invalid' });
  });

  it('enforces shape text and nested drawing counts exactly', async () => {
    const shapeParts = {
      'xl/drawings/drawing1.xml': drawingDocument(oneCellObject(shape(50))),
    };
    await expect(
      parseXlsx(await bytes(shapeParts), {
        errorMode: 'strict',
        limits: { maxTextCharacters: 28 },
      }),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(shapeParts, {
          errorMode: 'strict',
          limits: { maxTextCharacters: 27 },
        })
      ).diagnostic,
    ).toMatchObject({
      actual: 28,
      limit: 27,
      limitName: 'maxTextCharacters',
    });
    const groupParts = {
      'xl/drawings/drawing1.xml': drawingDocument(oneCellObject(group(60))),
    };
    await expect(
      parseXlsx(await bytes(groupParts), {
        errorMode: 'strict',
        limits: { maxDrawings: 3 },
      }),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(groupParts, {
          errorMode: 'strict',
          limits: { maxDrawings: 2 },
        })
      ).diagnostic,
    ).toMatchObject({ actual: 3, limit: 2, limitName: 'maxDrawings' });
  });

  it.each([
    'image/bmp',
    'image/gif',
    'image/jpeg',
    'image/png',
    'image/tiff',
    'image/webp',
    'image/x-emf',
    'image/x-wmf',
  ])(
    'accepts safely supported embedded image content type %s',
    async (contentType) => {
      const document = await parseXlsx(
        await bytes({
          '[Content_Types].xml': CONTENT_TYPES.replace(
            'ContentType="image/png"',
            `ContentType="${contentType}"`,
          ),
        }),
        { errorMode: 'strict' },
      );
      const sheet = document.sheets[0]!;
      expect(
        sheet.kind === 'worksheet' && sheet.drawings[0]?.object.kind === 'image'
          ? sheet.drawings[0].object.contentType
          : '',
      ).toBe(contentType);
    },
  );

  it('accepts exact crop and zero-sized same-cell anchor boundaries', async () => {
    const boundary = drawingDocument(
      twoCellAnchor(1, 0, 0, 0, 0).replace('l="1000"', 'l="-100000"'),
    );
    const document = await parseXlsx(
      await bytes({ 'xl/drawings/drawing1.xml': boundary }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.drawings[0] : undefined,
    ).toMatchObject({
      from: { column: 1, row: 1 },
      object: { crop: { left: -100 } },
      to: { column: 1, row: 1 },
    });
  });

  it.each([
    ['absolute', 'absolute'],
    ['oneCell', 'one-cell'],
    ['twoCell', 'two-cell'],
  ] as const)('normalizes two-cell edit mode %s', async (source, expected) => {
    const document = await parseXlsx(
      await bytes({
        'xl/drawings/drawing1.xml': drawingDocument(
          twoCellAnchor(1, 0, 0, 1, 1).replace(
            '<xdr:twoCellAnchor>',
            `<xdr:twoCellAnchor editAs="${source}">`,
          ),
        ),
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.editAs : undefined,
    ).toBe(expected);
  });

  it('preserves exact marker grid boundary and transform defaults', async () => {
    const minimal = drawingDocument(
      oneCellAnchor(1, 0, 0)
        .replace(' rot="60000" flipH="1" flipV="0"', '')
        .replace(' descr="Description 1"', ''),
    );
    const document = await parseXlsx(
      await bytes({ 'xl/drawings/drawing1.xml': minimal }),
      {
        errorMode: 'strict',
        limits: { maxColumnsPerWorksheet: 1, maxRowsPerWorksheet: 1 },
      },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined,
    ).toStrictEqual({
      contentType: 'image/png',
      crop: { bottom: 4, left: 1, right: 3, top: 2 },
      hidden: false,
      id: 1,
      kind: 'image',
      name: 'Image 1',
      part: 'xl/media/image1.png',
      transform: {
        flipHorizontal: false,
        flipVertical: false,
        rotation: 0,
      },
    });
  });

  it('preserves signed image rotation', async () => {
    const document = await parseXlsx(
      await bytes({
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellAnchor(1, 0, 0).replace('rot="60000"', 'rot="-60000"'),
        ),
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.drawings[0]?.object.transform.rotation
        : undefined,
    ).toBe(-1);
  });

  it('accepts decreasing offsets after advancing to another row or column', async () => {
    for (const [from, to] of [
      [marker(0, 0, 25_400, 25_400), marker(0, 1, 25_400, 12_700)],
      [marker(0, 0, 25_400, 12_700), marker(1, 0, 12_700, 12_700)],
    ] as const) {
      const document = await parseXlsx(
        await bytes({
          'xl/drawings/drawing1.xml': drawingDocument(
            twoCellMarkers(1, from, to),
          ),
        }),
        { errorMode: 'strict' },
      );
      expect(document.sheets).toHaveLength(1);
    }
  });

  it('parses textual transform booleans and absent crop defaults exactly', async () => {
    const xml = drawingDocument(
      oneCellAnchor(1, 0, 0)
        .replace('flipH="1" flipV="0"', 'flipH="true" flipV="false"')
        .replace('<a:srcRect l="1000" t="2000" r="3000" b="4000"/>', ''),
    );
    const document = await parseXlsx(
      await bytes({ 'xl/drawings/drawing1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined,
    ).toMatchObject({
      crop: { bottom: 0, left: 0, right: 0, top: 0 },
      transform: { flipHorizontal: true, flipVertical: false },
    });
  });

  it('distinguishes every reversed two-cell coordinate component', async () => {
    for (const [from, to] of [
      [marker(0, 1), marker(1, 0)],
      [marker(1, 0), marker(0, 1)],
      [marker(0, 0), marker(0, 0, 25_400, 12_699)],
      [marker(0, 0), marker(0, 0, 25_399, 12_700)],
    ] as const) {
      expect(
        (
          await capture({
            'xl/drawings/drawing1.xml': drawingDocument(
              twoCellMarkers(1, from, to),
            ),
          })
        ).diagnostic.message,
      ).toBe('Two-cell drawing anchor is reversed');
    }
  });

  it.each([
    ['A4', false],
    ['A0', false],
    ['D1', false],
    ['LEFT', false],
    ['A1', true],
    ['C3', true],
  ] as const)(
    'classifies two-cell selection boundary %s as %s',
    async (range, included) => {
      const selectionRange =
        range === 'A0' || range === 'LEFT' ? 'A1:A1' : range;
      const anchor =
        range === 'A0'
          ? twoCellAnchor(1, 0, 1, 2, 2)
          : range === 'LEFT'
            ? twoCellAnchor(1, 1, 0, 2, 2)
            : twoCellAnchor(1, 0, 0, 2, 2);
      const document = await parseXlsx(
        await bytes({
          'xl/drawings/drawing1.xml': drawingDocument(anchor),
        }),
        {
          errorMode: 'strict',
          selection: { ranges: { Sheet1: [selectionRange] } },
        },
      );
      const sheet = document.sheets[0]!;
      expect(
        sheet.kind === 'worksheet' ? sheet.drawings.length > 0 : false,
      ).toBe(included);
    },
  );

  it('round-trips drawing metadata through portable exact R0', async () => {
    const source = await bytes();
    const snapshot = await readXlsxRoundTrip(source);
    const output = await writeXlsxRoundTrip(
      await validateXlsxRoundTripJson(
        JSON.parse(JSON.stringify(snapshot)) as unknown,
      ),
    );
    expect(output.data).toEqual(source);
    expect(output.report.level).toBe('R0');
  });

  it('parses prefixed Strict drawing, relationship, and image ownership', async () => {
    const strictSheet = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictDrawing =
      'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing';
    const strictMain = 'http://purl.oclc.org/ooxml/drawingml/main';
    const strictRelationship =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const strictDrawingXml = drawingXml()
      .replaceAll(DRAWING_NS, strictDrawing)
      .replaceAll(DRAWING_MAIN_NS, strictMain)
      .replaceAll(XLSX_OFFICE_REL_TYPE.slice(0, -1), strictRelationship);
    const source = await createIndependentXlsx({
      '[Content_Types].xml': CONTENT_TYPES,
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelationship}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${strictRelationship}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/drawings/_rels/drawing1.xml.rels': DRAWING_RELS.replaceAll(
        XLSX_OFFICE_REL_TYPE,
        `${strictRelationship}/`,
      ),
      'xl/drawings/drawing1.xml': strictDrawingXml,
      'xl/media/image1.png': IMAGE_BYTES,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheet}" xmlns:r="${strictRelationship}"><s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets></s:workbook>`,
      'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replaceAll(
        XLSX_OFFICE_REL_TYPE,
        `${strictRelationship}/`,
      ),
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheet}" xmlns:r="${strictRelationship}"><s:sheetData/><s:drawing r:id="rIdDrawing"/></s:worksheet>`,
    });
    const document = await parseXlsx(source, { errorMode: 'strict' });
    const sheet = document.sheets[0]!;
    expect(sheet.kind === 'worksheet' ? sheet.drawings : []).toHaveLength(3);
  });

  it('parses an unprefixed drawing root with prefixed children', async () => {
    const xml = drawingXml()
      .replace('<xdr:wsDr ', `<wsDr xmlns="${DRAWING_NS}" `)
      .replace('</xdr:wsDr>', '</wsDr>');
    const document = await parseXlsx(
      await bytes({ 'xl/drawings/drawing1.xml': xml }),
      { errorMode: 'strict' },
    );
    expect(document.sheets).toHaveLength(1);
  });

  it.each([
    [
      { 'xl/worksheets/sheet1.xml': WORKSHEET.replace('rIdDrawing', '') },
      'Worksheet drawing relationship reference is invalid',
    ],
    [
      {
        'xl/worksheets/sheet1.xml': WORKSHEET.replace(
          '</worksheet>',
          '<drawing r:id="rIdDrawing"/></worksheet>',
        ),
      },
      'Worksheet contains duplicate drawing elements',
    ],
    [
      {
        'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replace(
          `Type="${DRAWING_RELATIONSHIP}"`,
          `Type="${XLSX_OFFICE_REL_TYPE}image"`,
        ),
      },
      'Worksheet drawing relationship is invalid',
    ],
    [
      {
        'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replace(
          'Target="../drawings/drawing1.xml"',
          'Target="https://example.invalid/drawing.xml" TargetMode="External"',
        ),
      },
      'Worksheet drawing relationship is invalid',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'application/vnd.openxmlformats-officedocument.drawing+xml',
          'application/xml',
        ),
      },
      'Worksheet drawing target has the wrong content type',
    ],
    [
      {
        'xl/drawings/drawing1.xml': `<wrong xmlns="${DRAWING_NS}"/>`,
      },
      'Worksheet drawing root is missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          DRAWING_NS,
          'urn:wrong',
        ),
      },
      'Worksheet drawing root has the wrong namespace',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'editAs="oneCell"',
          'editAs="bad"',
        ),
      },
      'Drawing anchor edit mode is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellAnchor(1, 0, 0).replace(
            '<xdr:oneCellAnchor>',
            '<xdr:oneCellAnchor editAs="oneCell">',
          ),
        ),
      },
      'Drawing anchor edit mode is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellAnchor(1, 0, 0).replace(
            '<xdr:ext cx="381000" cy="508000"/>',
            '<xdr:pos x="0" y="0"/><xdr:ext cx="381000" cy="508000"/>',
          ),
        ),
      },
      'Cell drawing anchor contains an absolute position',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          `<xdr:from>${marker(0, 0)}</xdr:from>`,
          `<xdr:from>${marker(3, 3)}</xdr:from>`,
        ),
      },
      'Two-cell drawing anchor is reversed',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'id="2" name="Image 2"',
          'id="1" name="Image 2"',
        ),
      },
      'Worksheet drawing contains duplicate object IDs',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'id="1" name="Image 1"',
          'id="0" name="Image 1"',
        ),
      },
      'Image ID is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'id="1" name="Image 1"',
          'id="-1" name="Image 1"',
        ),
      },
      'Image ID is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'name="Image 1"',
          'name=""',
        ),
      },
      'Image name is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'hidden="0"',
          'hidden="bad"',
        ),
      },
      'Image hidden flag is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'r:embed="image"',
          'r:link="external"',
        ),
      },
      'Externally linked drawing images are not loaded',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml('missing'),
      },
      'Image relationship is invalid',
    ],
    [
      {
        'xl/drawings/_rels/drawing1.xml.rels': DRAWING_RELS.replace(
          `Type="${IMAGE_RELATIONSHIP}"`,
          `Type="${XLSX_OFFICE_REL_TYPE}chart"`,
        ),
      },
      'Image relationship is invalid',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'ContentType="image/png"',
          'ContentType="image/svg+xml"',
        ),
      },
      'Drawing image content type is not safely supported',
    ],
    [
      { 'xl/media/image1.png': null },
      'Required XLSX image part is missing: xl/media/image1.png',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'l="1000" t="2000" r="3000" b="4000"',
          'l="60000" t="2000" r="40000" b="4000"',
        ),
      },
      'Image crop removes the complete image',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'flipH="1"',
          'flipH="bad"',
        ),
      },
      'Image horizontal-flip flag is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          '<xdr:pos x="-12700" y="-25400"/>',
          '',
        ),
      },
      'Absolute drawing position is missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'y="-25400"',
          'y="bad"',
        ),
      },
      'Drawing Y position is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'x="-12700"',
          'x="bad"',
        ),
      },
      'Drawing X position is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'x="-12700"',
          'x="x-12700"',
        ),
      },
      'Drawing X position is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'y="-25400"',
          'y="-25400x"',
        ),
      },
      'Drawing Y position is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'l="1000" t="2000" r="3000" b="4000"',
          'l="-100001" t="2000" r="3000" b="4000"',
        ),
      },
      'Image crop l is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(110)).replace('id="110"', 'id="-1"'),
        ),
      },
      'Drawing object ID is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(110)).replace(
            'descr="Shape description"',
            'hidden="bad"',
          ),
        ),
      },
      'Drawing object hidden flag is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(110)).replace(
            '<a:srgbClr val="aabbcc"/>',
            '<a:srgbClr val="xAABBCC"/>',
          ),
        ),
      },
      'Drawing color is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(110)).replace(
            '<a:srgbClr val="aabbcc"/>',
            '<a:sysClr val="window" lastClr="xAABBCC"/>',
          ),
        ),
      },
      'Drawing color is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(110)).replace(
            '<a:srgbClr val="aabbcc"/>',
            '<a:sysClr val="window" lastClr="AABBCCx"/>',
          ),
        ),
      },
      'Drawing color is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(140)).replace(
            '<a:solidFill><a:srgbClr val="aabbcc"/></a:solidFill>',
            '<a:solidFill>bad</a:solidFill>',
          ),
        ),
      },
      'Drawing color is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(110)).replace(
            '<a:srgbClr val="aabbcc"/>',
            '<a:srgbClr val="AABBCCx"/>',
          ),
        ),
      },
      'Drawing color is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(110)).replace('w="12700"', 'w="-1"'),
        ),
      },
      'Drawing line width is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(110)).replace(
            /<a:xfrm[^>]*>[\s\S]*?<\/a:xfrm>/u,
            '',
          ),
        ),
      },
      'Drawing object transform is missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(110)).replace('<a:off x="12700" y="25400"/>', ''),
        ),
      },
      'Drawing object transform is incomplete',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(110)).replace('prst="roundRect"', 'prst=""'),
        ),
      },
      'Drawing preset geometry is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(connector(120)).replace(
            'id="1" idx="2"',
            'id="-1" idx="2"',
          ),
        ),
      },
      'Drawing connector shape reference is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(group(130)).replace(
            'x="12700" y="25400"',
            'x="bad" y="25400"',
          ),
        ),
      },
      'Drawing group child X position is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(group(130)).replace(
            'x="12700" y="25400"',
            'x="12700" y="bad"',
          ),
        ),
      },
      'Drawing group child Y position is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(140)).replace(
            '<a:srgbClr val="aabbcc"/>',
            '<a:srgbClr val="AABBCC"/><a:schemeClr val="accent1"/>',
          ),
        ),
      },
      'Drawing color is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(140)).replace(
            '<a:srgbClr val="aabbcc"/>',
            '<a:foo val="AABBCC"/>',
          ),
        ),
      },
      'Drawing color is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(140)).replace('flipH="false"', 'flipH="bad"'),
        ),
      },
      'Drawing object horizontal-flip flag is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(140)).replace('flipV="true"', 'flipV="bad"'),
        ),
      },
      'Drawing object vertical-flip flag is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(140)).replace('rot="-60000"', 'rot="bad"'),
        ),
      },
      'Drawing object rotation is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(140)).replace('x="12700"', 'x="bad"'),
        ),
      },
      'Drawing object X position is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(140)).replace('y="25400"', 'y="bad"'),
        ),
      },
      'Drawing object Y position is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(140)).replace(
            /<a:p>[\s\S]*?<\/a:p><a:p>[\s\S]*?<\/a:p>/u,
            '',
          ),
        ),
      },
      'Drawing text is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(140)).replace(
            '<a:t>Shape &amp; text</a:t>',
            '<a:t><a:bad/></a:t>',
          ),
        ),
      },
      'Drawing text is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(connector(150)).replace('idx="2"', 'idx="-1"'),
        ),
      },
      'Drawing connector site is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(group(160)).replace(/<a:xfrm>[\s\S]*?<\/a:xfrm>/u, ''),
        ),
      },
      'Drawing group transform is missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(
            group(160).replace(
              /<xdr:sp>[\s\S]*?<\/xdr:sp>/u,
              '<xdr:sp>bad</xdr:sp>',
            ),
          ),
        ),
      },
      'Drawing group children are invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(70)).replace(
            /<xdr:nvSpPr>[\s\S]*?<\/xdr:nvSpPr>/u,
            '',
          ),
        ),
      },
      'Drawing object properties are missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(70)).replace('id="70"', 'id="0"'),
        ),
      },
      'Drawing object ID is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(70)).replace('name="Shape 70"', 'name=""'),
        ),
      },
      'Drawing object name is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(70)).replace(
            /<xdr:spPr>[\s\S]*?<\/xdr:spPr>/u,
            '',
          ),
        ),
      },
      'Drawing shape properties are missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(70)).replace(
            '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>',
            '',
          ),
        ),
      },
      'Drawing geometry is missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(70)).replace('val="aabbcc"', 'val="bad"'),
        ),
      },
      'Drawing color is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(connector(80)).replace('idx="2"', 'idx="bad"'),
        ),
      },
      'Drawing connector site is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(group(90)).replace('<a:chOff', '<a:wrong'),
        ),
      },
      'Drawing group child transform is missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellObject(shape(100)).replace(
            '</xdr:sp><xdr:clientData',
            `</xdr:sp>${connector(101)}<xdr:clientData`,
          ),
        ),
      },
      'Drawing anchor has multiple objects',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'l="1000" t="2000" r="3000" b="4000"',
          'l="100000" t="2000" r="0" b="4000"',
        ),
      },
      'Image crop removes the complete image',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'l="1000" t="2000" r="3000" b="4000"',
          'l="1000" t="60000" r="3000" b="40000"',
        ),
      },
      'Image crop removes the complete image',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          '<xdr:col>0</xdr:col>',
          '<xdr:col>x1</xdr:col>',
        ),
      },
      'Drawing anchor column is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          '<xdr:row>0</xdr:row>',
          '<xdr:row>1x</xdr:row>',
        ),
      },
      'Drawing anchor row is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          '<xdr:colOff>25400</xdr:colOff>',
          '<xdr:colOff>-1</xdr:colOff>',
        ),
      },
      'Drawing anchor column offset is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          '<xdr:rowOff>12700</xdr:rowOff>',
          '<xdr:rowOff>-1</xdr:rowOff>',
        ),
      },
      'Drawing anchor row offset is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          '<xdr:col>0</xdr:col>',
          '<xdr:col>-1</xdr:col>',
        ),
      },
      'Drawing anchor column is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          '<xdr:row>0</xdr:row>',
          '<xdr:row>-1</xdr:row>',
        ),
      },
      'Drawing anchor row is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          '<xdr:oneCellAnchor>bad</xdr:oneCellAnchor>',
        ),
      },
      'Worksheet drawing anchors are invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellAnchor(1, 0, 0).replace(
            `<xdr:from>${marker(0, 0)}</xdr:from>`,
            '',
          ),
        ),
      },
      'Drawing anchor marker is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          oneCellAnchor(1, 0, 0).replace(
            '<xdr:ext cx="381000" cy="508000"/>',
            '<xdr:wrong/>',
          ),
        ),
      },
      'Drawing extent is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace('l="1000"', 'l="bad"'),
      },
      'Image crop l is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'rot="60000"',
          'rot="bad"',
        ),
      },
      'Image rotation is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'flipV="0"',
          'flipV="bad"',
        ),
      },
      'Image vertical-flip flag is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          /<a:xfrm[^>]*>[\s\S]*?<\/a:xfrm>/u,
          '',
        ),
      },
      'Image transform is missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          '<xdr:cNvPr',
          '<xdr:wrong',
        ),
      },
      'Image properties are missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          /<xdr:blipFill>[\s\S]*?<\/xdr:blipFill>/u,
          '',
        ),
      },
      'Image fill is missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'r:embed="image"',
          'r:embed=""',
        ),
      },
      'Image relationship reference is invalid',
    ],
    [
      {
        'xl/drawings/_rels/drawing1.xml.rels': DRAWING_RELS.replace(
          'Target="../media/image1.png"',
          'Target="https://example.invalid/image.png" TargetMode="External"',
        ),
      },
      'Image relationship is invalid',
    ],
    [
      { 'xl/drawings/_rels/drawing1.xml.rels': null },
      'Image relationship is invalid',
    ],
    [
      { 'xl/drawings/drawing1.xml': null },
      'Required XLSX part is missing: xl/drawings/drawing1.xml',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'cx="381000"',
          'cx="-1"',
        ),
      },
      'Drawing width is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'cy="508000"',
          'cy="-1"',
        ),
      },
      'Drawing height is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          'l="1000" t="2000" r="3000" b="4000"',
          'l="100001" t="2000" r="3000" b="4000"',
        ),
      },
      'Image crop l is invalid',
    ],
  ] as const)(
    'rejects invalid drawing contract %#',
    async (overrides, message) => {
      expect((await capture(overrides)).diagnostic.message).toBe(message);
    },
  );

  it('enforces drawing and selection-work limits at exact boundaries', async () => {
    const exact = await parseXlsx(await bytes(), {
      errorMode: 'strict',
      limits: { maxDrawings: 3, maxReturnedCells: 2, maxScannedCells: 2 },
      selection: { ranges: { Sheet1: ['B2'] } },
    });
    expect(exact.sheets).toHaveLength(1);
    expect(
      (await capture({}, { errorMode: 'strict', limits: { maxDrawings: 2 } }))
        .diagnostic,
    ).toMatchObject({
      actual: 3,
      code: 'resource-limit-exceeded',
      limit: 2,
      limitName: 'maxDrawings',
    });
    expect(
      (
        await capture(
          {},
          {
            errorMode: 'strict',
            limits: { maxReturnedCells: 1, maxScannedCells: 1 },
            selection: { ranges: { Sheet1: ['B2'] } },
          },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      code: 'resource-limit-exceeded',
      limit: 1,
      limitName: 'maxScannedCells',
    });
    for (const [anchor, limits, limitName] of [
      [
        oneCellAnchor(1, 1, 0),
        { maxColumnsPerWorksheet: 1 },
        'maxColumnsPerWorksheet',
      ],
      [
        oneCellAnchor(1, 0, 1),
        { maxRowsPerWorksheet: 1 },
        'maxRowsPerWorksheet',
      ],
    ] as const) {
      expect(
        (
          await capture(
            { 'xl/drawings/drawing1.xml': drawingDocument(anchor) },
            { errorMode: 'strict', limits },
          )
        ).diagnostic,
      ).toMatchObject({
        actual: 2,
        code: 'resource-limit-exceeded',
        limit: 1,
        limitName,
      });
    }
  });
});

describe('XLSX image media session', () => {
  function reader(parts: Record<string, Uint8Array>): XlsxMediaReader {
    return {
      hasPart: (part) => parts[part] !== undefined,
      readBytes: (part) => Promise.resolve(parts[part]?.slice() ?? null),
    };
  }

  it.each([
    ['none', {}],
    ['base64', { base64: 'data:image/png;base64,AQID' }],
    ['blob', { blobUrl: 'blob:media' }],
    ['both', { base64: 'data:image/png;base64,AQID', blobUrl: 'blob:media' }],
  ] as const)('emits exact %s payload shape', async (mode, expected) => {
    vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:media');
    const session = new XlsxMediaSession(mode, defaultXlsxResourceLimits());
    const payload = await session.image(
      'xl/media/image.png',
      'image/png',
      reader({ 'xl/media/image.png': new Uint8Array([1, 2, 3]) }),
    );
    if (mode === 'none') expect(payload).toBeUndefined();
    else expect(payload).toStrictEqual({ byteLength: 3, ...expected });
  });

  it('rolls back only new cache and URLs and clears committed URLs once', async () => {
    const create = vi
      .spyOn(URL, 'createObjectURL')
      .mockReturnValueOnce('blob:a')
      .mockReturnValueOnce('blob:b')
      .mockReturnValueOnce('blob:b2');
    const revoke = vi
      .spyOn(URL, 'revokeObjectURL')
      .mockImplementation(() => {});
    const session = new XlsxMediaSession('blob', defaultXlsxResourceLimits());
    const mediaReader = reader({
      a: new Uint8Array([1]),
      b: new Uint8Array([2]),
    });
    await session.image('a', 'image/png', mediaReader);
    const checkpoint = session.checkpoint();
    await session.image('b', 'image/png', mediaReader);
    session.rollback(checkpoint);
    expect(revoke).toHaveBeenCalledExactlyOnceWith('blob:b');
    const retained = await session.image('a', 'image/png', mediaReader);
    const recreated = await session.image('b', 'image/png', mediaReader);
    expect(retained?.blobUrl).toBe('blob:a');
    expect(recreated?.blobUrl).toBe('blob:b2');
    expect(create).toHaveBeenCalledTimes(3);
    session.revokeAll();
    session.revokeAll();
    expect(revoke).toHaveBeenCalledTimes(3);
    expect(revoke).toHaveBeenCalledWith('blob:a');
    expect(revoke).toHaveBeenCalledWith('blob:b2');
  });

  it('validates metadata without reading excluded payloads', async () => {
    let reads = 0;
    const mediaReader: XlsxMediaReader = {
      hasPart: () => true,
      readBytes: () => {
        reads += 1;
        return Promise.resolve(IMAGE_BYTES);
      },
    };
    const session = new XlsxMediaSession('base64', defaultXlsxResourceLimits());
    expect(
      await session.image('image', 'image/png', mediaReader, false),
    ).toBeUndefined();
    expect(reads).toBe(0);
    await session.image('image', 'image/png', mediaReader);
    expect(reads).toBe(1);
  });

  it('reports a required image that disappears before its bounded read', async () => {
    const session = new XlsxMediaSession('base64', defaultXlsxResourceLimits());
    await expect(
      session.image('image', 'image/png', {
        hasPart: () => true,
        readBytes: () => Promise.resolve(null),
      }),
    ).rejects.toThrow('Required XLSX image part is missing: image');
  });

  it('creates Blob payloads with copied bytes and the declared content type', async () => {
    const create = vi
      .spyOn(URL, 'createObjectURL')
      .mockReturnValue('blob:typed');
    const session = new XlsxMediaSession('blob', defaultXlsxResourceLimits());
    await session.image(
      'image',
      'image/png',
      reader({ image: new Uint8Array([1, 2, 3]) }),
    );
    const blob = create.mock.calls[0]?.[0];
    expect(blob).toBeInstanceOf(Blob);
    if (!(blob instanceof Blob)) throw new Error('Expected Blob media payload');
    expect(blob.size).toBe(3);
    expect(blob.type).toBe('image/png');
  });
});

describe('XLSX drawing numeric normalization', () => {
  it.each([
    ['0', false, 0],
    ['4294967296', false, 4_294_967_296],
    ['-1', true, -1],
    ['-0', true, -0],
    ['9007199254740991', true, Number.MAX_SAFE_INTEGER],
  ] as const)('parses exact integer %s', (value, signed, expected) => {
    expect(parseXlsxDrawingInteger(value, signed, 'invalid', 'drawing')).toBe(
      expected,
    );
  });

  it.each([
    ['x1', false],
    ['1x', false],
    ['1a', false],
    ['-1', false],
    ['x-1', true],
    ['-1x', true],
    ['9007199254740992', true],
    ['01', false],
    ['1.0', false],
    ['+1', true],
    ['01', true],
    [' 1', true],
    ['1 ', true],
  ] as const)('rejects invalid integer %s', (value, signed) => {
    expect(() =>
      parseXlsxDrawingInteger(value, signed, 'invalid', 'drawing'),
    ).toThrow('invalid');
  });

  it.each([undefined, 1, null])('rejects non-string integer %#', (value) => {
    expect(() =>
      parseXlsxDrawingInteger(value, true, 'invalid', 'drawing'),
    ).toThrow('invalid');
  });
});
