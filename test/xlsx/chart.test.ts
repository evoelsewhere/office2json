import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  parseXlsxWithDiagnostics,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
  XlsxParseError,
} from '../../src/formats/xlsx';
import {
  parseXlsxChartFiniteNumber,
  parseXlsxChartUnsignedInteger,
} from '../../src/formats/xlsx/internal/chart';
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
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const OFFICE_REL_NS = XLSX_OFFICE_REL_TYPE.slice(0, -1);
const DRAWING_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}drawing`;
const CHART_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}chart`;

const CONTENT_TYPES = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`;

const WORKSHEET = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${OFFICE_REL_NS}"><sheetData/><drawing r:id="drawing"/></worksheet>`;
const WORKSHEET_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="drawing" Type="${DRAWING_RELATIONSHIP}" Target="../drawings/drawing1.xml"/></Relationships>`;
const DRAWING_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="chart" Type="${CHART_RELATIONSHIP}" Target="../charts/chart1.xml"/></Relationships>`;

function graphicFrame(id = 1, relationshipId = 'chart'): string {
  return `<xdr:graphicFrame>
    <xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}" descr="Chart description"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
    <xdr:xfrm rot="60000" flipH="1"><a:off x="-12700" y="25400"/><a:ext cx="127000" cy="254000"/></xdr:xfrm>
    <a:graphic><a:graphicData uri="${CHART_NS}"><c:chart xmlns:c="${CHART_NS}" r:id="${relationshipId}"/></a:graphicData></a:graphic>
  </xdr:graphicFrame>`;
}

function chartAnchor(frame: string): string {
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="381000" cy="508000"/>${frame}<xdr:clientData/></xdr:oneCellAnchor>`;
}

function drawingDocument(anchors: string): string {
  return `<xdr:wsDr xmlns:xdr="${DRAWING_NS}" xmlns:a="${DRAWING_MAIN_NS}" xmlns:c="${CHART_NS}" xmlns:r="${OFFICE_REL_NS}">${anchors}</xdr:wsDr>`;
}

function drawingXml(frame = graphicFrame()): string {
  return drawingDocument(chartAnchor(frame));
}

function stringSource(formula: string, values: readonly string[]): string {
  return `<c:strRef><c:f>${formula}</c:f><c:strCache><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`;
}

function numberSource(formula: string, values: readonly number[]): string {
  return `<c:numRef><c:f>${formula}</c:f><c:numCache><c:formatCode>0.00</c:formatCode><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`;
}

function chartXml(): string {
  return `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_MAIN_NS}">
    <c:roundedCorners val="1"/><c:style val="10"/>
    <c:chart>
      <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Sales &amp;</a:t></a:r><a:br/><a:r><a:t> Growth</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:autoTitleDeleted val="0"/>
      <c:plotArea><c:layout/>
        <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>
          <c:ser><c:idx val="0"/><c:order val="0"/>
            <c:tx>${stringSource('Sheet1!$B$1', ['Revenue'])}</c:tx>
            <c:spPr><a:solidFill><a:srgbClr val="abcdef"/></a:solidFill></c:spPr>
            <c:cat>${stringSource('Sheet1!$A$2:$A$3', ['Q1', 'Q2'])}</c:cat>
            <c:val>${numberSource('Sheet1!$B$2:$B$3', [12.5, 20])}</c:val>
            <c:marker><c:symbol val="circle"/><c:size val="7"/></c:marker><c:smooth val="false"/>
            <c:dLbls><c:dLblPos val="outEnd"/><c:showVal val="1"/></c:dLbls>
          </c:ser>
          <c:dLbls><c:showCatName val="true"/><c:showPercent val="0"/><c:separator>, </c:separator></c:dLbls>
          <c:gapWidth val="150"/><c:overlap val="-20"/><c:axId val="10"/><c:axId val="20"/>
        </c:barChart>
        <c:lineChart><c:grouping val="standard"/><c:varyColors/><c:axId val="10"/><c:axId val="20"/></c:lineChart>
        <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title><c:numFmt formatCode="General" sourceLinked="1"/><c:majorGridlines/><c:crossAx val="20"/><c:crosses val="autoZero"/></c:catAx>
        <c:valAx><c:axId val="20"/><c:scaling><c:orientation val="maxMin"/><c:min val="0"/><c:max val="100"/></c:scaling><c:delete val="false"/><c:axPos val="l"/><c:numFmt formatCode="0.0" sourceLinked="false"/><c:majorGridlines/><c:minorGridlines/><c:crossAx val="10"/><c:crossesAt val="0"/><c:majorUnit val="10"/><c:minorUnit val="2"/></c:valAx>
      </c:plotArea>
      <c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry><c:overlay val="true"/></c:legend>
      <c:plotVisOnly val="false"/><c:dispBlanksAs val="span"/><c:showDLblsOverMax val="1"/>
    </c:chart>
  </c:chartSpace>`;
}

function parts(overrides: XlsxBlackBoxOverrides = {}): XlsxBlackBoxOverrides {
  return {
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/charts/chart1.xml': chartXml(),
    'xl/drawings/_rels/drawing1.xml.rels': DRAWING_RELS,
    'xl/drawings/drawing1.xml': drawingXml(),
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
  throw new Error('Expected XLSX chart parsing to fail');
}

describe('XLSX charts', () => {
  it('parses chart frames, plots, series sources, caches, axes, titles, legends, colors, and styles', async () => {
    const document = await parseXlsx(await bytes(), { errorMode: 'strict' });
    const sheet = document.sheets[0]!;
    const object =
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
    expect(object).toMatchObject({
      autoTitleDeleted: false,
      description: 'Chart description',
      displayBlanksAs: 'span',
      hidden: false,
      id: 1,
      kind: 'chart',
      legend: {
        entries: [{ deleted: true, index: 0 }],
        overlay: true,
        position: 'right',
      },
      name: 'Chart 1',
      part: 'xl/charts/chart1.xml',
      plotVisibleOnly: false,
      roundedCorners: true,
      showDataLabelsOverMaximum: true,
      style: 10,
      title: { text: 'Sales &\n Growth' },
      transform: {
        flipHorizontal: true,
        flipVertical: false,
        height: 20,
        rotation: 1,
        width: 10,
        x: -1,
        y: 2,
      },
    });
    expect(object?.kind).toBe('chart');
    if (object?.kind !== 'chart') throw new Error('Expected chart object');
    expect(object.plots).toStrictEqual([
      {
        axisIds: [10, 20],
        barDirection: 'column',
        dataLabels: {
          separator: ', ',
          showBubbleSize: false,
          showCategoryName: true,
          showLegendKey: false,
          showPercent: false,
          showSeriesName: false,
          showValue: false,
        },
        gapWidth: 150,
        grouping: 'clustered',
        overlap: -20,
        series: [
          {
            categories: {
              formula: 'Sheet1!$A$2:$A$3',
              kind: 'string',
              pointCount: 2,
              points: [
                { index: 0, value: 'Q1' },
                { index: 1, value: 'Q2' },
              ],
            },
            color: { kind: 'rgb', value: 'ABCDEF' },
            dataLabels: {
              position: 'outEnd',
              showBubbleSize: false,
              showCategoryName: false,
              showLegendKey: false,
              showPercent: false,
              showSeriesName: false,
              showValue: true,
            },
            index: 0,
            marker: { size: 7, symbol: 'circle' },
            name: { formula: 'Sheet1!$B$1', text: 'Revenue' },
            order: 0,
            smooth: false,
            values: {
              formatCode: '0.00',
              formula: 'Sheet1!$B$2:$B$3',
              kind: 'number',
              pointCount: 2,
              points: [
                { index: 0, value: 12.5 },
                { index: 1, value: 20 },
              ],
            },
          },
        ],
        type: 'bar',
        varyColors: false,
      },
      {
        axisIds: [10, 20],
        grouping: 'standard',
        series: [],
        type: 'line',
        varyColors: true,
      },
    ]);
    expect(object.axes).toStrictEqual([
      {
        crossAxis: 20,
        crosses: 'autoZero',
        deleted: false,
        id: 10,
        kind: 'category',
        majorGridlines: true,
        minorGridlines: false,
        numberFormat: { code: 'General', sourceLinked: true },
        orientation: 'min-max',
        position: 'bottom',
        title: { text: 'Quarter' },
      },
      {
        crossAxis: 10,
        crossesAt: 0,
        deleted: false,
        id: 20,
        kind: 'value',
        majorGridlines: true,
        majorUnit: 10,
        maximum: 100,
        minimum: 0,
        minorGridlines: true,
        minorUnit: 2,
        numberFormat: { code: '0.0', sourceLinked: false },
        orientation: 'max-min',
        position: 'left',
      },
    ]);
    expect(JSON.parse(JSON.stringify(document))).toEqual(document);
  });

  it('recognizes every required common chart family in authored order', async () => {
    const families = [
      ['areaChart', 'area'],
      ['area3DChart', 'area-3d'],
      ['barChart', 'bar'],
      ['bar3DChart', 'bar-3d'],
      ['bubbleChart', 'bubble'],
      ['doughnutChart', 'doughnut'],
      ['lineChart', 'line'],
      ['line3DChart', 'line-3d'],
      ['ofPieChart', 'of-pie'],
      ['pieChart', 'pie'],
      ['pie3DChart', 'pie-3d'],
      ['radarChart', 'radar'],
      ['scatterChart', 'scatter'],
      ['stockChart', 'stock'],
      ['surfaceChart', 'surface'],
      ['surface3DChart', 'surface-3d'],
    ] as const;
    const plotXml = families
      .map(
        ([element]) => `<c:${element}><c:varyColors val="0"/></c:${element}>`,
      )
      .join('');
    const xml = `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea>${plotXml}</c:plotArea></c:chart></c:chartSpace>`;
    const document = await parseXlsx(
      await bytes({ 'xl/charts/chart1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const object =
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
    expect(object?.kind).toBe('chart');
    if (object?.kind !== 'chart') throw new Error('Expected chart object');
    expect(object.plots).toStrictEqual(
      families.map(([, type]) => ({
        axisIds: [],
        series: [],
        type,
        varyColors: false,
      })),
    );
    expect({
      axes: object.axes,
      autoTitleDeleted: object.autoTitleDeleted,
      displayBlanksAs: object.displayBlanksAs,
      legend: object.legend,
      plotVisibleOnly: object.plotVisibleOnly,
      roundedCorners: object.roundedCorners,
      showDataLabelsOverMaximum: object.showDataLabelsOverMaximum,
      style: object.style,
      title: object.title,
    }).toStrictEqual({
      axes: [],
      autoTitleDeleted: false,
      displayBlanksAs: 'gap',
      legend: undefined,
      plotVisibleOnly: true,
      roundedCorners: false,
      showDataLabelsOverMaximum: false,
      style: undefined,
      title: undefined,
    });
  });

  it('parses literal, multi-level, scatter, and bubble series sources', async () => {
    const xml = `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea>
      <c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Literal</c:v></c:tx><c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$1:$B$2</c:f><c:multiLvlStrCache><c:ptCount val="2"/><c:lvl><c:pt idx="0"><c:v>2025</c:v></c:pt><c:pt idx="1"><c:v>2026</c:v></c:pt></c:lvl><c:lvl><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat><c:val><c:numLit><c:formatCode>0</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart>
      <c:scatterChart><c:scatterStyle val="smoothMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:xVal><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:yVal></c:ser></c:scatterChart>
      <c:bubbleChart><c:bubbleScale val="75"/><c:ser><c:idx val="0"/><c:order val="0"/><c:bubbleSize><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>5</c:v></c:pt></c:numLit></c:bubbleSize></c:ser></c:bubbleChart>
    </c:plotArea></c:chart></c:chartSpace>`;
    const document = await parseXlsx(
      await bytes({ 'xl/charts/chart1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const object =
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
    expect(object?.kind).toBe('chart');
    if (object?.kind !== 'chart') throw new Error('Expected chart object');
    expect(object.plots).toStrictEqual([
      {
        axisIds: [],
        series: [
          {
            categories: {
              formula: 'Sheet1!$A$1:$B$2',
              kind: 'multi-level-string',
              levels: [
                [
                  { index: 0, value: '2025' },
                  { index: 1, value: '2026' },
                ],
                [
                  { index: 0, value: 'Q1' },
                  { index: 1, value: 'Q2' },
                ],
              ],
              pointCount: 2,
            },
            index: 0,
            name: { text: 'Literal' },
            order: 0,
            values: {
              formatCode: '0',
              kind: 'number',
              pointCount: 2,
              points: [
                { index: 0, value: 1 },
                { index: 1, value: 2 },
              ],
            },
          },
        ],
        type: 'bar',
        varyColors: false,
      },
      {
        axisIds: [],
        scatterStyle: 'smoothMarker',
        series: [
          {
            index: 0,
            order: 0,
            xValues: {
              kind: 'number',
              pointCount: 1,
              points: [{ index: 0, value: 3 }],
            },
            yValues: {
              kind: 'number',
              pointCount: 1,
              points: [{ index: 0, value: 4 }],
            },
          },
        ],
        type: 'scatter',
        varyColors: false,
      },
      {
        axisIds: [],
        bubbleScale: 75,
        series: [
          {
            bubbleSizes: {
              kind: 'number',
              pointCount: 1,
              points: [{ index: 0, value: 5 }],
            },
            index: 0,
            order: 0,
          },
        ],
        type: 'bubble',
        varyColors: false,
      },
    ]);
  });

  it('preserves every plot option, remaining axis kind, color kind, and mapped position', async () => {
    const xml = `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_MAIN_NS}"><c:chart><c:plotArea>
      <c:bar3DChart><c:barDir val="bar"/><c:gapDepth val="25"/><c:axId val="30"/><c:axId val="40"/></c:bar3DChart>
      <c:pieChart><c:firstSliceAng val="45"/></c:pieChart>
      <c:doughnutChart><c:holeSize val="55"/></c:doughnutChart>
      <c:radarChart><c:radarStyle val="filled"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr></c:ser></c:radarChart>
      <c:surfaceChart><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:sysClr val="windowText" lastClr="aabbcc"/></a:solidFill></c:spPr></c:ser></c:surfaceChart>
      <c:dateAx><c:axId val="30"/><c:scaling><c:logBase val="10"/></c:scaling><c:axPos val="r"/><c:crossAx val="40"/></c:dateAx>
      <c:serAx><c:axId val="40"/><c:scaling/><c:axPos val="t"/><c:crossAx val="30"/></c:serAx>
    </c:plotArea><c:legend><c:legendPos val="tr"/></c:legend></c:chart></c:chartSpace>`;
    const document = await parseXlsx(
      await bytes({ 'xl/charts/chart1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const object =
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
    expect(object?.kind).toBe('chart');
    if (object?.kind !== 'chart') throw new Error('Expected chart object');
    expect(object.plots).toStrictEqual([
      {
        axisIds: [30, 40],
        barDirection: 'bar',
        gapDepth: 25,
        series: [],
        type: 'bar-3d',
        varyColors: false,
      },
      {
        axisIds: [],
        firstSliceAngle: 45,
        series: [],
        type: 'pie',
        varyColors: false,
      },
      {
        axisIds: [],
        holeSize: 55,
        series: [],
        type: 'doughnut',
        varyColors: false,
      },
      {
        axisIds: [],
        radarStyle: 'filled',
        series: [
          {
            color: { kind: 'scheme', value: 'accent2' },
            index: 0,
            order: 0,
          },
        ],
        type: 'radar',
        varyColors: false,
      },
      {
        axisIds: [],
        series: [
          {
            color: {
              kind: 'system',
              lastColor: 'AABBCC',
              value: 'windowText',
            },
            index: 0,
            order: 0,
          },
        ],
        type: 'surface',
        varyColors: false,
      },
    ]);
    expect(object.axes).toStrictEqual([
      {
        crossAxis: 40,
        deleted: false,
        id: 30,
        kind: 'date',
        logBase: 10,
        majorGridlines: false,
        minorGridlines: false,
        orientation: 'min-max',
        position: 'right',
      },
      {
        crossAxis: 30,
        deleted: false,
        id: 40,
        kind: 'series',
        majorGridlines: false,
        minorGridlines: false,
        orientation: 'min-max',
        position: 'top',
      },
    ]);
    expect(object.legend).toStrictEqual({
      entries: [],
      overlay: false,
      position: 'top-right',
    });
  });

  it('parses string literals, system colors without fallbacks, and multiple title paragraphs', async () => {
    const xml = `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_MAIN_NS}"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>First</a:t></a:r></a:p><a:p><a:r><a:t>Second</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Literal name</c:v></c:pt></c:strLit></c:tx><c:spPr><a:solidFill><a:sysClr val="window"/></a:solidFill></c:spPr></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
    const document = await parseXlsx(
      await bytes({ 'xl/charts/chart1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const object =
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
    expect(object?.kind).toBe('chart');
    if (object?.kind !== 'chart') throw new Error('Expected chart object');
    expect(object.title).toStrictEqual({ text: 'First\nSecond' });
    expect(object.plots[0]?.series[0]).toStrictEqual({
      color: { kind: 'system', value: 'window' },
      index: 0,
      name: { text: 'Literal name' },
      order: 0,
    });
  });

  it.each([
    ['b', 'bottom'],
    ['l', 'left'],
    ['r', 'right'],
    ['t', 'top'],
    ['tr', 'top-right'],
  ] as const)('maps legend position %s', async (source, expected) => {
    const xml = `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:pieChart/></c:plotArea><c:legend><c:legendPos val="${source}"/></c:legend></c:chart></c:chartSpace>`;
    const document = await parseXlsx(
      await bytes({ 'xl/charts/chart1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const object =
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
    expect(object?.kind).toBe('chart');
    if (object?.kind !== 'chart') throw new Error('Expected chart object');
    expect(object.legend?.position).toBe(expected);
  });

  it('parses an unprefixed chart root with prefixed children', async () => {
    const xml = `<chartSpace xmlns="${CHART_NS}"><d:chart xmlns:d="${CHART_NS}"><d:plotArea><d:pieChart/></d:plotArea></d:chart></chartSpace>`;
    const document = await parseXlsx(
      await bytes({ 'xl/charts/chart1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object.kind : undefined,
    ).toBe('chart');
  });

  it('preserves optional chart omissions without inventing values', async () => {
    const xml = `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>A1</c:f><c:strCache><c:ptCount val="0"/></c:strCache></c:strRef></c:tx><c:spPr/><c:marker><c:symbol val="triangle"/></c:marker></c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart><c:catAx><c:axId val="1"/><c:scaling/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling/><c:crossAx val="1"/></c:valAx></c:plotArea><c:legend><c:legendEntry><c:idx val="2"/></c:legendEntry></c:legend></c:chart></c:chartSpace>`;
    const document = await parseXlsx(
      await bytes({ 'xl/charts/chart1.xml': xml }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    const object =
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
    expect(object?.kind).toBe('chart');
    if (object?.kind !== 'chart') throw new Error('Expected chart object');
    expect(object.plots[0]?.series[0]).toStrictEqual({
      index: 0,
      marker: { symbol: 'triangle' },
      name: { formula: 'A1', text: '' },
      order: 0,
    });
    expect(object.axes.map((axis) => axis.position)).toStrictEqual([
      undefined,
      undefined,
    ]);
    expect(object.axes[0]).not.toHaveProperty('position');
    expect(object.axes[1]).not.toHaveProperty('position');
    expect(object.legend).toStrictEqual({
      entries: [{ deleted: false, index: 2 }],
      overlay: false,
    });
  });

  it.each(['gap', 'span', 'zero'] as const)(
    'preserves chart blank display mode %s',
    async (mode) => {
      const xml = chartXml().replace(
        '<c:dispBlanksAs val="span"/>',
        `<c:dispBlanksAs val="${mode}"/>`,
      );
      const document = await parseXlsx(
        await bytes({ 'xl/charts/chart1.xml': xml }),
        { errorMode: 'strict' },
      );
      const sheet = document.sheets[0]!;
      const object =
        sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
      expect(object?.kind).toBe('chart');
      if (object?.kind !== 'chart') throw new Error('Expected chart object');
      expect(object.displayBlanksAs).toBe(mode);
    },
  );

  it.each([
    ['0', false],
    ['false', false],
    ['1', true],
    ['true', true],
  ] as const)(
    'parses chart number-format source-linked flag %s',
    async (source, expected) => {
      const xml = chartXml().replace(
        'sourceLinked="1"',
        `sourceLinked="${source}"`,
      );
      const document = await parseXlsx(
        await bytes({ 'xl/charts/chart1.xml': xml }),
        { errorMode: 'strict' },
      );
      const sheet = document.sheets[0]!;
      const object =
        sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
      expect(object?.kind).toBe('chart');
      if (object?.kind !== 'chart') throw new Error('Expected chart object');
      expect(object.axes[0]?.numberFormat?.sourceLinked).toBe(expected);
    },
  );

  it('parses Strict chart namespaces and owner relationships', async () => {
    const strictSheet = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictDrawing =
      'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing';
    const strictMain = 'http://purl.oclc.org/ooxml/drawingml/main';
    const strictChart = 'http://purl.oclc.org/ooxml/drawingml/chart';
    const strictRelationship =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const source = await createIndependentXlsx({
      '[Content_Types].xml': CONTENT_TYPES,
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelationship}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${strictRelationship}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/charts/chart1.xml': chartXml()
        .replaceAll(CHART_NS, strictChart)
        .replaceAll(DRAWING_MAIN_NS, strictMain),
      'xl/drawings/_rels/drawing1.xml.rels': DRAWING_RELS.replaceAll(
        XLSX_OFFICE_REL_TYPE,
        `${strictRelationship}/`,
      ),
      'xl/drawings/drawing1.xml': drawingXml()
        .replaceAll(DRAWING_NS, strictDrawing)
        .replaceAll(DRAWING_MAIN_NS, strictMain)
        .replaceAll(CHART_NS, strictChart)
        .replaceAll(OFFICE_REL_NS, strictRelationship),
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheet}" xmlns:r="${strictRelationship}"><s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets></s:workbook>`,
      'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replaceAll(
        XLSX_OFFICE_REL_TYPE,
        `${strictRelationship}/`,
      ),
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheet}" xmlns:r="${strictRelationship}"><s:sheetData/><s:drawing r:id="drawing"/></s:worksheet>`,
    });
    const document = await parseXlsx(source, { errorMode: 'strict' });
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.drawings[0]?.object.kind : undefined,
    ).toBe('chart');
  });

  it('round-trips chart metadata through portable exact R0', async () => {
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

  it('validates but omits a chart outside selected cell anchors', async () => {
    const options = {
      errorMode: 'strict' as const,
      selection: { ranges: { Sheet1: ['B2'] } },
    };
    const document = await parseXlsx(await bytes(), options);
    const sheet = document.sheets[0]!;
    expect(sheet.kind === 'worksheet' ? sheet.drawings : []).toEqual([]);
    await expect(
      parseXlsx(await bytes({ 'xl/charts/chart1.xml': '<broken' }), options),
    ).rejects.toBeInstanceOf(XlsxParseError);
  });

  it('enforces maxCharts exactly across drawing parts', async () => {
    await expect(
      parseXlsx(await bytes(), {
        errorMode: 'strict',
        limits: { maxCharts: 1 },
      }),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(
          {
            'xl/drawings/drawing1.xml': drawingDocument(
              `${chartAnchor(graphicFrame(1))}${chartAnchor(graphicFrame(2))}`,
            ),
          },
          { errorMode: 'strict', limits: { maxCharts: 1 } },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      code: 'resource-limit-exceeded',
      limit: 1,
      limitName: 'maxCharts',
      part: 'xl/charts/chart1.xml',
    });
  });

  it('skips an empty drawing anchor and parses a chart nested in a group', async () => {
    const empty = await parseXlsx(
      await bytes({
        'xl/drawings/drawing1.xml': drawingDocument(chartAnchor('')),
      }),
      { errorMode: 'strict' },
    );
    const emptySheet = empty.sheets[0]!;
    expect(emptySheet.kind === 'worksheet' ? emptySheet.drawings : []).toEqual(
      [],
    );

    const group = `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="10" name="Chart group"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm></xdr:grpSpPr>${graphicFrame(11)}</xdr:grpSp>`;
    const nested = await parseXlsx(
      await bytes({
        'xl/drawings/drawing1.xml': drawingDocument(chartAnchor(group)),
      }),
      { errorMode: 'strict' },
    );
    const nestedSheet = nested.sheets[0]!;
    const object =
      nestedSheet.kind === 'worksheet'
        ? nestedSheet.drawings[0]?.object
        : undefined;
    expect(object?.kind).toBe('group');
    if (object?.kind !== 'group') throw new Error('Expected drawing group');
    expect(object.children.map((child) => child.kind)).toStrictEqual(['chart']);
  });

  it('enforces chart formula and returned-text budgets at exact boundaries', async () => {
    const formulaChart = (formula: string) =>
      `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>${formula}</c:f><c:strCache><c:ptCount val="0"/></c:strCache></c:strRef></c:cat></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
    await expect(
      parseXlsx(await bytes({ 'xl/charts/chart1.xml': formulaChart('A') }), {
        errorMode: 'strict',
        limits: { maxFormulaCharacters: 1 },
      }),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(
          { 'xl/charts/chart1.xml': formulaChart('AA') },
          { errorMode: 'strict', limits: { maxFormulaCharacters: 1 } },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxFormulaCharacters',
      part: 'xl/charts/chart1.xml',
    });

    const textChart = `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_MAIN_NS}"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>X</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart/></c:plotArea></c:chart></c:chartSpace>`;
    await expect(
      parseXlsx(await bytes({ 'xl/charts/chart1.xml': textChart }), {
        errorMode: 'strict',
        limits: { maxTextCharacters: 10 },
      }),
    ).resolves.toBeDefined();
    expect(
      (
        await capture(
          { 'xl/charts/chart1.xml': textChart },
          { errorMode: 'strict', limits: { maxTextCharacters: 9 } },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 10,
      limit: 9,
      limitName: 'maxTextCharacters',
      part: 'xl/charts/chart1.xml',
    });
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
  ] as const)(
    'parses chart numeric lexical form %s',
    async (source, expected) => {
      const xml = chartXml().replace('<c:v>12.5</c:v>', `<c:v>${source}</c:v>`);
      const document = await parseXlsx(
        await bytes({ 'xl/charts/chart1.xml': xml }),
        { errorMode: 'strict' },
      );
      const sheet = document.sheets[0]!;
      const object =
        sheet.kind === 'worksheet' ? sheet.drawings[0]?.object : undefined;
      expect(object?.kind).toBe('chart');
      if (object?.kind !== 'chart') throw new Error('Expected chart object');
      const sourceData = object.plots[0]?.series[0]?.values;
      expect(
        sourceData?.kind === 'number' ? sourceData.points[0]?.value : null,
      ).toBe(expected);
    },
  );

  it.each([
    'x1',
    '1x',
    '1a2',
    '.',
    '+',
    '1e',
    '1e+',
    '1e-',
    '1a3',
    ' 1',
    '1 ',
    '1e309',
  ])('rejects chart numeric lexical form %s', async (source) => {
    const xml = chartXml().replace('<c:v>12.5</c:v>', `<c:v>${source}</c:v>`);
    expect(
      (await capture({ 'xl/charts/chart1.xml': xml })).diagnostic,
    ).toMatchObject({ message: 'Chart numeric cache value is invalid' });
  });

  it.each(['-1', '01', '1.0', 'x1', '1x', '9007199254740992'])(
    'rejects chart unsigned integer lexical form %s',
    async (source) => {
      const xml = chartXml().replace(
        '<c:idx val="0"/>',
        `<c:idx val="${source}"/>`,
      );
      expect(
        (await capture({ 'xl/charts/chart1.xml': xml })).diagnostic,
      ).toMatchObject({ message: 'Chart series index is invalid' });
    },
  );

  it.each([undefined, null, 1, [], {}])(
    'rejects non-string chart numeric value %#',
    (source) => {
      expect(() =>
        parseXlsxChartFiniteNumber(
          source,
          'Chart test number is invalid',
          'xl/charts/chart1.xml',
        ),
      ).toThrow('Chart test number is invalid');
    },
  );

  it.each([undefined, null, 1, [], {}])(
    'rejects non-string chart unsigned integer %#',
    (source) => {
      expect(() =>
        parseXlsxChartUnsignedInteger(
          source,
          'Chart test integer is invalid',
          'xl/charts/chart1.xml',
        ),
      ).toThrow('Chart test integer is invalid');
    },
  );

  it('recovers an invalid optional chart in tolerant mode and rejects it in strict mode', async () => {
    const overrides = { 'xl/charts/chart1.xml': '<broken' };
    const tolerant = await parseXlsxWithDiagnostics(await bytes(overrides));
    const sheet = tolerant.document.sheets[0]!;
    expect(sheet.kind === 'worksheet' ? sheet.drawings : []).toEqual([]);
    expect(tolerant.diagnostics).toMatchObject([
      {
        code: 'xml-parse-failed',
        part: 'xl/charts/chart1.xml',
        severity: 'warning',
      },
    ]);
    await expect(
      parseXlsx(await bytes(overrides), { errorMode: 'strict' }),
    ).rejects.toBeInstanceOf(XlsxParseError);
  });

  it.each([
    [
      {
        'xl/drawings/_rels/drawing1.xml.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
      },
      'Chart relationship is invalid',
    ],
    [
      { 'xl/drawings/drawing1.xml': drawingXml(graphicFrame(1, 'missing')) },
      'Chart relationship is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml().replace(
          `uri="${CHART_NS}"`,
          'uri="bad"',
        ),
      },
      'Chart graphic data URI is invalid',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
          'application/xml',
        ),
      },
      'Chart target has the wrong content type',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:axId val="20"/>',
          '<c:axId val="99"/>',
        ),
      },
      'Chart plot references a missing axis',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:pt idx="1"><c:v>Q2</c:v></c:pt>',
          '<c:pt idx="0"><c:v>Q2</c:v></c:pt>',
        ),
      },
      'Chart cache contains a duplicate point index',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:v>12.5</c:v>',
          '<c:v>NaN</c:v>',
        ),
      },
      'Chart numeric cache value is invalid',
    ],
    [
      {
        'xl/drawings/_rels/drawing1.xml.rels': DRAWING_RELS.replace(
          'Target="../charts/chart1.xml"',
          'Target="https://example.invalid/chart.xml" TargetMode="External"',
        ),
      },
      'Externally linked charts are not loaded',
    ],
    [
      {
        'xl/charts/chart1.xml': `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:funnelChart/></c:plotArea></c:chart></c:chartSpace>`,
      },
      'Chart family is not supported',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:dLbls><c:showCatName',
          '<c:dLbls>bad</c:dLbls><c:dLbls><c:showCatName',
        ),
      },
      'Chart data labels are invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt>',
          '<c:ptCount val="1"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt>',
        ),
      },
      'Chart cache point exceeds its declared count',
    ],
    [
      { 'xl/charts/chart1.xml': `<wrong xmlns="${CHART_NS}"/>` },
      'Chart root is missing',
    ],
    [
      { 'xl/charts/chart1.xml': chartXml().replace(CHART_NS, 'urn:wrong') },
      'Chart root has the wrong namespace',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml()
          .replace('<c:chart>', '<c:notChart>')
          .replace('</c:chart>', '</c:notChart>'),
      },
      'Chart definition is missing',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml()
          .replace('<c:plotArea>', '<c:notPlotArea>')
          .replace('</c:plotArea>', '</c:notPlotArea>'),
      },
      'Chart plot area is missing',
    ],
    [
      {
        'xl/charts/chart1.xml': `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:barChart>bad</c:barChart></c:plotArea></c:chart></c:chartSpace>`,
      },
      'Chart plot area is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:layout/></c:plotArea></c:chart></c:chartSpace>`,
      },
      'Chart contains no supported plot',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '</c:plotArea>',
          '<c:catAx><c:axId val="10"/><c:scaling/></c:catAx></c:plotArea>',
        ),
      },
      'Chart contains a duplicate axis ID',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:crossAx val="20"/>',
          '<c:crossAx val="99"/>',
        ),
      },
      'Chart axis references a missing cross axis',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:dispBlanksAs val="span"/>',
          '<c:dispBlanksAs val="bad"/>',
        ),
      },
      'Chart blank-display mode is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:style val="10"/>',
          '<c:style val="bad"/>',
        ),
      },
      'Chart style is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:showVal val="1"/>',
          '<c:showVal val="bad"/>',
        ),
      },
      'Chart data-label value flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:showCatName val="true"/>',
          '<c:showCatName val="bad"/>',
        ),
      },
      'Chart data-label category-name flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:showPercent val="0"/>',
          '<c:showPercent val="bad"/>',
        ),
      },
      'Chart data-label percent flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:showCatName val="true"/>',
          '<c:showBubbleSize val="bad"/><c:showCatName val="true"/>',
        ),
      },
      'Chart data-label bubble-size flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:showCatName val="true"/>',
          '<c:showLegendKey val="bad"/><c:showCatName val="true"/>',
        ),
      },
      'Chart data-label legend-key flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:showCatName val="true"/>',
          '<c:showSerName val="bad"/><c:showCatName val="true"/>',
        ),
      },
      'Chart data-label series-name flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:smooth val="false"/>',
          '<c:smooth val="bad"/>',
        ),
      },
      'Chart series smooth flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:varyColors val="0"/>',
          '<c:varyColors val="bad"/>',
        ),
      },
      'Chart vary-colors flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:delete val="0"/>',
          '<c:delete val="bad"/>',
        ),
      },
      'Chart axis delete flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:delete val="1"/>',
          '<c:delete val="bad"/>',
        ),
      },
      'Chart legend entry delete flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:overlay val="true"/>',
          '<c:overlay val="bad"/>',
        ),
      },
      'Chart legend overlay flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:autoTitleDeleted val="0"/>',
          '<c:autoTitleDeleted val="bad"/>',
        ),
      },
      'Chart auto-title-delete flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:plotVisOnly val="false"/>',
          '<c:plotVisOnly val="bad"/>',
        ),
      },
      'Chart visible-only flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:roundedCorners val="1"/>',
          '<c:roundedCorners val="bad"/>',
        ),
      },
      'Chart rounded-corners flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:showDLblsOverMax val="1"/>',
          '<c:showDLblsOverMax val="bad"/>',
        ),
      },
      'Chart labels-over-maximum flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:showVal val="1"/>',
          '<c:showVal>bad</c:showVal>',
        ),
      },
      'Chart data-label value flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:spPr><a:solidFill><a:srgbClr val="abcdef"/></a:solidFill></c:spPr>',
          '<c:spPr>bad</c:spPr>',
        ),
      },
      'Chart shape properties are invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<a:solidFill><a:srgbClr val="abcdef"/></a:solidFill>',
          '<a:solidFill>bad</a:solidFill>',
        ),
      },
      'Chart color is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<a:srgbClr val="abcdef"/>',
          '<a:srgbClr val="xABCDEF"/>',
        ),
      },
      'Chart color is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<a:srgbClr val="abcdef"/>',
          '<a:srgbClr val="ABCDEFx"/>',
        ),
      },
      'Chart color is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:cat><c:strRef>',
          '<c:cat>bad</c:cat><c:cat><c:strRef>',
        ),
      },
      'Chart data source is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '</c:strRef></c:cat>',
          '</c:strRef><c:numLit/></c:cat>',
        ),
      },
      'Chart data source is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:f>Sheet1!$A$2:$A$3</c:f>',
          '<c:f><c:x/></c:f>',
        ),
      },
      'Chart data source formula is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml()
          .replace('<c:strCache>', '<c:wrong>')
          .replace('</c:strCache>', '</c:wrong>'),
      },
      'Chart data cache is missing',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:formatCode>0.00</c:formatCode>',
          '<c:formatCode><c:x/></c:formatCode>',
        ),
      },
      'Chart cache format code is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:ptCount val="2"/>',
          '<c:ptCount val="bad"/>',
        ),
      },
      'Chart cache point count is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:pt idx="0"><c:v>Revenue</c:v></c:pt>',
          '<c:pt>bad</c:pt>',
        ),
      },
      'Chart cache points are invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:pt idx="0"><c:v>Revenue</c:v></c:pt>',
          '<c:pt idx="bad"><c:v>Revenue</c:v></c:pt>',
        ),
      },
      'Chart cache point index is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:pt idx="0"><c:v>Revenue</c:v></c:pt>',
          '<c:pt idx="0"><c:v><c:x/></c:v></c:pt>',
        ),
      },
      'Chart cache point value is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace('<c:idx val="0"/>', ''),
      },
      'Chart series identity is missing',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace('<c:order val="0"/>', ''),
      },
      'Chart series identity is missing',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:marker><c:symbol val="circle"/><c:size val="7"/></c:marker>',
          '<c:marker>bad</c:marker>',
        ),
      },
      'Chart marker is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:size val="7"/>',
          '<c:size val="bad"/>',
        ),
      },
      'Chart marker size is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:symbol val="circle"/>',
          '<c:symbol/>',
        ),
      },
      'Chart marker symbol is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '</c:ser>',
          '</c:ser><c:ser><c:idx val="0"/><c:order val="1"/></c:ser>',
        ),
      },
      'Chart plot contains duplicate series identity',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '</c:ser>',
          '</c:ser><c:ser><c:idx val="1"/><c:order val="0"/></c:ser>',
        ),
      },
      'Chart plot contains duplicate series identity',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:barDir val="col"/>',
          '<c:barDir val="bad"/>',
        ),
      },
      'Chart bar direction is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:gapWidth val="150"/>',
          '<c:gapWidth val="bad"/>',
        ),
      },
      'Chart gap width is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:overlap val="-20"/>',
          '<c:overlap val="bad"/>',
        ),
      },
      'Chart overlap is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:grouping val="clustered"/>',
          '<c:grouping/>',
        ),
      },
      'Chart grouping is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:axId val="10"/>',
          '<c:axId val="bad"/>',
        ),
      },
      'Chart plot axis reference is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:catAx><c:axId val="10"/>',
          '<c:catAx>',
        ),
      },
      'Chart axis ID is missing',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:scaling><c:orientation val="minMax"/></c:scaling>',
          '<c:wrong/>',
        ),
      },
      'Chart axis scaling is missing',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:orientation val="minMax"/>',
          '<c:orientation val="bad"/>',
        ),
      },
      'Chart axis orientation is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:axPos val="b"/>',
          '<c:axPos val="bad"/>',
        ),
      },
      'Chart axis position is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:numFmt formatCode="General" sourceLinked="1"/>',
          '<c:numFmt sourceLinked="1"/>',
        ),
      },
      'Chart axis number format is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          'sourceLinked="1"',
          'sourceLinked="bad"',
        ),
      },
      'Chart axis number-format source flag is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:crossAx val="20"/>',
          '<c:crossAx val="bad"/>',
        ),
      },
      'Chart cross-axis reference is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:crossesAt val="0"/>',
          '<c:crossesAt val="bad"/>',
        ),
      },
      'Chart axis crossing value is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:majorUnit val="10"/>',
          '<c:majorUnit val="bad"/>',
        ),
      },
      'Chart axis major unit is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:max val="100"/>',
          '<c:max val="bad"/>',
        ),
      },
      'Chart axis maximum is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:min val="0"/>',
          '<c:min val="bad"/>',
        ),
      },
      'Chart axis minimum is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:minorUnit val="2"/>',
          '<c:minorUnit val="bad"/>',
        ),
      },
      'Chart axis minor unit is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:legendPos val="r"/>',
          '<c:legendPos val="bad"/>',
        ),
      },
      'Chart legend position is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry>',
          '<c:legendEntry>bad</c:legendEntry>',
        ),
      },
      'Chart legend entries are invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:legendEntry><c:idx val="0"/>',
          '<c:legendEntry>',
        ),
      },
      'Chart legend entry index is missing',
    ],
    [
      {
        'xl/charts/chart1.xml': `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:barChart><c:ser>bad</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`,
      },
      'Chart plot structure is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:gapWidth val="150"/>',
          '<c:bubbleScale val="bad"/><c:gapWidth val="150"/>',
        ),
      },
      'Chart bubble scale is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:gapWidth val="150"/>',
          '<c:firstSliceAng val="bad"/><c:gapWidth val="150"/>',
        ),
      },
      'Chart first-slice angle is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:gapWidth val="150"/>',
          '<c:gapDepth val="bad"/><c:gapWidth val="150"/>',
        ),
      },
      'Chart gap depth is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:gapWidth val="150"/>',
          '<c:holeSize val="bad"/><c:gapWidth val="150"/>',
        ),
      },
      'Chart hole size is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:grouping val="clustered"/>',
          '<c:radarStyle/><c:grouping val="clustered"/>',
        ),
      },
      'Chart radar style is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:grouping val="clustered"/>',
          '<c:scatterStyle/><c:grouping val="clustered"/>',
        ),
      },
      'Chart scatter style is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:catAx><c:axId val="10"/>',
          '<c:catAx><c:axId val="bad"/>',
        ),
      },
      'Chart axis ID is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:orientation val="minMax"/>',
          '<c:logBase val="bad"/><c:orientation val="minMax"/>',
        ),
      },
      'Chart axis logarithm base is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:crosses val="autoZero"/>',
          '<c:crosses/>',
        ),
      },
      'Chart axis crossing is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:legend><c:legendPos',
          '<c:legend>bad</c:legend><c:legend><c:legendPos',
        ),
      },
      'Chart legend is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:title><c:tx><c:rich>',
          '<c:title>bad</c:title><c:title><c:tx><c:rich>',
        ),
      },
      'Chart text is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:title><c:tx><c:numLit/></c:tx></c:title><c:plotArea><c:pieChart/></c:plotArea></c:chart></c:chartSpace>`,
      },
      'Chart text source is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:ptCount val="1"/>',
          '<c:ptCount/>',
        ),
      },
      'Chart cache point count is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:multiLvlStrRef><c:f>A1</c:f></c:multiLvlStrRef></c:cat></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`,
      },
      'Chart multi-level cache is missing',
    ],
    [
      {
        'xl/charts/chart1.xml': `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:multiLvlStrRef><c:f>A1</c:f><c:multiLvlStrCache><c:lvl>bad</c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`,
      },
      'Chart multi-level cache is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<a:srgbClr val="abcdef"/>',
          '<a:sysClr val="window" lastClr="xAABBCC"/>',
        ),
      },
      'Chart color is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<a:srgbClr val="abcdef"/>',
          '<a:sysClr val="window" lastClr="AABBCCx"/>',
        ),
      },
      'Chart color is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_MAIN_NS}"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>bad</a:p></c:rich></c:tx></c:title><c:plotArea><c:pieChart/></c:plotArea></c:chart></c:chartSpace>`,
      },
      'Chart text is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<a:t>Sales &amp;</a:t>',
          '<a:t><a:x/></a:t>',
        ),
      },
      'Chart text is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:separator>, </c:separator>',
          '<c:separator><c:x/></c:separator>',
        ),
      },
      'Chart data-label separator is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:dLblPos val="outEnd"/>',
          '<c:dLblPos/>',
        ),
      },
      'Chart data-label position is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:idx val="0"/>',
          '<c:idx/>',
        ),
      },
      'Chart series index is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:order val="0"/>',
          '<c:order/>',
        ),
      },
      'Chart series order is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:size val="7"/>',
          '<c:size/>',
        ),
      },
      'Chart marker size is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:catAx><c:axId val="10"/>',
          '<c:catAx><c:axId/>',
        ),
      },
      'Chart axis ID is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:orientation val="minMax"/>',
          '<c:orientation/>',
        ),
      },
      'Chart axis orientation is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:axPos val="b"/>',
          '<c:axPos/>',
        ),
      },
      'Chart axis position is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:crossAx val="20"/>',
          '<c:crossAx/>',
        ),
      },
      'Chart cross-axis reference is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:legendPos val="r"/>',
          '<c:legendPos/>',
        ),
      },
      'Chart legend position is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:legendEntry><c:idx val="0"/>',
          '<c:legendEntry><c:idx/>',
        ),
      },
      'Chart legend entry index is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:style val="10"/>',
          '<c:style/>',
        ),
      },
      'Chart style is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<a:srgbClr val="abcdef"/>',
          '<a:srgbClr val="abcdef"/><a:schemeClr val="accent1"/>',
        ),
      },
      'Chart color is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<a:srgbClr val="abcdef"/>',
          '<a:srgbClr>bad</a:srgbClr>',
        ),
      },
      'Chart color is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<a:srgbClr val="abcdef"/>',
          '<a:hslClr val="abcdef"/>',
        ),
      },
      'Chart color is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v><c:x/></c:v></c:tx></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`,
      },
      'Chart text is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:order val="0"/>',
          '<c:order val="bad"/>',
        ),
      },
      'Chart series order is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:barDir val="col"/>',
          '<c:barDir/>',
        ),
      },
      'Chart bar direction is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:numFmt formatCode="General" sourceLinked="1"/>',
          '<c:numFmt>bad</c:numFmt>',
        ),
      },
      'Chart axis number format is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:legendEntry><c:idx val="0"/>',
          '<c:legendEntry><c:idx val="bad"/>',
        ),
      },
      'Chart legend entry index is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          '<c:dispBlanksAs val="span"/>',
          '<c:dispBlanksAs/>',
        ),
      },
      'Chart blank-display mode is invalid',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          `xmlns:c="${CHART_NS}"`,
          `xmlns:c="${CHART_NS}" xmlns:duplicate="${CHART_NS}"`,
        ),
      },
      'Chart root has the wrong namespace',
    ],
    [
      {
        'xl/charts/chart1.xml': chartXml().replace(
          CHART_NS,
          'http://purl.oclc.org/ooxml/drawingml/chart',
        ),
      },
      'Chart root has the wrong namespace',
    ],
    [
      {
        'xl/charts/chart1.xml': `<x:chartSpace xmlns:x="urn:wrong" xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:pieChart/></c:plotArea></c:chart></x:chartSpace>`,
      },
      'Chart root has the wrong namespace',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml(
          graphicFrame().replace(/<a:graphic>[\s\S]*<\/a:graphic>/u, ''),
        ),
      },
      'Chart graphic data is missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml(
          graphicFrame().replace(
            `<c:chart xmlns:c="${CHART_NS}" r:id="chart"/>`,
            '<c:wrong/>',
          ),
        ),
      },
      'Chart reference is missing',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingXml(
          graphicFrame().replace(' r:id="chart"', ''),
        ),
      },
      'Chart relationship reference is invalid',
    ],
    [
      {
        'xl/drawings/_rels/drawing1.xml.rels': DRAWING_RELS.replace(
          CHART_RELATIONSHIP,
          `${XLSX_OFFICE_REL_TYPE}image`,
        ),
      },
      'Chart relationship is invalid',
    ],
    [
      {
        'xl/drawings/drawing1.xml': drawingDocument(
          chartAnchor('<xdr:graphicFrame>bad</xdr:graphicFrame>'),
        ),
      },
      'Drawing anchor objects are invalid',
    ],
  ] as const)(
    'rejects invalid chart contract %#',
    async (overrides, message) => {
      expect(
        (await capture(overrides)).diagnostic.message,
        JSON.stringify(overrides),
      ).toBe(message);
    },
  );
});
