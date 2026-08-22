import { describe, expect, it } from 'vitest';

import type {
  Image,
  Group,
  Chart,
  CommonChart,
  PptxDocument,
  PptxParseOptions,
  Shape,
  Table,
  Text,
} from '../../src/formats/pptx/types';
import type { PptxSvgRenderResult } from '../../src/formats/pptx/render-types';
import type {
  PptxSceneDocument,
  PptxSceneSlide,
} from '../../src/formats/pptx/scene-types';
import {
  verifyPowerPointCreation,
  verifyPowerPointCreationWithParser,
} from '../../src/formats/pptx/writer/verify';

function emptySlide(key: string): PptxSceneSlide {
  return { elements: [], key };
}

function textSlide(key: string): PptxSceneSlide {
  return {
    elements: [
      {
        authored: {
          transform: { height: 40, width: 160, x: 10, y: 20 },
        },
        key: `${key}-text`,
        resolved: { hidden: false },
        text: {
          body: {},
          paragraphs: [
            {
              children: [{ key: `${key}-run`, text: 'Text', type: 'run' }],
              key: `${key}-paragraph`,
            },
          ],
        },
        type: 'text',
      },
    ],
    key,
  };
}

function shapeSlide(key: string): PptxSceneSlide {
  return {
    elements: [
      {
        authored: {
          fillColor: '#F97316',
          geometry: 'ellipse',
          lineColor: '#0F172A',
          lineWidth: 2,
          transform: { height: 40, width: 160, x: 10, y: 20 },
        },
        key: `${key}-shape`,
        resolved: { hidden: false },
        type: 'shape',
      },
    ],
    key,
  };
}

const IMAGE_BYTES = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
]);

function imageSlide(key: string): PptxSceneSlide {
  return {
    elements: [
      {
        authored: {
          transform: { height: 40, width: 160, x: 10, y: 20 },
        },
        key: `${key}-image`,
        mediaKey: 'media-1',
        resolved: { hidden: false },
        type: 'image',
      },
    ],
    key,
  };
}

function tableSlide(key: string): PptxSceneSlide {
  return {
    elements: [
      {
        authored: {
          transform: { height: 40, width: 160, x: 10, y: 20 },
        },
        columns: [160],
        key: `${key}-table`,
        resolved: { hidden: false },
        rows: [
          {
            cells: [
              {
                text: {
                  body: {},
                  paragraphs: [
                    {
                      children: [
                        { key: `${key}-run`, text: 'Expected', type: 'run' },
                      ],
                      key: `${key}-paragraph`,
                    },
                  ],
                },
              },
            ],
            height: 40,
          },
        ],
        type: 'table',
      },
    ],
    key,
  };
}

function groupSlide(key: string): PptxSceneSlide {
  return {
    elements: [
      {
        authored: {
          transform: {
            childSpace: { height: 40, width: 160, x: 0, y: 0 },
            height: 40,
            width: 160,
            x: 10,
            y: 20,
          },
        },
        elements: [
          {
            authored: {
              transform: { height: 10, width: 20, x: 1, y: 2 },
            },
            key: `${key}-child`,
            resolved: { hidden: false },
            type: 'shape',
          },
        ],
        key: `${key}-group`,
        resolved: { hidden: false },
        type: 'group',
      },
    ],
    key,
  };
}

function chartSlide(key: string): PptxSceneSlide {
  return {
    elements: [
      {
        authored: {
          transform: { height: 200, width: 400, x: 10, y: 20 },
        },
        barDirection: 'col',
        chartType: 'barChart',
        grouping: 'clustered',
        key: `${key}-chart`,
        resolved: { hidden: false },
        series: [
          {
            categories: ['A', 'B'],
            color: '#4F46E5',
            key: `${key}-series`,
            name: 'Series',
            values: [1, 2],
          },
        ],
        type: 'chart',
      },
    ],
    key,
  };
}

function nestedChartSlide(key: string): PptxSceneSlide {
  const chartElement = chartSlide(key).elements[0]!;
  return {
    elements: [
      {
        authored: {
          transform: {
            childSpace: { height: 200, width: 400, x: 0, y: 0 },
            height: 200,
            width: 400,
            x: 10,
            y: 20,
          },
        },
        elements: [chartElement],
        key: `${key}-group`,
        resolved: { hidden: false },
        type: 'group',
      },
    ],
    key,
  };
}

function scene(slides: PptxSceneSlide[]): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides,
    themes: [],
  };
}

function generatedSlide(): PptxDocument['slides'][number] {
  return {
    elements: [],
    fill: { type: 'color', value: '#FFFFFF' },
    layoutElements: [],
    note: '',
  };
}

function document(slideCount: number): PptxDocument {
  return {
    size: { height: 540, width: 960 },
    slides: Array.from({ length: slideCount }, generatedSlide),
    themeColors: [],
    usedFonts: [],
  };
}

function generatedText(text = 'Text'): Text {
  return {
    borderColor: '#000000',
    borderStrokeDasharray: '0',
    borderType: 'solid',
    borderWidth: 0,
    content: `<p><span>${text}</span></p>`,
    fill: null,
    height: 40,
    id: '2',
    isFlipH: false,
    isFlipV: false,
    isVertical: false,
    left: 10,
    name: 'Text Box 2',
    order: 0,
    rotate: 0,
    top: 20,
    type: 'text',
    vAlign: 'up',
    width: 160,
    wrap: true,
  };
}

function generatedShape(shapType: string, text = 'Text'): Shape {
  return {
    ...generatedText(text),
    shapType,
    type: 'shape',
  };
}

function generatedNativeShape(): Shape {
  return {
    ...generatedShape('ellipse', ''),
    borderColor: '#0F172A',
    borderWidth: 2,
    fill: { type: 'color', value: '#F97316' },
  };
}

function generatedImage(): Image {
  return {
    base64: 'data:image/png;base64,iVBORw0KGgo=',
    blob: '',
    borderColor: '#000000',
    borderStrokeDasharray: '0',
    borderType: 'solid',
    borderWidth: 0,
    geom: 'rect',
    height: 40,
    id: '2',
    isFlipH: false,
    isFlipV: false,
    left: 10,
    order: 0,
    ref: 'ppt/media/image1.png',
    rotate: 0,
    top: 20,
    type: 'image',
    width: 160,
  };
}

function generatedTable(text = 'Expected'): Table {
  return {
    borders: {},
    colWidths: [160],
    data: [[{ borders: {}, text, vAlign: 'up' }]],
    height: 40,
    id: '2',
    isFlipH: false,
    isFlipV: false,
    left: 10,
    name: 'Table 2',
    order: 0,
    rotate: 0,
    rowHeights: [40],
    top: 20,
    type: 'table',
    width: 160,
  };
}

function generatedGroup(): Group {
  return {
    childSpace: { height: 40, width: 160, x: 0, y: 0 },
    elements: [
      {
        ...generatedShape('rect', ''),
        height: 10,
        id: '3',
        left: 1,
        top: 2,
        width: 20,
      },
    ],
    height: 40,
    id: '2',
    isFlipH: false,
    isFlipV: false,
    left: 10,
    order: 0,
    rotate: 0,
    top: 20,
    type: 'group',
    width: 160,
  };
}

function generatedChart(): CommonChart {
  return {
    barDir: 'col',
    chartType: 'barChart',
    colors: ['#4F46E5'],
    data: [
      {
        key: 'Series',
        values: [
          { x: '0', y: 1 },
          { x: '1', y: 2 },
        ],
        xlabels: { '0': 'A', '1': 'B' },
      },
    ],
    grouping: 'clustered',
    height: 200,
    id: '2',
    left: 10,
    order: 0,
    top: 20,
    type: 'chart',
    width: 400,
  };
}

function rendered(documentValue: PptxDocument): PptxSvgRenderResult {
  return {
    slides: documentValue.slides.map((_slide, index) => ({
      data: new TextEncoder().encode(
        '<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg"></svg>',
      ),
      format: 'svg',
      height: documentValue.size.height,
      mimeType: 'image/svg+xml',
      slideNumber: index + 1,
      warnings: [],
      width: documentValue.size.width,
    })),
  };
}

describe('PowerPoint creation verification', () => {
  it('passes the generated bytes and exact strict limits to the parser', async () => {
    const data = new Uint8Array([1, 2, 3]);
    const input = scene([emptySlide('slide-1')]);
    let receivedData: Uint8Array | undefined;
    let receivedOptions: PptxParseOptions | undefined;
    let renderedDocument: PptxDocument | undefined;
    const parsed = document(1);

    await verifyPowerPointCreationWithParser(
      data,
      input,
      (value, options) => {
        receivedData = value;
        receivedOptions = options;
        return Promise.resolve(parsed);
      },
      (value) => {
        renderedDocument = value;
        return rendered(value);
      },
    );

    expect(receivedData).toBe(data);
    expect(receivedOptions).toEqual({
      audioMode: 'none',
      errorMode: 'strict',
      imageMode: 'none',
      limits: { maxEntries: 11, maxSlides: 1 },
      videoMode: 'none',
    });
    expect(renderedDocument).toBe(parsed);
  });

  it('does not count ordinary elements as owned chart parts', async () => {
    const input = scene([textSlide('slide-1')]);
    const output = document(1);
    output.slides[0]?.elements.push(generatedText());
    let receivedOptions: PptxParseOptions | undefined;

    await verifyPowerPointCreationWithParser(
      new Uint8Array(),
      input,
      (_data, options) => {
        receivedOptions = options;
        return Promise.resolve(output);
      },
      rendered,
    );

    expect(receivedOptions?.limits?.maxEntries).toBe(11);
  });

  it.each([
    ['top-level', chartSlide('slide-1')],
    ['nested', nestedChartSlide('slide-1')],
  ] as const)(
    'counts one %s chart part in parser limits',
    async (_name, inputSlide) => {
      let receivedOptions: PptxParseOptions | undefined;
      await expect(
        verifyPowerPointCreationWithParser(
          new Uint8Array(),
          scene([inputSlide]),
          (_data, options) => {
            receivedOptions = options;
            return Promise.reject(new Error('stop after parser options'));
          },
          rendered,
        ),
      ).rejects.toThrow('stop after parser options');

      expect(receivedOptions?.limits?.maxEntries).toBe(12);
    },
  );

  it('verifies text and transform semantics before accepting a render', async () => {
    const output = document(1);
    const outputSlide = output.slides[0];
    if (outputSlide === undefined) throw new Error('Expected output slide');
    outputSlide.elements.push(generatedText());

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([textSlide('slide-1')]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).resolves.toBeUndefined();
  });

  it('accepts only the authored non-rect text geometry', async () => {
    const inputSlide = textSlide('slide-1');
    const expected = inputSlide.elements[0];
    if (expected?.type !== 'text') throw new Error('Expected input text');
    expected.authored.geometry = 'roundRect';
    const output = document(1);
    const outputSlide = output.slides[0];
    if (outputSlide === undefined) throw new Error('Expected output slide');
    outputSlide.elements.push(generatedShape('roundRect'));

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([inputSlide]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).resolves.toBeUndefined();

    outputSlide.elements = [generatedShape('ellipse')];
    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([inputSlide]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      new Error(
        'Generated PowerPoint text element missing at slide 1, element 1',
      ),
    );
  });

  it('verifies native shape transform and visual styling', async () => {
    const output = document(1);
    output.slides[0]?.elements.push(generatedNativeShape());

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([shapeSlide('slide-1')]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).resolves.toBeUndefined();
  });

  it('verifies an unstyled native rect without inventing optional styling', async () => {
    const inputSlide: PptxSceneSlide = {
      elements: [
        {
          authored: {
            transform: { height: 40, width: 160, x: 10, y: 20 },
          },
          key: 'plain-shape',
          resolved: { hidden: false },
          type: 'shape',
        },
      ],
      key: 'slide-1',
    };
    const output = document(1);
    output.slides[0]?.elements.push(generatedShape('rect', ''));

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([inputSlide]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).resolves.toBeUndefined();
  });

  it.each<[string, Partial<Shape> | undefined, string]>([
    ['missing shape', undefined, 'Generated PowerPoint shape missing'],
    [
      'geometry',
      { shapType: 'roundRect' },
      'Generated PowerPoint shape missing',
    ],
    ['fill', { fill: null }, 'Generated PowerPoint shape fill mismatch'],
    [
      'fill color',
      { fill: { type: 'color', value: '#FFFFFF' } },
      'Generated PowerPoint shape fill mismatch',
    ],
    [
      'line',
      { borderColor: '#FFFFFF' },
      'Generated PowerPoint shape line mismatch',
    ],
    [
      'line width',
      { borderWidth: 3 },
      'Generated PowerPoint shape line width mismatch',
    ],
    ['transform', { left: 11 }, 'Generated PowerPoint transform mismatch'],
  ])('rejects native shape %s mismatches', async (_name, change, message) => {
    const output = document(1);
    if (change !== undefined) {
      output.slides[0]?.elements.push({
        ...generatedNativeShape(),
        ...change,
      });
    } else if (output.slides[0] !== undefined) {
      output.slides[0].elements = new Array<Shape>(1);
    }

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([shapeSlide('slide-1')]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(`${message} at slide 1, element 1`);
  });

  it('verifies native image transform, media data, and geometry', async () => {
    const input = scene([imageSlide('slide-1')]);
    input.media = [
      { data: IMAGE_BYTES, key: 'media-1', mimeType: 'image/png' },
    ];
    const output = document(1);
    output.slides[0]?.elements.push(generatedImage());

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        input,
        () => Promise.resolve(output),
        rendered,
      ),
    ).resolves.toBeUndefined();
  });

  it('dispatches native table verification through the creation boundary', async () => {
    const output = document(1);
    output.slides[0]?.elements.push(generatedTable('Wrong'));

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([tableSlide('slide-1')]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      'Generated PowerPoint table text mismatch at slide 1, element 1, row 1, cell 1',
    );
  });

  it('verifies native chart transform, type, data, color, and options', async () => {
    const output = document(1);
    output.slides[0]?.elements.push(generatedChart());

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([chartSlide('slide-1')]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).resolves.toBeUndefined();
  });

  it.each([
    ['missing', undefined, 'chart missing'],
    ['transform', { left: 11 }, 'chart transform mismatch'],
    ['type', { chartType: 'lineChart' }, 'chart type mismatch'],
    ['data', { data: [] }, 'chart data mismatch'],
    ['color', { colors: ['#FFFFFF'] }, 'chart color mismatch'],
    ['options', { grouping: 'stacked' }, 'chart option mismatch'],
  ] as const)(
    'rejects a native chart %s mismatch',
    async (_name, replacement, message) => {
      const output = document(1);
      if (replacement === undefined) {
        output.slides[0]!.elements = new Array<Chart>(1);
      } else {
        output.slides[0]?.elements.push({
          ...generatedChart(),
          ...replacement,
        } as Chart);
      }

      await expect(
        verifyPowerPointCreationWithParser(
          new Uint8Array(),
          scene([chartSlide('slide-1')]),
          () => Promise.resolve(output),
          rendered,
        ),
      ).rejects.toThrow(
        `Generated PowerPoint ${message} at slide 1, element 1`,
      );
    },
  );

  it.each([
    ['x', { left: 11 }],
    ['y', { top: 21 }],
    ['width', { width: 401 }],
    ['height', { height: 201 }],
  ] as const)(
    'rejects a native chart transform %s mismatch',
    async (_key, replacement) => {
      const output = document(1);
      output.slides[0]?.elements.push({
        ...generatedChart(),
        ...replacement,
      });

      await expect(
        verifyPowerPointCreationWithParser(
          new Uint8Array(),
          scene([chartSlide('slide-1')]),
          () => Promise.resolve(output),
          rendered,
        ),
      ).rejects.toThrow(
        'Generated PowerPoint chart transform mismatch at slide 1, element 1',
      );
    },
  );

  it('requires an authored chart transform at its exact location', async () => {
    const input = chartSlide('slide-1');
    const expected = input.elements[0];
    if (expected?.type !== 'chart') throw new Error('Expected chart');
    delete expected.authored.transform;
    const output = document(1);
    output.slides[0]?.elements.push(generatedChart());

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([input]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      'Expected PowerPoint authored transform missing at slide 1, element 1',
    );
  });

  it('verifies an authored chart without a series color', async () => {
    const input = chartSlide('slide-1');
    const expected = input.elements[0];
    if (expected?.type !== 'chart') throw new Error('Expected chart');
    delete expected.series[0]!.color;
    const generated = generatedChart();
    generated.colors = [''];
    const output = document(1);
    output.slides[0]?.elements.push(generated);

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([input]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).resolves.toBeUndefined();
  });

  it('verifies line chart color independently from bar charts', async () => {
    const input = chartSlide('slide-1');
    const expected = input.elements[0];
    if (expected?.type !== 'chart') throw new Error('Expected chart');
    expected.chartType = 'lineChart';
    expected.grouping = 'standard';
    delete expected.barDirection;
    const generated = generatedChart();
    generated.chartType = 'lineChart';
    generated.grouping = 'standard';
    delete generated.barDir;
    generated.colors = ['#FFFFFF'];
    const output = document(1);
    output.slides[0]?.elements.push(generated);

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([input]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      'Generated PowerPoint chart color mismatch at slide 1, element 1',
    );
  });

  it.each(['pieChart', 'doughnutChart'] as const)(
    'verifies native %s options without series-color matching',
    async (chartType) => {
      const input = chartSlide('slide-1');
      const expected = input.elements[0];
      if (expected?.type !== 'chart') throw new Error('Expected chart');
      expected.chartType = chartType;
      delete expected.barDirection;
      delete expected.grouping;
      if (chartType === 'doughnutChart') expected.holeSize = 50;
      const generated = generatedChart();
      generated.chartType = chartType;
      generated.colors = [];
      delete generated.barDir;
      delete generated.grouping;
      if (chartType === 'doughnutChart') generated.holeSize = '50';
      const output = document(1);
      output.slides[0]?.elements.push(generated);

      await expect(
        verifyPowerPointCreationWithParser(
          new Uint8Array(),
          scene([input]),
          () => Promise.resolve(output),
          rendered,
        ),
      ).resolves.toBeUndefined();
    },
  );

  it.each(['barDir', 'grouping', 'holeSize', 'marker'] as const)(
    'rejects chart option mismatch %s independently',
    async (option) => {
      const input = chartSlide('slide-1');
      const expected = input.elements[0];
      if (expected?.type !== 'chart') throw new Error('Expected chart');
      const generated = generatedChart();
      if (option === 'barDir') generated.barDir = 'bar';
      if (option === 'grouping') generated.grouping = 'stacked';
      if (option === 'holeSize') {
        expected.chartType = 'doughnutChart';
        expected.holeSize = 50;
        delete expected.barDirection;
        delete expected.grouping;
        generated.chartType = 'doughnutChart';
        generated.holeSize = '51';
        delete generated.barDir;
        delete generated.grouping;
      }
      if (option === 'marker') {
        expected.chartType = 'lineChart';
        expected.grouping = 'standard';
        expected.marker = true;
        delete expected.barDirection;
        generated.chartType = 'lineChart';
        generated.grouping = 'standard';
        generated.marker = false;
        delete generated.barDir;
      }
      const output = document(1);
      output.slides[0]?.elements.push(generated);

      await expect(
        verifyPowerPointCreationWithParser(
          new Uint8Array(),
          scene([input]),
          () => Promise.resolve(output),
          rendered,
        ),
      ).rejects.toThrow(
        'Generated PowerPoint chart option mismatch at slide 1, element 1',
      );
    },
  );

  it('reports a missing group at its exact slide and element location', async () => {
    const secondSlide = groupSlide('slide-2');
    secondSlide.elements.unshift(textSlide('decoy').elements[0]!);
    const output = document(2);
    output.slides[1]?.elements.push(
      generatedText(),
      generatedShape('rect', ''),
    );

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([emptySlide('slide-1'), secondSlide]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      'Generated PowerPoint group missing at slide 2, element 2',
    );
  });

  it('recursively verifies every generated group child', async () => {
    const output = document(1);
    const group = generatedGroup();
    const child = group.elements[0];
    if (child?.type !== 'shape') throw new Error('Expected generated child');
    child.shapType = 'ellipse';
    output.slides[0]?.elements.push(group);

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([groupSlide('slide-1')]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      'Generated PowerPoint shape missing at slide 1, element 1',
    );
  });

  it('resolves image verification media by key instead of array position', async () => {
    const input = scene([imageSlide('slide-1')]);
    input.media = [
      {
        data: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
        key: 'decoy',
        mimeType: 'image/png',
      },
      { data: IMAGE_BYTES, key: 'media-1', mimeType: 'image/png' },
    ];
    const output = document(1);
    output.slides[0]?.elements.push(generatedImage());

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        input,
        () => Promise.resolve(output),
        rendered,
      ),
    ).resolves.toBeUndefined();
  });

  it('verifies every native image crop edge exactly', async () => {
    const input = scene([imageSlide('slide-1')]);
    input.media = [
      { data: IMAGE_BYTES, key: 'media-1', mimeType: 'image/png' },
    ];
    const expected = input.slides[0]?.elements[0];
    if (expected?.type !== 'image') throw new Error('Expected image scene');
    expected.crop = { bottom: -20, left: 30, right: 0, top: 10.125 };
    const output = document(1);
    output.slides[0]?.elements.push({
      ...generatedImage(),
      rect: { b: -20, l: 30, r: 0, t: 10.125 },
    });

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        input,
        () => Promise.resolve(output),
        rendered,
      ),
    ).resolves.toBeUndefined();
  });

  it.each([
    ['missing crop', undefined],
    ['bottom edge', { b: -19, l: 30, r: 0, t: 10.125 }],
    ['left edge', { b: -20, l: 31, r: 0, t: 10.125 }],
    ['right edge', { b: -20, l: 30, r: 1, t: 10.125 }],
    ['top edge', { b: -20, l: 30, r: 0, t: 11.125 }],
  ])('rejects native image %s mismatch', async (_name, rect) => {
    const input = scene([imageSlide('slide-1')]);
    input.media = [
      { data: IMAGE_BYTES, key: 'media-1', mimeType: 'image/png' },
    ];
    const expected = input.slides[0]?.elements[0];
    if (expected?.type !== 'image') throw new Error('Expected image scene');
    expected.crop = { bottom: -20, left: 30, right: 0, top: 10.125 };
    const output = document(1);
    output.slides[0]?.elements.push({
      ...generatedImage(),
      ...(rect === undefined ? {} : { rect }),
    });

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        input,
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      'Generated PowerPoint image crop mismatch at slide 1, element 1',
    );
  });

  it.each([
    ['missing image', undefined, 'Generated PowerPoint image missing'],
    [
      'media data',
      { base64: 'data:image/png;base64,AAAA' },
      'Generated PowerPoint image data mismatch',
    ],
    [
      'geometry',
      { geom: 'ellipse' },
      'Generated PowerPoint image geometry mismatch',
    ],
    [
      'crop',
      { rect: { b: 0, l: 1, r: 0, t: 0 } },
      'Generated PowerPoint image crop mismatch',
    ],
    ['transform', { left: 11 }, 'Generated PowerPoint transform mismatch'],
  ])('rejects native image %s mismatches', async (_name, change, message) => {
    const input = scene([imageSlide('slide-1')]);
    input.media = [
      { data: IMAGE_BYTES, key: 'media-1', mimeType: 'image/png' },
    ];
    const output = document(1);
    if (change !== undefined) {
      output.slides[0]?.elements.push({ ...generatedImage(), ...change });
    } else if (output.slides[0] !== undefined) {
      output.slides[0].elements = new Array<Image>(1);
    }

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        input,
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(`${message} at slide 1, element 1`);
  });

  it('rejects native image media missing from the scene inventory', async () => {
    const output = document(1);
    output.slides[0]?.elements.push(generatedImage());

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([imageSlide('slide-1')]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      'Expected PowerPoint image media missing at slide 1, element 1',
    );
  });

  it('rejects a generated text value that differs from the source scene', async () => {
    const output = document(1);
    const outputSlide = output.slides[0];
    if (outputSlide === undefined) throw new Error('Expected output slide');
    outputSlide.elements.push(generatedText('Wrong'));

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([textSlide('slide-1')]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      new Error('Generated PowerPoint text mismatch at slide 1, element 1'),
    );
  });

  it.each([
    ['x', { left: 11 }],
    ['y', { top: 21 }],
    ['width', { width: 161 }],
    ['height', { height: 41 }],
    ['rotation', { rotate: 1 }],
    ['horizontal flip', { isFlipH: true }],
    ['vertical flip', { isFlipV: true }],
  ])('rejects a generated transform with wrong %s', async (_name, change) => {
    const output = document(1);
    const outputSlide = output.slides[0];
    if (outputSlide === undefined) throw new Error('Expected output slide');
    const element = generatedText();
    Object.assign(element, change);
    outputSlide.elements.push(element);

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([textSlide('slide-1')]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      new Error(
        'Generated PowerPoint transform mismatch at slide 1, element 1',
      ),
    );
  });

  it('verifies run, break, and field text in source order', async () => {
    const inputSlide = textSlide('slide-1');
    const element = inputSlide.elements[0];
    if (element?.type !== 'text') throw new Error('Expected input text');
    const paragraph = element.text.paragraphs[0];
    if (paragraph === undefined) throw new Error('Expected input paragraph');
    paragraph.children = [
      { key: 'run', text: 'First', type: 'run' },
      { key: 'break', type: 'break' },
      { fieldType: 'slidenum', key: 'field', text: 'Second', type: 'field' },
    ];
    const output = document(1);
    const outputSlide = output.slides[0];
    if (outputSlide === undefined) throw new Error('Expected output slide');
    outputSlide.elements.push(generatedText('First<br>Second'));

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([inputSlide]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).resolves.toBeUndefined();
  });

  it('verifies paragraph separators instead of concatenating paragraphs', async () => {
    const inputSlide = textSlide('slide-1');
    const element = inputSlide.elements[0];
    if (element?.type !== 'text') throw new Error('Expected input text');
    element.text.paragraphs.push({
      children: [{ key: 'second-run', text: 'More', type: 'run' }],
      key: 'second-paragraph',
    });
    const output = document(1);
    const outputSlide = output.slides[0];
    if (outputSlide === undefined) throw new Error('Expected output slide');
    const generated = generatedText();
    generated.content = '<p><span>Text</span></p><p><span>More</span></p>';
    outputSlide.elements.push(generated);

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([inputSlide]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).resolves.toBeUndefined();
  });

  it('rejects a sparse generated text element with an exact location', async () => {
    const output = document(1);
    const outputSlide = output.slides[0];
    if (outputSlide === undefined) throw new Error('Expected output slide');
    outputSlide.elements = new Array<
      PptxDocument['slides'][number]['elements'][number]
    >(1);

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([textSlide('slide-1')]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      new Error(
        'Generated PowerPoint text element missing at slide 1, element 1',
      ),
    );
  });

  it('rejects a source text element without an authored transform', async () => {
    const inputSlide = textSlide('slide-1');
    const element = inputSlide.elements[0];
    if (element?.type !== 'text') throw new Error('Expected input text');
    delete element.authored.transform;
    const output = document(1);
    const outputSlide = output.slides[0];
    if (outputSlide === undefined) throw new Error('Expected output slide');
    outputSlide.elements.push(generatedText());

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([inputSlide]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      new Error(
        'Expected PowerPoint authored transform missing at slide 1, element 1',
      ),
    );
  });

  it('reports a non-text source element at its exact nested location', async () => {
    const inputSlide = textSlide('slide-2');
    inputSlide.elements.push({
      authored: {},
      feature: 'shape',
      key: 'unsupported',
      resolved: { hidden: false },
      type: 'unsupported',
    });
    const output = document(2);
    const outputSlide = output.slides[1];
    if (outputSlide === undefined) throw new Error('Expected output slide');
    outputSlide.elements.push(generatedText(), generatedText());

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([emptySlide('slide-1'), inputSlide]),
        () => Promise.resolve(output),
        rendered,
      ),
    ).rejects.toThrow(
      new Error(
        'Expected PowerPoint text element missing at slide 2, element 2',
      ),
    );
  });

  it('rejects a render that omits a generated slide', async () => {
    const output = document(1);

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([emptySlide('slide-1')]),
        () => Promise.resolve(output),
        () => ({ slides: [] }),
      ),
    ).rejects.toThrow(
      new Error(
        'Generated PowerPoint render count mismatch: expected 1, received 0',
      ),
    );
  });

  it.each([
    ['wrong width', { width: 959 }],
    ['wrong height', { height: 539 }],
    ['wrong slide number', { slideNumber: 2 }],
    ['wrong MIME type', { mimeType: 'text/plain' }],
    ['wrong format', { format: 'png' }],
    ['empty bytes', { data: new Uint8Array() }],
  ])('rejects a render with %s', async (_name, replacement) => {
    const output = document(1);
    const result = rendered(output);
    const slide = result.slides[0];
    if (slide === undefined) throw new Error('Expected rendered slide');
    Object.assign(slide, replacement);

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([emptySlide('slide-1')]),
        () => Promise.resolve(output),
        () => result,
      ),
    ).rejects.toThrow(
      new Error('Generated PowerPoint visual invariant mismatch on slide 1'),
    );
  });

  it.each([
    '<svg><script/></svg>',
    '<svg><foreignObject/></svg>',
    '<svg><image href="http://example.test/image.png"/></svg>',
    '<svg><image href="https://example.test/image.png"/></svg>',
    '<svg><image href="blob:unsafe"/></svg>',
    '<svg><image src="file:///tmp/unsafe.png"/></svg>',
    'not svg',
  ])('rejects unsafe rendered source %j', async (source) => {
    const output = document(1);
    const result = rendered(output);
    const slide = result.slides[0];
    if (slide === undefined) throw new Error('Expected rendered slide');
    slide.data = new TextEncoder().encode(source);

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([emptySlide('slide-1')]),
        () => Promise.resolve(output),
        () => result,
      ),
    ).rejects.toThrow(
      new Error('Generated PowerPoint unsafe SVG output on slide 1'),
    );
  });

  it('accepts an SVG root without an XML declaration', async () => {
    const output = document(1);
    const result = rendered(output);
    const slide = result.slides[0];
    if (slide === undefined) throw new Error('Expected rendered slide');
    slide.data = new TextEncoder().encode(
      '<svg xmlns="http://www.w3.org/2000/svg"></svg>',
    );

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([emptySlide('slide-1')]),
        () => Promise.resolve(output),
        () => result,
      ),
    ).resolves.toBeUndefined();
  });

  it('rejects bytes before an otherwise valid SVG root', async () => {
    const output = document(1);
    const result = rendered(output);
    const slide = result.slides[0];
    if (slide === undefined) throw new Error('Expected rendered slide');
    slide.data = new TextEncoder().encode(
      'unsafe<svg xmlns="http://www.w3.org/2000/svg"></svg>',
    );

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([emptySlide('slide-1')]),
        () => Promise.resolve(output),
        () => result,
      ),
    ).rejects.toThrow(
      new Error('Generated PowerPoint unsafe SVG output on slide 1'),
    );
  });

  it('uses a positive verification limit for a zero-slide package', async () => {
    let receivedOptions: PptxParseOptions | undefined;

    await verifyPowerPointCreationWithParser(
      new Uint8Array(),
      scene([]),
      (_value, options) => {
        receivedOptions = options;
        return Promise.resolve(document(0));
      },
    );

    expect(receivedOptions?.limits).toEqual({ maxEntries: 9, maxSlides: 1 });
  });

  it('rejects a different parsed slide count with exact evidence', async () => {
    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([emptySlide('slide-1')]),
        () => Promise.resolve(document(0)),
      ),
    ).rejects.toThrow(
      new Error(
        'Generated PowerPoint slide count mismatch: expected 1, received 0',
      ),
    );
  });

  it('rejects an independently wrong generated width', async () => {
    const output = document(0);
    output.size.width = 959;

    await expect(
      verifyPowerPointCreationWithParser(new Uint8Array(), scene([]), () =>
        Promise.resolve(output),
      ),
    ).rejects.toThrow(
      new Error(
        'Generated PowerPoint size mismatch: expected 960x540, received 959x540',
      ),
    );
  });

  it('rejects an independently wrong generated height', async () => {
    const output = document(0);
    output.size.height = 539;

    await expect(
      verifyPowerPointCreationWithParser(new Uint8Array(), scene([]), () =>
        Promise.resolve(output),
      ),
    ).rejects.toThrow(
      new Error(
        'Generated PowerPoint size mismatch: expected 960x540, received 960x539',
      ),
    );
  });

  it('rejects a present slide with the wrong element count', async () => {
    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([textSlide('slide-1')]),
        () => Promise.resolve(document(1)),
      ),
    ).rejects.toThrow(
      new Error(
        'Generated PowerPoint element count mismatch on slide 1: expected 1, received 0',
      ),
    );
  });

  it('identifies a missing generated slide at its one-based index', async () => {
    const output = document(2);
    output.slides = new Array<PptxDocument['slides'][number]>(2);
    output.slides[0] = generatedSlide();

    await expect(
      verifyPowerPointCreationWithParser(
        new Uint8Array(),
        scene([emptySlide('slide-1'), emptySlide('slide-2')]),
        () => Promise.resolve(output),
      ),
    ).rejects.toThrow(
      new Error(
        'Generated PowerPoint element count mismatch on slide 2: expected 0, received 0',
      ),
    );
  });

  it('uses the real strict parser at the production wrapper', async () => {
    await expect(
      verifyPowerPointCreation(new Uint8Array([1, 2, 3]), scene([])),
    ).rejects.toMatchObject({ name: 'PptxParseError' });
  });
});
