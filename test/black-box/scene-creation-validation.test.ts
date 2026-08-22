import { describe, expect, it } from 'vitest';

import { validatePptxScene } from '../../src';

function creationScene(): Record<string, unknown> {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {
              transform: {
                height: 80,
                rotation: 15,
                width: 300,
                x: 20,
                y: 30,
              },
            },
            key: 'text-1',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    {
                      key: 'run-1',
                      properties: { fontSize: 18 },
                      text: 'Hello',
                      type: 'run',
                    },
                  ],
                  key: 'paragraph-1',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'slide-1',
      },
    ],
    themes: [],
  };
}

function element(scene: Record<string, unknown>): Record<string, unknown> {
  const slides = scene.slides as Record<string, unknown>[];
  const slide = slides[0] as Record<string, unknown>;
  const elements = slide.elements as Record<string, unknown>[];
  return elements[0] as Record<string, unknown>;
}

function validateCreation(value: unknown) {
  return validatePptxScene(value, { profile: 'create-text-v1' });
}

function validateNativeCreation(value: unknown) {
  return validatePptxScene(value, { profile: 'create-native-v1' });
}

function addNativeImage(
  scene: Record<string, unknown>,
  authored: Record<string, unknown> = {
    transform: { height: 100, width: 160, x: 40, y: 50 },
  },
): Record<string, unknown> {
  scene.media = [
    {
      data: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      key: 'image-1',
      mimeType: 'image/png',
    },
  ];
  const image = {
    authored,
    key: 'picture-1',
    mediaKey: 'image-1',
    resolved: { hidden: false },
    type: 'image',
  };
  const slide = (scene.slides as Record<string, unknown>[])[0];
  if (slide === undefined) throw new Error('Expected slide');
  slide.elements = [image];
  return image;
}

function tableText(key: string): Record<string, unknown> {
  return {
    body: { anchor: 'center' },
    paragraphs: [
      {
        children: [{ key: `${key}-run`, text: key, type: 'run' }],
        key: `${key}-paragraph`,
      },
    ],
  };
}

function addNativeTable(
  scene: Record<string, unknown>,
): Record<string, unknown> {
  const table = {
    authored: {
      transform: { height: 80, width: 300, x: 40, y: 50 },
    },
    columns: [100, 200],
    key: 'table-1',
    resolved: { hidden: false },
    rows: [
      {
        cells: [{ text: tableText('cell-1') }, { text: tableText('cell-2') }],
        height: 80,
      },
    ],
    type: 'table',
  };
  const slide = (scene.slides as Record<string, unknown>[])[0];
  if (slide === undefined) throw new Error('Expected slide');
  slide.elements = [table];
  return table;
}

function firstRun(scene: Record<string, unknown>): Record<string, unknown> {
  const text = element(scene).text as Record<string, unknown>;
  const paragraphs = text.paragraphs as Record<string, unknown>[];
  const paragraph = paragraphs[0] as Record<string, unknown>;
  const children = paragraph.children as Record<string, unknown>[];
  return children[0] as Record<string, unknown>;
}

describe('PowerPoint creation scene validation', () => {
  it('accepts the bounded source-free text profile', () => {
    expect(validateCreation(creationScene())).toEqual({
      issues: [],
      valid: true,
    });
  });

  it('accepts bounded visual styling for rich text templates', () => {
    const scene = creationScene();
    const slides = scene.slides as Record<string, unknown>[];
    const slide = slides[0] as Record<string, unknown>;
    slide.backgroundColor = '#0F172A';
    const authored = element(scene).authored as Record<string, unknown>;
    authored.fillColor = '#1E293B';
    authored.geometry = 'roundRect';
    authored.lineColor = '#38BDF8';
    authored.lineWidth = 1.5;
    firstRun(scene).properties = {
      bold: true,
      color: '#F8FAFC',
      fontSize: 18,
    };

    expect(validateCreation(scene)).toEqual({ issues: [], valid: true });
  });

  it('accepts native shapes with bounded geometry and styling', () => {
    const scene = creationScene();
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements = [
      {
        authored: {
          fillColor: '#1E293B',
          geometry: 'roundRect',
          lineColor: '#38BDF8',
          lineWidth: 1.5,
          transform: { height: 120, width: 240, x: 40, y: 50 },
        },
        key: 'shape-1',
        resolved: { hidden: false },
        type: 'shape',
      },
    ];

    expect(validateNativeCreation(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene).issues).toContainEqual({
      code: 'unsupported-feature',
      message: 'Creation profile create-text-v1 supports text elements only',
      path: '$.slides[0].elements[0]',
    });
  });

  it('requires a serializable authored transform for native shapes', () => {
    const scene = creationScene();
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements = [
      {
        authored: { geometry: 'ellipse' },
        key: 'shape-1',
        resolved: { hidden: false },
        type: 'shape',
      },
    ];

    expect(validateNativeCreation(scene).issues).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Creation profile create-native-v1 requires an authored shape transform',
      path: '$.slides[0].elements[0].authored.transform',
    });
  });

  it('accepts structured native tables with exact grid dimensions', () => {
    const scene = creationScene();
    addNativeTable(scene);

    expect(validateNativeCreation(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene).issues).toContainEqual({
      code: 'unsupported-feature',
      message: 'Creation profile create-text-v1 supports text elements only',
      path: '$.slides[0].elements[0]',
    });
  });

  it.each([
    ['width', 301, 'column widths'],
    ['height', 81, 'row heights'],
  ])(
    'requires native table transform %s to match authored %s',
    (key, value, source) => {
      const scene = creationScene();
      const table = addNativeTable(scene);
      const authored = table.authored as Record<string, unknown>;
      const transform = authored.transform as Record<string, unknown>;
      transform[key] = value;

      expect(validateNativeCreation(scene).issues).toContainEqual({
        code: 'invalid-scene-document',
        message: `Table transform ${key} must equal the sum of its ${source}`,
        path: `$.slides[0].elements[0].authored.transform.${key}`,
      });
    },
  );

  it('requires bounded table spans and exact continuation flags', () => {
    const scene = creationScene();
    const table = addNativeTable(scene);
    table.rows = [
      {
        cells: [
          { colSpan: 2, rowSpan: 2, text: tableText('origin') },
          { hMerge: true, text: tableText('top-right') },
        ],
        height: 40,
      },
      {
        cells: [
          { text: tableText('bottom-left'), vMerge: true },
          {
            hMerge: true,
            text: tableText('bottom-right'),
            vMerge: true,
          },
        ],
        height: 40,
      },
    ];

    expect(validateNativeCreation(scene)).toEqual({ issues: [], valid: true });

    const rows = table.rows as Record<string, unknown>[];
    const bottom = rows[1] as Record<string, unknown>;
    const cells = bottom.cells as Record<string, unknown>[];
    delete cells[1]?.hMerge;
    expect(validateNativeCreation(scene).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Table merge continuation flags do not match its spans',
      path: '$.slides[0].elements[0].rows[1].cells[1]',
    });

    const origin = (
      (rows[0] as Record<string, unknown>).cells as Record<string, unknown>[]
    )[0];
    if (origin === undefined) throw new Error('Expected origin');
    origin.colSpan = 3;
    expect(validateNativeCreation(scene).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Table span exceeds the grid bounds',
      path: '$.slides[0].elements[0].rows[0].cells[0]',
    });
  });

  it('validates native table cells, borders, and exact row width', () => {
    const scene = creationScene();
    const table = addNativeTable(scene);
    const rows = table.rows as Record<string, unknown>[];
    const row = rows[0] as Record<string, unknown>;
    const cells = row.cells as Record<string, unknown>[];
    const first = cells[0] as Record<string, unknown>;
    first.fillColor = 'blue';
    first.borders = {
      top: { color: '#334155', style: 'double', width: 0 },
    };
    row.cells = [first];

    const issues = validateNativeCreation(scene).issues;
    expect(issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          path: '$.slides[0].elements[0].rows[0].cells',
        }),
        expect.objectContaining({
          path: '$.slides[0].elements[0].rows[0].cells[0].fillColor',
        }),
        expect.objectContaining({
          path: '$.slides[0].elements[0].rows[0].cells[0].borders.top.style',
        }),
        expect.objectContaining({
          path: '$.slides[0].elements[0].rows[0].cells[0].borders.top.width',
        }),
      ]),
    );
  });

  it('accepts signature-checked native image media and references', () => {
    const scene = creationScene();
    scene.media = [
      {
        data: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        key: 'image-1',
        mimeType: 'image/png',
      },
    ];
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements = [
      {
        authored: {
          transform: { height: 100, width: 160, x: 40, y: 50 },
        },
        key: 'picture-1',
        mediaKey: 'image-1',
        resolved: { hidden: false },
        type: 'image',
      },
    ];

    expect(validateNativeCreation(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene).issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'unsupported-feature',
          path: '$.media',
        }),
        expect.objectContaining({
          code: 'unsupported-feature',
          path: '$.slides[0].elements[0]',
        }),
      ]),
    );
  });

  it.each([
    ['image/png', new Uint8Array([0x89, 0x50]), 'PNG'],
    ['image/jpeg', new Uint8Array([0xff, 0xd8, 0xff, 0x00]), 'JPEG'],
  ])('rejects invalid %s media signatures', (mimeType, data, label) => {
    const scene = creationScene();
    scene.media = [{ data, key: 'image-1', mimeType }];

    expect(validateNativeCreation(scene).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: `${label} media data has an invalid signature`,
      path: '$.media[0].data',
    });
  });

  it('accepts a canonical JPEG signature', () => {
    const scene = creationScene();
    scene.media = [
      {
        data: new Uint8Array([0xff, 0xd8, 0xff, 0x00, 0xff, 0xd9]),
        key: 'image-1',
        mimeType: 'image/jpeg',
      },
    ];

    expect(validateNativeCreation(scene)).toEqual({ issues: [], valid: true });
  });

  it.each([
    ['prefix', new Uint8Array([0xfe, 0xd8, 0xff, 0x00, 0xff, 0xd9])],
    ['third marker byte', new Uint8Array([0xff, 0xd8, 0xfe, 0x00, 0xff, 0xd9])],
    ['penultimate byte', new Uint8Array([0xff, 0xd8, 0xff, 0x00, 0x00, 0xd9])],
    ['final byte', new Uint8Array([0xff, 0xd8, 0xff, 0x00, 0xff, 0x00])],
  ])('rejects a JPEG with an invalid %s', (_name, data) => {
    const scene = creationScene();
    scene.media = [{ data, key: 'image-1', mimeType: 'image/jpeg' }];

    expect(validateNativeCreation(scene).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'JPEG media data has an invalid signature',
      path: '$.media[0].data',
    });
  });

  it.each([
    ['non-binary data', 'bytes'],
    ['empty data', new Uint8Array()],
  ])('rejects %s in native media', (_name, data) => {
    const scene = creationScene();
    scene.media = [{ data, key: 'image-1', mimeType: 'image/png' }];

    expect(validateNativeCreation(scene).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected non-empty Uint8Array media data',
      path: '$.media[0].data',
    });
  });

  it('rejects unsupported media types with an exact diagnostic', () => {
    const scene = creationScene();
    scene.media = [
      { data: new Uint8Array([1]), key: 'image-1', mimeType: 'image/gif' },
    ];

    expect(validateNativeCreation(scene).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected image/png or image/jpeg media type',
      path: '$.media[0].mimeType',
    });
  });

  it('requires native image media keys and authored transforms', () => {
    const scene = creationScene();
    const image = addNativeImage(scene, {});
    delete image.mediaKey;

    expect(validateNativeCreation(scene).issues).toEqual(
      expect.arrayContaining([
        {
          code: 'invalid-scene-document',
          message: 'Expected a media key',
          path: '$.slides[0].elements[0].mediaKey',
        },
        {
          code: 'unsupported-feature',
          message:
            'Creation profile create-native-v1 requires an authored image transform',
          path: '$.slides[0].elements[0].authored.transform',
        },
      ]),
    );
  });

  it('accepts bounded native image crop percentages', () => {
    const scene = creationScene();
    const image = addNativeImage(scene);
    image.crop = { bottom: 1.015, left: 30, right: 0, top: 10.125 };

    expect(validateNativeCreation(scene)).toEqual({ issues: [], valid: true });
  });

  it.each([
    { bottom: 0, left: -100, right: 0, top: 0 },
    { bottom: 0, left: 100, right: -1, top: 0 },
  ])('accepts exact native image crop boundary %j', (crop) => {
    const scene = creationScene();
    const image = addNativeImage(scene);
    image.crop = crop;

    expect(validateNativeCreation(scene)).toEqual({ issues: [], valid: true });
  });

  it('reports exact quantization and native crop range messages', () => {
    const quantizedScene = creationScene();
    const quantized = addNativeImage(quantizedScene);
    quantized.crop = { bottom: 0, left: 0.0001, right: 0, top: 0 };
    expect(validateNativeCreation(quantizedScene).issues).toContainEqual({
      code: 'invalid-numeric-value',
      message:
        'Image crop must be a finite percentage with at most three decimal places',
      path: '$.slides[0].elements[0].crop.left',
    });

    const rangeScene = creationScene();
    const ranged = addNativeImage(rangeScene);
    ranged.crop = { bottom: 0, left: 101, right: -2, top: 0 };
    expect(validateNativeCreation(rangeScene).issues).toContainEqual({
      code: 'invalid-numeric-value',
      message: 'Native image crop must be from -100 through 100',
      path: '$.slides[0].elements[0].crop.left',
    });

    const nonCascadingScene = creationScene();
    const nonCascading = addNativeImage(nonCascadingScene);
    nonCascading.crop = { bottom: 0, left: 101, right: 0, top: 0 };
    expect(validateNativeCreation(nonCascadingScene).issues).toEqual([
      {
        code: 'invalid-numeric-value',
        message: 'Native image crop must be from -100 through 100',
        path: '$.slides[0].elements[0].crop.left',
      },
    ]);
  });

  it('preserves finite source crops outside the native write profile', () => {
    const scene = creationScene();
    const image = addNativeImage(scene);
    image.crop = { bottom: 140, left: 120, right: 0, top: 110 };

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(validateNativeCreation(scene).issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'invalid-numeric-value',
          path: '$.slides[0].elements[0].crop.bottom',
        }),
        expect.objectContaining({
          code: 'invalid-numeric-value',
          path: '$.slides[0].elements[0].crop.left',
        }),
        expect.objectContaining({
          code: 'invalid-numeric-value',
          path: '$.slides[0].elements[0].crop.top',
        }),
      ]),
    );
  });

  it.each([
    ['missing edge', { bottom: 0, left: 0, right: 0 }, 'top'],
    [
      'unknown edge',
      { bottom: 0, diagonal: 1, left: 0, right: 0, top: 0 },
      'diagonal',
    ],
    ['string edge', { bottom: 0, left: '1', right: 0, top: 0 }, 'left'],
    [
      'non-finite edge',
      { bottom: 0, left: 0, right: Infinity, top: 0 },
      'right',
    ],
    ['low edge', { bottom: -100.001, left: 0, right: 0, top: 0 }, 'bottom'],
    ['high edge', { bottom: -2, left: 0, right: 0, top: 101 }, 'top'],
    ['precision', { bottom: 0, left: 0.0001, right: 0, top: 0 }, 'left'],
  ])('rejects image crop %s', (_name, crop, field) => {
    const scene = creationScene();
    const image = addNativeImage(scene);
    image.crop = crop;

    expect(validateNativeCreation(scene).issues).toContainEqual(
      expect.objectContaining({
        path: `$.slides[0].elements[0].crop.${field}`,
      }),
    );
  });

  it.each([
    [
      'horizontal',
      { bottom: 0, left: 60, right: 40, top: 0 },
      'Horizontal image crop must leave a positive visible region',
    ],
    [
      'vertical',
      { bottom: 50, left: 0, right: 0, top: 50 },
      'Vertical image crop must leave a positive visible region',
    ],
  ])('rejects a collapsed %s image crop', (_name, crop, message) => {
    const scene = creationScene();
    const image = addNativeImage(scene);
    image.crop = crop;

    expect(validateNativeCreation(scene).issues).toContainEqual({
      code: 'invalid-numeric-value',
      message,
      path: '$.slides[0].elements[0].crop',
    });
  });

  it.each(['fillColor', 'geometry', 'lineColor', 'lineWidth'])(
    'rejects authored image shape styling property %s',
    (property) => {
      const scene = creationScene();
      addNativeImage(scene, {
        [property]: property === 'lineWidth' ? 1 : '#FFFFFF',
        transform: { height: 100, width: 160, x: 40, y: 50 },
      });

      expect(validateNativeCreation(scene).issues).toContainEqual({
        code: 'unsupported-feature',
        message:
          'Creation profile create-native-v1 does not apply shape styling to images',
        path: '$.slides[0].elements[0].authored',
      });
    },
  );

  it('reports exact text-profile image and media exclusions', () => {
    const scene = creationScene();
    addNativeImage(scene);
    const issues = validateCreation(scene).issues;

    expect(issues).toContainEqual({
      code: 'unsupported-feature',
      message: 'Creation profile create-text-v1 supports text elements only',
      path: '$.slides[0].elements[0]',
    });
    expect(issues).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Creation profile create-text-v1 does not support media resources',
      path: '$.media',
    });
  });

  it('rejects opaque elements from the native creation profile', () => {
    const scene = creationScene();
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements = [
      {
        authored: {},
        feature: 'chart',
        key: 'opaque-1',
        resolved: { hidden: false },
        type: 'unsupported',
      },
    ];

    expect(validateNativeCreation(scene).issues).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Creation profile create-native-v1 does not support opaque elements',
      path: '$.slides[0].elements[0]',
    });
  });

  it('rejects dangling media references and unsupported image styling', () => {
    const scene = creationScene();
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements = [
      {
        authored: {
          fillColor: '#FFFFFF',
          transform: { height: 100, width: 160, x: 40, y: 50 },
        },
        key: 'picture-1',
        mediaKey: 'missing-image',
        resolved: { hidden: false },
        type: 'image',
      },
    ];

    expect(validateNativeCreation(scene).issues).toEqual(
      expect.arrayContaining([
        {
          code: 'unsupported-feature',
          message:
            'Creation profile create-native-v1 does not apply shape styling to images',
          path: '$.slides[0].elements[0].authored',
        },
        {
          code: 'invalid-hierarchy-reference',
          message: 'Reference points to an unknown public key',
          path: '$.slides[0].elements[0].mediaKey',
        },
      ]),
    );
  });

  it.each(['ellipse', 'rect', 'roundRect'])(
    'accepts text shape geometry %s',
    (geometry) => {
      const scene = creationScene();
      const authored = element(scene).authored as Record<string, unknown>;
      authored.geometry = geometry;

      expect(validateCreation(scene)).toEqual({ issues: [], valid: true });
    },
  );

  it.each([
    [
      '$.slides[0].backgroundColor',
      (scene: Record<string, unknown>) => {
        const slides = scene.slides as Record<string, unknown>[];
        const slide = slides[0] as Record<string, unknown>;
        slide.backgroundColor = '0F172A';
      },
    ],
    [
      '$.slides[0].elements[0].authored.fillColor',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        authored.fillColor = '#XYZ123';
      },
    ],
    [
      '$.slides[0].elements[0].authored.lineColor',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        authored.lineColor = '#1234';
      },
    ],
    [
      '$.slides[0].elements[0].text.paragraphs[0].children[0].properties.color',
      (scene: Record<string, unknown>) => {
        firstRun(scene).properties = { color: 'red' };
      },
    ],
  ])('rejects a malformed visual color at %s', (path, mutate) => {
    const scene = creationScene();
    mutate(scene);

    expect(validateCreation(scene).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected a #RRGGBB color',
      path,
    });
  });

  it.each(['x#0F172A', '#0F172Ax'])(
    'rejects a color with data outside the #RRGGBB boundary: %s',
    (backgroundColor) => {
      const scene = creationScene();
      const slides = scene.slides as Record<string, unknown>[];
      const slide = slides[0] as Record<string, unknown>;
      slide.backgroundColor = backgroundColor;

      expect(validateCreation(scene).issues).toContainEqual({
        code: 'invalid-scene-document',
        message: 'Expected a #RRGGBB color',
        path: '$.slides[0].backgroundColor',
      });
    },
  );

  it('rejects unsupported geometry and unsafe line widths', () => {
    const scene = creationScene();
    const authored = element(scene).authored as Record<string, unknown>;
    authored.geometry = 'star';
    authored.lineWidth = 0;

    expect(validateCreation(scene).issues).toEqual(
      expect.arrayContaining([
        {
          code: 'invalid-scene-document',
          message: 'Unknown text shape geometry',
          path: '$.slides[0].elements[0].authored.geometry',
        },
        {
          code: 'invalid-numeric-value',
          message: 'Expected a positive finite number',
          path: '$.slides[0].elements[0].authored.lineWidth',
        },
      ]),
    );
  });

  it.each([
    [1_000_000_000_000, 'Value exceeds the safe OOXML integer range'],
    [0.000_039, 'Value must round to a positive OOXML integer'],
  ])(
    'applies serialized line width bounds only to creation for %s',
    (lineWidth, message) => {
      const scene = creationScene();
      const authored = element(scene).authored as Record<string, unknown>;
      authored.lineWidth = lineWidth;

      expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
      expect(validateCreation(scene).issues).toContainEqual({
        code: 'invalid-numeric-value',
        message,
        path: '$.slides[0].elements[0].authored.lineWidth',
      });
    },
  );

  it('requires authored geometry without changing the general scene profile', () => {
    const scene = creationScene();
    element(scene).authored = {};

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene)).toEqual({
      issues: [
        {
          code: 'unsupported-feature',
          message:
            'Creation profile create-text-v1 requires an authored text transform',
          path: '$.slides[0].elements[0].authored.transform',
        },
      ],
      valid: false,
    });
  });

  it('rejects explicit hierarchy until its writer mapping ships', () => {
    const scene = creationScene();
    scene.themes = [{ key: 'theme-1' }];
    scene.masters = [{ elements: [], key: 'master-1', themeKey: 'theme-1' }];
    scene.layouts = [{ elements: [], key: 'layout-1', masterKey: 'master-1' }];
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide) slide.layoutKey = 'layout-1';

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene).issues).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Creation profile create-text-v1 generates its own minimal hierarchy',
      path: '$',
    });
  });

  it('rejects placeholders until owner and inheritance serialization ships', () => {
    const scene = creationScene();
    element(scene).placeholder = { role: 'slide-instance' };

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene).issues).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Creation profile create-text-v1 does not support placeholders yet',
      path: '$.slides[0].elements[0].placeholder',
    });
  });

  it('does not apply text geometry requirements to preservation-only elements', () => {
    const scene = creationScene();
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide) {
      slide.elements = [
        {
          authored: {},
          feature: 'chart',
          key: 'chart-1',
          resolved: { hidden: false },
          type: 'unsupported',
        },
      ];
    }

    expect(validateCreation(scene)).toEqual({
      issues: [
        {
          code: 'unsupported-feature',
          message:
            'Creation profile create-text-v1 supports text elements only',
          path: '$.slides[0].elements[0]',
        },
      ],
      valid: false,
    });
  });

  it.each([
    [
      '$.size.width',
      (scene: Record<string, unknown>) => {
        scene.size = { height: 540, width: 1_000_000_000_000 };
      },
    ],
    [
      '$.size.height',
      (scene: Record<string, unknown>) => {
        scene.size = { height: 1_000_000_000_000, width: 960 };
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.x',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.x = 1_000_000_000_000;
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.y',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.y = -1_000_000_000_000;
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.width',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.width = 1_000_000_000_000;
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.height',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.height = 1_000_000_000_000;
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.rotation',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.rotation = 1_000_000_000_000;
      },
    ],
    [
      '$.slides[0].elements[0].text.paragraphs[0].children[0].properties.fontSize',
      (scene: Record<string, unknown>) => {
        const text = element(scene).text as Record<string, unknown>;
        const paragraphs = text.paragraphs as Record<string, unknown>[];
        const paragraph = paragraphs[0] as Record<string, unknown>;
        const children = paragraph.children as Record<string, unknown>[];
        const run = children[0] as Record<string, unknown>;
        run.properties = { fontSize: 1_000_000_000_000_000 };
      },
    ],
  ] as const)(
    'rejects a value outside the serializable integer range at %s',
    (path, mutate) => {
      const scene = creationScene();
      mutate(scene);

      expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
      expect(validateCreation(scene).issues).toContainEqual({
        code: 'invalid-numeric-value',
        message: 'Value exceeds the safe OOXML integer range',
        path,
      });
    },
  );

  it('checks resolved transform ranges before a future writer can depend on them', () => {
    const scene = creationScene();
    element(scene).resolved = {
      hidden: false,
      transform: {
        height: 80,
        rotation: -1_000_000_000_000,
        width: 300,
        x: 20,
        y: 30,
      },
    };

    expect(validateCreation(scene).issues).toContainEqual({
      code: 'invalid-numeric-value',
      message: 'Value exceeds the safe OOXML integer range',
      path: '$.slides[0].elements[0].resolved.transform.rotation',
    });
  });

  it('does not add duplicate range errors for non-numeric values', () => {
    const scene = creationScene();
    scene.size = { height: 540, width: Number.POSITIVE_INFINITY };

    expect(validateCreation(scene)).toEqual({
      issues: [
        {
          code: 'invalid-numeric-value',
          message: 'Expected a positive finite number',
          path: '$.size.width',
        },
      ],
      valid: false,
    });
  });

  it.each([
    [
      '$.size.width',
      (scene: Record<string, unknown>) => {
        scene.size = { height: 540, width: 0.000_039 };
      },
    ],
    [
      '$.size.height',
      (scene: Record<string, unknown>) => {
        scene.size = { height: 0.000_039, width: 960 };
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.width',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.width = 0.000_039;
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.height',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.height = 0.000_039;
      },
    ],
    [
      '$.slides[0].elements[0].resolved.transform.width',
      (scene: Record<string, unknown>) => {
        element(scene).resolved = {
          hidden: false,
          transform: {
            height: 80,
            width: 0.000_039,
            x: 20,
            y: 30,
          },
        };
      },
    ],
    [
      '$.slides[0].elements[0].text.paragraphs[0].children[0].properties.fontSize',
      (scene: Record<string, unknown>) => {
        const text = element(scene).text as Record<string, unknown>;
        const paragraphs = text.paragraphs as Record<string, unknown>[];
        const paragraph = paragraphs[0] as Record<string, unknown>;
        const children = paragraph.children as Record<string, unknown>[];
        const run = children[0] as Record<string, unknown>;
        run.properties = { fontSize: 0.004 };
      },
    ],
  ] as const)(
    'rejects a positive value that quantizes to zero at %s',
    (path, mutate) => {
      const scene = creationScene();
      mutate(scene);

      expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
      expect(validateCreation(scene).issues).toContainEqual({
        code: 'invalid-numeric-value',
        message: 'Value must round to a positive OOXML integer',
        path,
      });
    },
  );

  it('accepts the smallest positive values that quantize to one', () => {
    const scene = creationScene();
    scene.size = { height: 0.000_04, width: 0.000_04 };
    const authored = element(scene).authored as Record<string, unknown>;
    const transform = authored.transform as Record<string, unknown>;
    transform.height = 0.000_04;
    transform.width = 0.000_04;
    const text = element(scene).text as Record<string, unknown>;
    const paragraphs = text.paragraphs as Record<string, unknown>[];
    const paragraph = paragraphs[0] as Record<string, unknown>;
    paragraph.endProperties = { fontSize: 0.006 };

    expect(validateCreation(scene)).toEqual({ issues: [], valid: true });
  });

  it('accepts signed positions and rotation in the creation profile', () => {
    const scene = creationScene();
    const authored = element(scene).authored as Record<string, unknown>;
    authored.transform = {
      height: 80,
      rotation: -45,
      width: 300,
      x: -20,
      y: -0.000_04,
    };

    expect(validateCreation(scene)).toEqual({ issues: [], valid: true });
  });

  it('rejects a scene beyond the bounded creation slide count', () => {
    const scene = creationScene();
    scene.slides = new Array(10_001);

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene)).toEqual({
      issues: [
        {
          code: 'unsupported-feature',
          message:
            'Creation profile create-text-v1 supports at most 10000 slides',
          path: '$.slides',
        },
      ],
      valid: false,
    });
  });

  it('rejects an oversized element graph only in the creation profile', () => {
    const scene = creationScene();
    const slides = scene.slides as Record<string, unknown>[];
    const firstSlide = slides[0];
    if (firstSlide) firstSlide.elements = new Array(5_001);

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene)).toEqual({
      issues: [
        {
          code: 'resource-limit-exceeded',
          message:
            'Creation profile create-text-v1 supports at most 5000 elements',
          path: '$.slides',
        },
      ],
      valid: false,
    });
  });

  it('does not traverse creation resources after structural validation fails', () => {
    const scene = creationScene();
    scene.schemaVersion = 1;
    const slides = scene.slides as Record<string, unknown>[];
    const firstSlide = slides[0];
    if (firstSlide) firstSlide.elements = new Array(5_001);

    expect(validateCreation(scene)).toEqual({
      issues: [
        {
          code: 'unsupported-schema-version',
          message: 'Only PowerPoint scene schema version 2 is supported',
          path: '$.schemaVersion',
        },
      ],
      valid: false,
    });
  });
});
