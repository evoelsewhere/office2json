import { describe, expect, it } from 'vitest';

import { decodeBase64 } from '../../src/common/binary/base64';
import { createPptx } from '../../src/formats/pptx/creator';
import {
  applyPptxRoundTripOperationsToPreview,
  normalizePptxRoundTripGroupTransform,
  normalizePptxRoundTripImageCrop,
  normalizePptxRoundTripTransform,
  replacePptxRoundTripText,
  setPptxRoundTripImageCrop,
  setPptxRoundTripTextTransform,
  validatePptxRoundTripReplaceTextRequest,
} from '../../src/formats/pptx/roundtrip/edit';
import { readPptxRoundTrip } from '../../src/formats/pptx/roundtrip/read';
import type { PptxSceneDocument } from '../../src/formats/pptx/scene-types';

function scene(): PptxSceneDocument {
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
              transform: { height: 80, width: 300, x: 20, y: 30 },
            },
            key: 'source-text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    { key: 'source-run', text: 'Before', type: 'run' },
                  ],
                  key: 'source-paragraph',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'source-slide',
      },
    ],
    themes: [],
  };
}

async function snapshot() {
  const created = await createPptx(scene());
  return readPptxRoundTrip(created.data);
}

function nativeTextOwnersScene(): PptxSceneDocument {
  const textBody = (key: string, value: string) => ({
    body: {},
    paragraphs: [
      {
        children: [{ key: `${key}-run`, text: value, type: 'run' as const }],
        key: `${key}-paragraph`,
      },
    ],
  });
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
              transform: { height: 80, width: 240, x: 20, y: 30 },
            },
            columns: [120, 120],
            key: 'source-table',
            resolved: { hidden: false },
            rows: [
              {
                cells: [
                  { text: textBody('cell-a', 'Alpha') },
                  { text: textBody('cell-b', 'Beta') },
                ],
                height: 80,
              },
            ],
            type: 'table',
          },
          {
            authored: {
              transform: {
                childSpace: { height: 100, width: 200, x: 0, y: 0 },
                height: 100,
                width: 200,
                x: 300,
                y: 30,
              },
            },
            elements: [
              {
                authored: {
                  transform: { height: 50, width: 100, x: 0, y: 0 },
                },
                key: 'nested-text',
                resolved: { hidden: false },
                text: textBody('nested-text', 'Nested'),
                type: 'text',
              },
            ],
            key: 'source-group',
            resolved: { hidden: false },
            type: 'group',
          },
        ],
        key: 'source-slide',
      },
    ],
    themes: [],
  };
}

async function nativeTextOwnersSnapshot() {
  const created = await createPptx(nativeTextOwnersScene());
  return readPptxRoundTrip(created.data);
}

const IMAGE_BYTES = decodeBase64(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
);

async function imageSnapshot(cropped = false) {
  const document: PptxSceneDocument = {
    layouts: [],
    masters: [],
    media: [{ data: IMAGE_BYTES, key: 'media', mimeType: 'image/png' }],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {
              transform: { height: 90, width: 120, x: 20, y: 30 },
            },
            ...(cropped
              ? { crop: { bottom: -20, left: 30, right: 0, top: 10 } }
              : {}),
            key: 'image',
            mediaKey: 'media',
            resolved: { hidden: false },
            type: 'image',
          },
        ],
        key: 'slide',
      },
    ],
    themes: [],
  };
  return readPptxRoundTrip((await createPptx(document)).data);
}

describe('PowerPoint round-trip text edit binding', () => {
  it('binds an exact preview precondition without mutating source state', async () => {
    const source = await snapshot();
    const before = structuredClone(source);
    const edited = await replacePptxRoundTripText(source, {
      targetKey: 'slide-1-element-1-run-1',
      value: 'After <&',
    });

    expect(source).toEqual(before);
    expect(edited.document).toEqual(source.document);
    expect(edited.source).toEqual(source.source);
    expect(edited.operations).toEqual([
      {
        expectedText: 'Before',
        id: 'replace-text-1',
        kind: 'replace-text',
        targetKey: 'slide-1-element-1-run-1',
        value: 'After <&',
      },
    ]);
    expect(edited.supportProfile).toEqual({
      effectiveLevel: 'R2',
      id: 'pptx-roundtrip-text-v1',
      producerMatrix: [],
      version: '1',
    });
    expect(edited.consistency.operationsSha256).not.toBe(
      source.consistency.operationsSha256,
    );
    expect(edited.consistency.semanticPreviewSha256).toBe(
      source.consistency.semanticPreviewSha256,
    );
  });

  it('is deterministic across isolated calls', async () => {
    const source = await snapshot();
    const [first, second] = await Promise.all([
      replacePptxRoundTripText(source, {
        targetKey: 'slide-1-element-1-run-1',
        value: 'After',
      }),
      replacePptxRoundTripText(source, {
        targetKey: 'slide-1-element-1-run-1',
        value: 'After',
      }),
    ]);

    expect(second).toEqual(first);
  });

  it.each([
    ['table cell', 'slide-1-element-1-row-1-cell-2-run-1', 'Updated cell'],
    ['nested group text', 'slide-1-element-2-element-1-run-1', 'Updated group'],
  ])('binds and previews native %s edits', async (_name, targetKey, value) => {
    const source = await nativeTextOwnersSnapshot();
    const edited = await replacePptxRoundTripText(source, {
      targetKey,
      value,
    });

    expect(edited.supportProfile).toMatchObject({
      effectiveLevel: 'R2',
      id: 'pptx-roundtrip-native-v1',
    });
    expect(edited.operations).toMatchObject([
      { kind: 'replace-text', targetKey, value },
    ]);
    expect(
      JSON.stringify(applyPptxRoundTripOperationsToPreview(edited)),
    ).toContain(value);
  });

  it.each([
    ['', 'After', 'target key must be a non-empty string'],
    ['missing', 'After', 'target key does not exist'],
    ['slide-1-element-1-run-1', 'Before', 'must change the target value'],
    ['slide-1-element-1-run-1', 'bad\u0000', 'is not safe XML text'],
  ])(
    'rejects target %s with an invalid edit',
    async (targetKey, value, message) => {
      const editing = replacePptxRoundTripText(await snapshot(), {
        targetKey,
        value,
      });
      await expect(editing).rejects.toMatchObject({
        code: 'invalid-edit-operation',
      });
      await expect(editing).rejects.toThrow(message);
    },
  );

  it('rejects a second operation for the same target', async () => {
    const first = await replacePptxRoundTripText(await snapshot(), {
      targetKey: 'slide-1-element-1-run-1',
      value: 'After',
    });

    await expect(
      replacePptxRoundTripText(first, {
        targetKey: 'slide-1-element-1-run-1',
        value: 'Again',
      }),
    ).rejects.toMatchObject({
      code: 'invalid-edit-operation',
      message: 'PowerPoint text edit target is already scheduled',
    });
  });

  it('applies a transform only to the exact text element target', async () => {
    const transformed = await setPptxRoundTripTextTransform(await snapshot(), {
      targetKey: 'slide-1-element-1',
      value: { height: 90, width: 320, x: 40, y: 50 },
    });
    transformed.document.slides[0]?.elements.push(
      {
        authored: {},
        key: 'decoy-text',
        resolved: {
          hidden: false,
          transform: { height: 10, width: 20, x: 1, y: 2 },
        },
        text: {
          body: {},
          paragraphs: [
            {
              children: [{ key: 'decoy-run', text: 'Decoy', type: 'run' }],
              key: 'decoy-paragraph',
            },
          ],
        },
        type: 'text',
      },
      {
        authored: {},
        feature: 'shape',
        key: 'slide-1-element-1',
        resolved: {
          hidden: false,
          transform: { height: 5, width: 6, x: 3, y: 4 },
        },
        type: 'unsupported',
      },
    );

    const preview = applyPptxRoundTripOperationsToPreview(transformed);

    expect(preview.slides[0]?.elements[0]?.resolved.transform).toEqual({
      flipHorizontal: false,
      flipVertical: false,
      height: 90,
      rotation: 0,
      width: 320,
      x: 40,
      y: 50,
    });
    expect(preview.slides[0]?.elements[1]?.resolved.transform).toEqual({
      height: 10,
      width: 20,
      x: 1,
      y: 2,
    });
    expect(preview.slides[0]?.elements[2]?.resolved.transform).toEqual({
      height: 5,
      width: 6,
      x: 3,
      y: 4,
    });
  });

  it('applies text only to the exact run target', async () => {
    const edited = await replacePptxRoundTripText(await snapshot(), {
      targetKey: 'slide-1-element-1-run-1',
      value: 'After',
    });
    const element = edited.document.slides[0]?.elements[0];
    if (element?.type !== 'text') throw new Error('Expected editable text');
    const paragraph = element.text.paragraphs[0];
    if (paragraph === undefined) throw new Error('Expected paragraph');
    paragraph.children.push(
      { key: 'decoy-run', text: 'Decoy', type: 'run' },
      { key: 'slide-1-element-1-run-1', type: 'break' },
    );

    const preview = applyPptxRoundTripOperationsToPreview(edited);
    const previewElement = preview.slides[0]?.elements[0];
    if (previewElement?.type !== 'text')
      throw new Error('Expected preview text');
    expect(previewElement.text.paragraphs[0]?.children).toMatchObject([
      { key: 'slide-1-element-1-run-1', text: 'After', type: 'run' },
      { key: 'decoy-run', text: 'Decoy', type: 'run' },
      { key: 'slide-1-element-1-run-1', type: 'break' },
    ]);
  });

  it.each([
    ['replace-text', 'missing-run'],
    ['set-transform', 'missing-element'],
  ] as const)(
    'rejects a disappeared %s preview target',
    async (kind, targetKey) => {
      const value = await snapshot();
      value.operations =
        kind === 'replace-text'
          ? [
              {
                expectedText: 'Before',
                id: 'replace-text-1',
                kind,
                targetKey,
                value: 'After',
              },
            ]
          : [
              {
                expectedTransform: {
                  flipHorizontal: false,
                  flipVertical: false,
                  height: 80,
                  rotation: 0,
                  width: 300,
                  x: 20,
                  y: 30,
                },
                id: 'set-transform-1',
                kind,
                targetKey,
                value: {
                  flipHorizontal: false,
                  flipVertical: false,
                  height: 90,
                  rotation: 0,
                  width: 320,
                  x: 40,
                  y: 50,
                },
              },
            ];

      expect(() => applyPptxRoundTripOperationsToPreview(value)).toThrow(
        kind === 'replace-text'
          ? `PowerPoint text edit verification target disappeared: ${targetKey}`
          : `PowerPoint transform verification target disappeared: ${targetKey}`,
      );
    },
  );

  it.each([
    ['', 'PowerPoint transform target key must be a non-empty string'],
    [7, 'PowerPoint transform target key must be a non-empty string'],
    ['missing', 'PowerPoint transform target key does not exist'],
  ])('rejects transform target %j', async (targetKey, message) => {
    await expect(
      setPptxRoundTripTextTransform(await snapshot(), {
        targetKey: targetKey as string,
        value: { height: 90, width: 320, x: 40, y: 50 },
      }),
    ).rejects.toThrow(message);
  });

  it.each([null, [], 7])('rejects transform value %j', async (value) => {
    await expect(
      setPptxRoundTripTextTransform(await snapshot(), {
        targetKey: 'slide-1-element-1',
        value: value as never,
      }),
    ).rejects.toThrow('PowerPoint transform value must be an object');
  });

  it('rejects a transform target without resolved geometry', async () => {
    const value = await snapshot();
    const element = value.document.slides[0]?.elements[0];
    if (element === undefined) throw new Error('Expected element');
    delete element.resolved.transform;

    await expect(
      setPptxRoundTripTextTransform(value, {
        targetKey: element.key,
        value: { height: 90, width: 320, x: 40, y: 50 },
      }),
    ).rejects.toThrow('PowerPoint transform target has no resolved transform');
  });

  it('rejects non-string text request fields', async () => {
    await expect(
      replacePptxRoundTripText(await snapshot(), {
        targetKey: 7 as never,
        value: 'After',
      }),
    ).rejects.toThrow(
      'PowerPoint text edit target key must be a non-empty string',
    );
    await expect(
      replacePptxRoundTripText(await snapshot(), {
        targetKey: 'slide-1-element-1-run-1',
        value: 7 as never,
      }),
    ).rejects.toThrow('PowerPoint text edit value must be a string');
  });

  it('enforces replace-text limits at exact code-unit and UTF-8 boundaries', () => {
    expect(() =>
      validatePptxRoundTripReplaceTextRequest(
        { targetKey: 'run', value: '1234' },
        4,
      ),
    ).not.toThrow();
    for (const value of ['12345', '😀😀']) {
      expect(() =>
        validatePptxRoundTripReplaceTextRequest({ targetKey: 'run', value }, 4),
      ).toThrow('PowerPoint text edit value exceeds the XML part byte limit');
    }
  });

  it.each([
    { height: 10, width: 20, x: '1', y: 2 },
    { height: 10, width: 20, x: Number.NaN, y: 2 },
    { height: 10, rotation: '1', width: 20, x: 1, y: 2 },
    { height: 10, rotation: Number.NaN, width: 20, x: 1, y: 2 },
    { height: 10, width: Number.MAX_VALUE, x: 1, y: 2 },
    { height: 10, rotation: Number.MAX_VALUE, width: 20, x: 1, y: 2 },
  ])('rejects invalid transform directly: %j', (value) => {
    expect(() => normalizePptxRoundTripTransform(value as never)).toThrow(
      'PowerPoint transform value is not a valid scene transform',
    );
  });

  it('normalizes optional transform fields directly', () => {
    expect(
      normalizePptxRoundTripTransform({
        height: 10,
        width: 20,
        x: 1,
        y: 2,
      }),
    ).toEqual({
      flipHorizontal: false,
      flipVertical: false,
      height: 10,
      rotation: 0,
      width: 20,
      x: 1,
      y: 2,
    });
  });

  it('binds, previews, and removes native image crops', async () => {
    const crop = { bottom: -20, left: 30, right: 0, top: 10 };
    const added = await setPptxRoundTripImageCrop(await imageSnapshot(), {
      targetKey: 'slide-1-element-1',
      value: crop,
    });
    expect(added.operations).toEqual([
      {
        expectedCrop: null,
        id: 'set-image-crop-1',
        kind: 'set-image-crop',
        targetKey: 'slide-1-element-1',
        value: crop,
      },
    ]);
    expect(added.supportProfile.id).toBe('pptx-roundtrip-native-v1');
    expect(
      applyPptxRoundTripOperationsToPreview(added).slides[0]?.elements[0],
    ).toMatchObject({ crop, type: 'image' });

    const removed = await setPptxRoundTripImageCrop(await imageSnapshot(true), {
      targetKey: 'slide-1-element-1',
      value: null,
    });
    expect(removed.operations).toMatchObject([
      { expectedCrop: crop, kind: 'set-image-crop', value: null },
    ]);
    expect(
      applyPptxRoundTripOperationsToPreview(removed).slides[0]?.elements[0],
    ).not.toHaveProperty('crop');
  });

  it.each([
    [undefined, 'must be an object or null'],
    [[], 'must be an object or null'],
    [{ bottom: 0, left: 0, right: 0 }, 'has an invalid shape'],
    [
      { bottom: 0, extra: 0, left: 0, right: 0, top: 0 },
      'has an invalid shape',
    ],
    [
      { bottom: 0, left: 0.0001, right: 0, top: 0 },
      'has an invalid percentage',
    ],
    [
      { bottom: 0, left: 100.001, right: 0, top: 0 },
      'has an invalid percentage',
    ],
    [
      { bottom: 50, left: 0, right: 0, top: 50 },
      'must leave a positive visible region',
    ],
  ])('rejects invalid image crop %j', (value, message) => {
    expect(() => normalizePptxRoundTripImageCrop(value as never)).toThrow(
      message,
    );
  });

  it('rejects missing, no-op, and duplicate image crop targets', async () => {
    await expect(
      setPptxRoundTripImageCrop(await imageSnapshot(), {
        targetKey: 'missing',
        value: { bottom: 0, left: 10, right: 0, top: 0 },
      }),
    ).rejects.toThrow('PowerPoint image transform target key does not exist');
    await expect(
      setPptxRoundTripImageCrop(await imageSnapshot(), {
        targetKey: 'slide-1-element-1',
        value: null,
      }),
    ).rejects.toThrow(
      'PowerPoint image crop edit must change the target value',
    );
    const first = await setPptxRoundTripImageCrop(await imageSnapshot(), {
      targetKey: 'slide-1-element-1',
      value: { bottom: 0, left: 10, right: 0, top: 0 },
    });
    await expect(
      setPptxRoundTripImageCrop(first, {
        targetKey: 'slide-1-element-1',
        value: { bottom: 0, left: 20, right: 0, top: 0 },
      }),
    ).rejects.toThrow('PowerPoint image crop target is already scheduled');
  });

  it.each([null, [], 7])('rejects group transform value %j', (value) => {
    expect(() => normalizePptxRoundTripGroupTransform(value as never)).toThrow(
      'PowerPoint group transform value must be an object',
    );
  });

  it('rejects unknown group transform fields', () => {
    expect(() =>
      normalizePptxRoundTripGroupTransform({
        childSpace: { height: 10, width: 20, x: 1, y: 2 },
        extra: true,
        height: 30,
        width: 40,
        x: 3,
        y: 4,
      } as never),
    ).toThrow('PowerPoint group transform value is not valid');
  });

  it.each([
    null,
    [],
    { extra: true, height: 10, width: 20, x: 1, y: 2 },
    { height: 10, width: 0, x: 1, y: 2 },
    { height: 0, width: 20, x: 1, y: 2 },
    { height: 10, width: 20, x: Number.NaN, y: 2 },
    { height: 10, width: 20, x: 1, y: Number.MAX_VALUE },
  ])('rejects invalid group child space %j', (childSpace) => {
    expect(() =>
      normalizePptxRoundTripGroupTransform({
        childSpace,
        height: 30,
        width: 40,
        x: 3,
        y: 4,
      } as never),
    ).toThrow(/PowerPoint group child space/);
  });

  it.each([7, 'bad'])(
    'rejects primitive group child space %j',
    (childSpace) => {
      expect(() =>
        normalizePptxRoundTripGroupTransform({
          childSpace,
          height: 30,
          width: 40,
          x: 3,
          y: 4,
        } as never),
      ).toThrow('PowerPoint group child space must be an object');
    },
  );

  it('normalizes explicit and omitted group optional fields', () => {
    expect(
      normalizePptxRoundTripGroupTransform({
        childSpace: { height: 10, width: 20, x: 1, y: 2 },
        flipHorizontal: true,
        flipVertical: true,
        height: 30,
        rotation: -90,
        width: 40,
        x: 3,
        y: 4,
      }),
    ).toEqual({
      childSpace: { height: 10, width: 20, x: 1, y: 2 },
      flipHorizontal: true,
      flipVertical: true,
      height: 30,
      rotation: -90,
      width: 40,
      x: 3,
      y: 4,
    });
    expect(
      normalizePptxRoundTripGroupTransform({
        childSpace: { height: 10, width: 20, x: 1, y: 2 },
        height: 30,
        width: 40,
        x: 3,
        y: 4,
      }),
    ).toMatchObject({
      flipHorizontal: false,
      flipVertical: false,
      rotation: 0,
    });
  });
});
