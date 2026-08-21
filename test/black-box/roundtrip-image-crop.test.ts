import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { decodeBase64 } from '../../src/common/binary/base64';
import {
  createPptx,
  parsePptxRoundTripJson,
  readPptxRoundTrip,
  renderPptxToSvg,
  serializePptxRoundTripJson,
  setPptxRoundTripImageCrop,
  setPptxRoundTripImageTransform,
  writePptxRoundTrip,
  type PptxSceneDocument,
} from '../../src';

const PNG_BYTES = decodeBase64(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
);
const SOURCE_CROP = { bottom: -20, left: 30, right: 0, top: 10.125 };
const CHANGED_CROP = { bottom: 5, left: 10, right: 15, top: 20 };

function scene(cropped = false): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [{ data: PNG_BYTES, key: 'media', mimeType: 'image/png' }],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {
              transform: { height: 90, width: 120, x: 500, y: 300 },
            },
            ...(cropped ? { crop: SOURCE_CROP } : {}),
            key: 'picture',
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
}

async function payloads(data: Uint8Array): Promise<Map<string, Uint8Array>> {
  const archive = await JSZip.loadAsync(data);
  const result = new Map<string, Uint8Array>();
  for (const part of Object.values(archive.files)) {
    if (!part.dir) result.set(part.name, await part.async('uint8array'));
  }
  return result;
}

describe('native PowerPoint image crop editing', () => {
  it('adds crop and transform operations while preserving exact media bytes', async () => {
    const created = await createPptx(scene());
    const snapshot = await readPptxRoundTrip(created.data);
    const cropped = await setPptxRoundTripImageCrop(snapshot, {
      targetKey: 'slide-1-element-1',
      value: SOURCE_CROP,
    });
    const transform = {
      flipHorizontal: true,
      flipVertical: false,
      height: 130,
      rotation: 15,
      width: 170,
      x: 450,
      y: 260,
    };
    const edited = await setPptxRoundTripImageTransform(cropped, {
      targetKey: 'slide-1-element-1',
      value: transform,
    });
    const output = await writePptxRoundTrip(edited);
    const [sourceParts, outputParts, verified, rendered] = await Promise.all([
      payloads(created.data),
      payloads(output.data),
      readPptxRoundTrip(output.data),
      renderPptxToSvg(output.data, { slideNumbers: [1] }),
    ]);

    expect(output.report).toMatchObject({
      level: 'R2',
      patchedPartCount: 1,
      supportProfile: { id: 'pptx-roundtrip-native-v1' },
    });
    expect(output.report.operations).toMatchObject([
      { kind: 'set-image-crop', status: 'verified' },
      { kind: 'set-transform', status: 'verified' },
    ]);
    expect(verified.document.slides[0]?.elements[0]).toMatchObject({
      crop: SOURCE_CROP,
      resolved: { transform },
      type: 'image',
    });
    for (const [name, source] of sourceParts) {
      const result = outputParts.get(name);
      expect(result, name).toBeDefined();
      if (name === 'ppt/slides/slide1.xml') {
        expect(result).not.toEqual(source);
      } else {
        expect(result, name).toEqual(source);
      }
    }
    expect(outputParts.get('ppt/media/image1.png')).toEqual(PNG_BYTES);
    const svg = new TextDecoder().decode(rendered.slides[0]?.data);
    expect(svg).toContain('overflow="hidden"><image x="-');
    expect(svg).toContain('data:image/png;base64,');
  });

  it('replaces crop through portable JSON and removes it explicitly', async () => {
    const created = await createPptx(scene(true));
    const snapshot = await readPptxRoundTrip(created.data);
    const changed = await setPptxRoundTripImageCrop(snapshot, {
      targetKey: 'slide-1-element-1',
      value: CHANGED_CROP,
    });
    const portable = await serializePptxRoundTripJson(changed);
    const restored = await parsePptxRoundTripJson(
      JSON.parse(JSON.stringify(portable)),
    );
    const changedOutput = await writePptxRoundTrip(restored);
    expect(
      (await readPptxRoundTrip(changedOutput.data)).document.slides[0]
        ?.elements[0],
    ).toMatchObject({ crop: CHANGED_CROP, type: 'image' });

    const removed = await setPptxRoundTripImageCrop(snapshot, {
      targetKey: 'slide-1-element-1',
      value: null,
    });
    const removedOutput = await writePptxRoundTrip(removed);
    expect(
      (await readPptxRoundTrip(removedOutput.data)).document.slides[0]
        ?.elements[0],
    ).not.toHaveProperty('crop');
  });

  it('is deterministic across independent crop writes', async () => {
    const created = await createPptx(scene());
    const snapshot = await readPptxRoundTrip(created.data);
    const [first, second] = await Promise.all(
      [0, 1].map(async () => {
        const edited = await setPptxRoundTripImageCrop(snapshot, {
          targetKey: 'slide-1-element-1',
          value: SOURCE_CROP,
        });
        return writePptxRoundTrip(edited);
      }),
    );

    expect(second).toEqual(first);
  });
});
