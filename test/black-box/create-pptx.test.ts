import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { decodeBase64 } from '../../src/common/binary/base64';
import {
  createPptx,
  parsePptx,
  PptxWriteError,
  renderPptxToSvg,
  type PptxElement,
  type PptxFidelityLevel,
  type PptxSceneDocument,
  type PptxSceneSlide,
} from '../../src';

function textContent(element: PptxElement | undefined): string {
  if (!element || !('content' in element)) {
    throw new Error('Expected a generated text-bearing element');
  }
  return element.content;
}

function slide(key: string, text: string): PptxSceneSlide {
  return {
    elements: [
      {
        authored: {
          transform: { height: 80, width: 300, x: 20, y: 30 },
        },
        key: `${key}-text`,
        name: `${key} text`,
        resolved: { hidden: false },
        text: {
          body: { anchor: 'center', wrap: true },
          paragraphs: [
            {
              children: [
                {
                  key: `${key}-run`,
                  properties: { bold: true, fontSize: 18 },
                  text,
                  type: 'run',
                },
              ],
              key: `${key}-paragraph`,
            },
          ],
        },
        type: 'text',
      },
    ],
    key,
    name: `${key} name`,
  };
}

function creationScene(): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [slide('first', 'Hello <& world'), slide('second', 'Slide two')],
    themes: [],
  };
}

const PNG_BYTES = decodeBase64(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
);

describe('PowerPoint creation through the public API', () => {
  it('exposes producer-verified creation as a distinct fidelity level', () => {
    const level: PptxFidelityLevel = 'C3';

    expect(level).toBe('C3');
  });

  it('creates a strict-readable and Office-free rendered text presentation with an explicit C2 report', async () => {
    const result = await createPptx(creationScene());
    const parsed = await parsePptx(result.data, {
      audioMode: 'none',
      errorMode: 'strict',
      imageMode: 'none',
      videoMode: 'none',
    });

    expect(Array.from(result.data.slice(0, 4))).toEqual([
      0x50, 0x4b, 0x03, 0x04,
    ]);
    expect(parsed.size).toEqual({ height: 540, width: 960 });
    expect(parsed.slides).toHaveLength(2);
    expect(parsed.slides.map((value) => value.elements.length)).toEqual([1, 1]);
    expect(textContent(parsed.slides[0]?.elements[0])).toContain(
      'Hello&nbsp;&lt;&amp;&nbsp;world',
    );
    expect(textContent(parsed.slides[1]?.elements[0])).toContain(
      'Slide&nbsp;two',
    );
    expect(result.report).toEqual({
      addedPartCount: 13,
      copiedPartCount: 0,
      diagnostics: [],
      level: 'C2',
      operations: [],
      patchedPartCount: 0,
      producerEvidence: [],
      rebuiltPartCount: 0,
      removedPartCount: 0,
      supportProfile: {
        effectiveLevel: 'C2',
        id: 'pptx-create-text-v1',
        producerMatrix: [],
        version: '1',
      },
    });
  });

  it('is byte deterministic across concurrent public calls', async () => {
    const input = creationScene();
    const [first, second, third] = await Promise.all([
      createPptx(input),
      createPptx(input),
      createPptx(input),
    ]);

    expect(second.data).toEqual(first.data);
    expect(third.data).toEqual(first.data);
    expect(second.report).toEqual(first.report);
    expect(third.report).toEqual(first.report);
  });

  it('strict-parses and Office-free renders authored visual styling', async () => {
    const scene = creationScene();
    const first = scene.slides[0];
    if (first === undefined) throw new Error('Expected first slide');
    first.backgroundColor = '#0F172A';
    const firstElement = first.elements[0];
    if (firstElement?.type !== 'text') throw new Error('Expected text element');
    firstElement.authored.fillColor = '#1E293B';
    firstElement.authored.geometry = 'roundRect';
    firstElement.authored.lineColor = '#38BDF8';
    firstElement.authored.lineWidth = 1.5;
    const firstRun = firstElement.text.paragraphs[0]?.children[0];
    if (firstRun?.type !== 'run') throw new Error('Expected text run');
    firstRun.properties = {
      ...firstRun.properties,
      color: '#F8FAFC',
      fontFamily: 'Aptos Display',
    };

    const created = await createPptx(scene);
    const [parsed, rendered] = await Promise.all([
      parsePptx(created.data, { errorMode: 'strict', imageMode: 'none' }),
      renderPptxToSvg(created.data, { slideNumbers: [1] }),
    ]);
    expect(parsed.slides[0]?.fill).toEqual({
      type: 'color',
      value: '#0F172A',
    });
    expect(parsed.slides[0]?.elements[0]).toMatchObject({
      borderColor: '#38BDF8',
      borderWidth: 1.5,
      fill: { type: 'color', value: '#1E293B' },
      shapType: 'roundRect',
      type: 'shape',
    });
    const svg = new TextDecoder()
      .decode(rendered.slides[0]?.data)
      .toUpperCase();
    expect(svg).toContain('#0F172A');
    expect(svg).toContain('#1E293B');
    expect(svg).toContain('#38BDF8');
    expect(svg).toContain('#F8FAFC');
    expect(svg).toContain('HELLO');
  });

  it('creates a native non-text shape and reports its distinct profile', async () => {
    const scene = creationScene();
    const first = scene.slides[0];
    if (first === undefined) throw new Error('Expected first slide');
    first.elements.push({
      authored: {
        fillColor: '#F97316',
        geometry: 'ellipse',
        lineColor: '#0F172A',
        lineWidth: 2,
        transform: {
          flipHorizontal: true,
          height: 120,
          rotation: 12,
          width: 180,
          x: 420,
          y: 220,
        },
      },
      description: 'Native decorative shape',
      key: 'native-shape',
      name: 'Native ellipse',
      resolved: { hidden: false },
      type: 'shape',
    });

    const created = await createPptx(scene);
    const [parsed, rendered] = await Promise.all([
      parsePptx(created.data, { errorMode: 'strict', imageMode: 'none' }),
      renderPptxToSvg(created.data, { slideNumbers: [1] }),
    ]);

    expect(created.report.supportProfile).toEqual({
      effectiveLevel: 'C2',
      id: 'pptx-create-native-v1',
      producerMatrix: [],
      version: '1',
    });
    expect(parsed.slides[0]?.elements[1]).toMatchObject({
      borderColor: '#0F172A',
      borderWidth: 2,
      fill: { type: 'color', value: '#F97316' },
      height: 120,
      isFlipH: true,
      left: 420,
      rotate: 12,
      shapType: 'ellipse',
      top: 220,
      type: 'shape',
      width: 180,
    });
    const svg = new TextDecoder()
      .decode(rendered.slides[0]?.data)
      .toUpperCase();
    expect(svg).toContain('#F97316');
    expect(svg).toContain('#0F172A');
  });

  it('creates a native image with exact media, relationships, and rendering', async () => {
    const scene = creationScene();
    scene.media = [
      {
        data: PNG_BYTES,
        key: 'logo-media',
        mimeType: 'image/png',
      },
    ];
    scene.slides[0]?.elements.push({
      authored: {
        transform: {
          flipVertical: true,
          height: 90,
          rotation: 5,
          width: 120,
          x: 500,
          y: 300,
        },
      },
      crop: { bottom: -20, left: 30, right: 0, top: 10.125 },
      description: 'Native image',
      key: 'logo-picture',
      mediaKey: 'logo-media',
      name: 'Logo picture',
      resolved: { hidden: false },
      type: 'image',
    });

    const created = await createPptx(scene);
    const [archive, parsed, rendered] = await Promise.all([
      JSZip.loadAsync(created.data),
      parsePptx(created.data, { errorMode: 'strict', imageMode: 'base64' }),
      renderPptxToSvg(created.data, { slideNumbers: [1] }),
    ]);

    expect(created.report.supportProfile.id).toBe('pptx-create-native-v1');
    await expect(
      archive.file('ppt/media/image1.png')?.async('uint8array'),
    ).resolves.toEqual(PNG_BYTES);
    await expect(
      archive.file('[Content_Types].xml')?.async('string'),
    ).resolves.toContain('<Default Extension="png" ContentType="image/png"/>');
    await expect(
      archive.file('ppt/slides/_rels/slide1.xml.rels')?.async('string'),
    ).resolves.toContain(
      'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"',
    );
    await expect(
      archive.file('ppt/slides/slide1.xml')?.async('string'),
    ).resolves.toContain('<a:srcRect l="30000" t="10125" r="0" b="-20000"/>');
    const parsedImage = parsed.slides[0]?.elements[1];
    expect(parsedImage).toMatchObject({
      height: 90,
      isFlipV: true,
      left: 500,
      rect: { b: -20, l: 30, r: 0, t: 10.125 },
      rotate: 5,
      top: 300,
      type: 'image',
      width: 120,
    });
    if (parsedImage?.type !== 'image') throw new Error('Expected image');
    expect(parsedImage.base64).toMatch(/^data:image\/png;base64,/);
    expect(new TextDecoder().decode(rendered.slides[0]?.data)).toContain(
      'data:image/png;base64,',
    );
  });

  it('owns native media bytes before asynchronous archive generation', async () => {
    const scene = creationScene();
    const callerBytes = new Uint8Array(PNG_BYTES);
    scene.media = [{ data: callerBytes, key: 'image', mimeType: 'image/png' }];
    scene.slides[0]?.elements.push({
      authored: { transform: { height: 10, width: 10, x: 0, y: 0 } },
      key: 'picture',
      mediaKey: 'image',
      resolved: { hidden: false },
      type: 'image',
    });

    const creating = createPptx(scene);
    callerBytes.fill(0);
    const created = await creating;
    const archive = await JSZip.loadAsync(created.data);

    await expect(
      archive.file('ppt/media/image1.png')?.async('uint8array'),
    ).resolves.toEqual(PNG_BYTES);
  });

  it('contains only the declared deterministic package inventory', async () => {
    const result = await createPptx(creationScene());
    const archive = await JSZip.loadAsync(result.data);

    expect(Object.keys(archive.files)).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'ppt/presentation.xml',
      'ppt/_rels/presentation.xml.rels',
      'ppt/slideMasters/slideMaster1.xml',
      'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      'ppt/slideLayouts/slideLayout1.xml',
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
      'ppt/theme/theme1.xml',
      'ppt/slides/slide1.xml',
      'ppt/slides/_rels/slide1.xml.rels',
      'ppt/slides/slide2.xml',
      'ppt/slides/_rels/slide2.xml.rels',
    ]);
  });

  it('rejects invalid input before returning package bytes', async () => {
    const input = creationScene();
    input.slides[0]?.elements.splice(0, 1);
    const invalid = input as unknown as Record<string, unknown>;
    invalid.schemaVersion = 1;

    const promise = createPptx(input);
    await expect(promise).rejects.toBeInstanceOf(PptxWriteError);
    await expect(promise).rejects.toMatchObject({
      code: 'invalid-scene',
      issues: [
        {
          code: 'unsupported-schema-version',
          message: 'Only PowerPoint scene schema version 2 is supported',
          path: '$.schemaVersion',
        },
      ],
      message: 'PowerPoint scene is not valid for creation',
    });
    await expect(promise).rejects.not.toHaveProperty('data');
  });
});
