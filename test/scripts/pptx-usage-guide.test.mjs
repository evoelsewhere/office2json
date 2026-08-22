import { readFile } from 'node:fs/promises';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

import * as rootApi from '../../src/index.ts';
import * as nodeApi from '../../src/formats/pptx/node.ts';

const guidePath = path.resolve('docs', 'pptx-usage-guide.md');

describe('PowerPoint usage guide', () => {
  it('documents exported end-to-end APIs and honest capability boundaries', async () => {
    const [guide, metadata, readme] = await Promise.all([
      readFile(guidePath, 'utf8'),
      readFile(path.resolve('package.json'), 'utf8').then(JSON.parse),
      readFile(path.resolve('README.md'), 'utf8'),
    ]);

    for (const name of [
      'createPptx',
      'parsePptx',
      'parsePptxRoundTripJson',
      'parsePptxWithDiagnostics',
      'readPptxRoundTrip',
      'renderPptxDocumentToSvg',
      'renderPptxToSvg',
      'replacePptxRoundTripText',
      'serializePptxRoundTripJson',
      'setPptxRoundTripTextTransform',
      'setPptxRoundTripShapeTransform',
      'setPptxRoundTripImageTransform',
      'setPptxRoundTripImageCrop',
      'setPptxRoundTripTableTransform',
      'setPptxRoundTripGroupTransform',
      'setPptxRoundTripChartTransform',
      'validatePptxScene',
      'writePptxRoundTrip',
    ]) {
      expect(rootApi[name], name).toBeTypeOf('function');
      expect(guide, name).toContain(name);
    }
    for (const name of ['renderPptxDocumentToPng', 'renderPptxToPng']) {
      expect(nodeApi[name], name).toBeTypeOf('function');
      expect(guide, name).toContain(name);
    }

    expect(guide).toContain(`OAKit \`${metadata.version}\``);
    expect(guide).toMatch(
      /\|\s*Arbitrary PPTX creation\/editing\s*\|\s*Not claimed\s*\|/,
    );
    expect(guide).toMatch(
      /\|\s*Pixel-identical rendering\s*\|\s*Not claimed\s*\|/,
    );
    expect(guide).toContain('Never edit `snapshot.document`');
    expect(guide).toContain('error.diagnostic.code');
    expect(guide).toContain('## Five-minute verified workflow');
    expect(guide).toContain("profile: 'create-native-v1'");
    expect(guide).toContain("operation.status !== 'verified'");
    expect(guide).toContain('function* walkElements');
    expect(guide).toMatch(/at most 100,000 aggregate chart\s+points/);
    expect(guide).not.toContain('editableScene');
    expect(guide).toContain('oakit --version');
    expect(readme).toContain(
      '[PowerPoint usage guide](docs/pptx-usage-guide.md)',
    );
    expect((guide.match(/^```/gm)?.length ?? 0) % 2).toBe(0);
  });

  it('keeps the verified native quick start executable', async () => {
    const scene = {
      schemaVersion: 2,
      size: { width: 960, height: 540 },
      themes: [],
      masters: [],
      layouts: [],
      media: [],
      slides: [
        {
          key: 'overview-slide',
          elements: [
            {
              type: 'text',
              key: 'overview-title',
              authored: {
                transform: { x: 56, y: 42, width: 848, height: 72 },
              },
              resolved: { hidden: false },
              text: {
                body: { anchor: 'center', wrap: true },
                paragraphs: [
                  {
                    key: 'overview-title-paragraph',
                    children: [
                      {
                        type: 'run',
                        key: 'overview-title-run',
                        text: 'Quarterly review',
                      },
                    ],
                  },
                ],
              },
            },
            {
              type: 'shape',
              key: 'accent',
              authored: {
                fillColor: '#4F46E5',
                geometry: 'roundRect',
                transform: { x: 56, y: 132, width: 240, height: 286 },
              },
              resolved: { hidden: false },
            },
            {
              type: 'chart',
              key: 'revenue-chart',
              authored: {
                transform: { x: 328, y: 132, width: 576, height: 286 },
              },
              resolved: { hidden: false },
              chartType: 'barChart',
              barDirection: 'col',
              grouping: 'clustered',
              series: [
                {
                  key: 'revenue-series',
                  name: 'Revenue',
                  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
                  values: [18, 24, 31, 43],
                  color: '#4F46E5',
                },
              ],
            },
          ],
        },
      ],
    };

    expect(
      rootApi.validatePptxScene(scene, { profile: 'create-native-v1' }),
    ).toMatchObject({ valid: true, issues: [] });

    const created = await rootApi.createPptx(scene);
    await expect(
      rootApi.parsePptx(created.data, {
        errorMode: 'strict',
        imageMode: 'none',
      }),
    ).resolves.toMatchObject({ slides: [{ elements: expect.any(Array) }] });
    await expect(
      nodeApi.renderPptxToPng(created.data, { slideNumbers: [1] }),
    ).resolves.toMatchObject({ slides: [{ format: 'png', slideNumber: 1 }] });

    const snapshot = await rootApi.readPptxRoundTrip(created.data);
    const titleRun = snapshot.document.slides
      .flatMap((slide) => slide.elements)
      .filter((element) => element.type === 'text')
      .flatMap((element) => element.text.paragraphs)
      .flatMap((paragraph) => paragraph.children)
      .find(
        (child) => child.type === 'run' && child.text === 'Quarterly review',
      );
    expect(titleRun).toBeDefined();

    const edited = await rootApi.replacePptxRoundTripText(snapshot, {
      targetKey: titleRun.key,
      value: 'Quarterly review — approved',
    });
    const written = await rootApi.writePptxRoundTrip(edited);

    expect(written.report.level).toBe('R2');
    expect(written.report.operations).toEqual([
      expect.objectContaining({ kind: 'replace-text', status: 'verified' }),
    ]);
    await expect(
      rootApi.parsePptx(written.data, {
        errorMode: 'strict',
        imageMode: 'none',
      }),
    ).resolves.toBeDefined();
  }, 30_000);

  it('keeps every relative guide link resolvable', async () => {
    const guide = await readFile(guidePath, 'utf8');
    const links = [...guide.matchAll(/\]\(([^)]+)\)/g)]
      .map((match) => match[1])
      .filter(
        (target) =>
          !target.startsWith('#') &&
          !target.startsWith('http://') &&
          !target.startsWith('https://'),
      );

    expect(links.length).toBeGreaterThan(0);
    for (const target of links) {
      await expect(
        readFile(path.resolve(path.dirname(guidePath), target)),
        target,
      ).resolves.toBeDefined();
    }
  });
});
