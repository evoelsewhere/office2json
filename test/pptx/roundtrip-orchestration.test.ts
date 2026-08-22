import { describe, expect, it } from 'vitest';

import { resolvePptxResourceLimits } from '../../src/formats/pptx/internal/resource-limits';
import { patchPptxOperations } from '../../src/formats/pptx/roundtrip/orchestration';
import type { PptxRoundTripOperation } from '../../src/formats/pptx/roundtrip/types';
import { parse } from '../../src/formats/pptx/parser';
import { createMinimalPptx } from './fixture';

describe('PowerPoint patch orchestration', () => {
  it('rejects a parsed document with a different slide count', async () => {
    const data = await createMinimalPptx();
    const document = await parse(data, { imageMode: 'none' });
    document.slides = [];

    await expect(
      patchPptxOperations(data, document, [], resolvePptxResourceLimits()),
    ).rejects.toThrow(
      'PowerPoint text edit slide order does not match the parsed document',
    );
  });

  it('rejects unsafe operation indices and non-text document targets', async () => {
    const data = await createMinimalPptx();
    const document = await parse(data, { imageMode: 'none' });
    const slide = document.slides[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements = [
      {
        base64: '',
        blob: '',
        borderColor: '#000000',
        borderStrokeDasharray: '0',
        borderType: 'solid',
        borderWidth: 0,
        geom: 'rect',
        height: 10,
        id: '2',
        isFlipH: false,
        isFlipV: false,
        left: 1,
        order: 0,
        ref: 'image.png',
        rotate: 0,
        top: 2,
        type: 'image',
        width: 10,
      },
    ];
    const replaceOperation: PptxRoundTripOperation = {
      expectedText: 'Before',
      id: 'replace-text-1',
      kind: 'replace-text',
      targetKey: 'slide-1-element-1-run-1',
      value: 'After',
    };
    await expect(
      patchPptxOperations(
        data,
        document,
        [replaceOperation],
        resolvePptxResourceLimits(),
      ),
    ).rejects.toThrow(
      'PowerPoint text edit target is not a native text element',
    );
    (document.slides[0]?.elements[0] as unknown as { type: string }).type =
      'audio';

    const transformOperation: PptxRoundTripOperation = {
      expectedTransform: {
        flipHorizontal: false,
        flipVertical: false,
        height: 10,
        rotation: 0,
        width: 10,
        x: 1,
        y: 2,
      },
      id: 'set-transform-1',
      kind: 'set-transform',
      targetKey: 'slide-1-element-1',
      value: {
        flipHorizontal: false,
        flipVertical: false,
        height: 20,
        rotation: 0,
        width: 20,
        x: 3,
        y: 4,
      },
    };
    await expect(
      patchPptxOperations(
        data,
        document,
        [transformOperation],
        resolvePptxResourceLimits(),
      ),
    ).rejects.toThrow(
      'PowerPoint transform target is not a slide-owned text, shape, image, table, chart, or group element',
    );

    (document.slides[0]?.elements[0] as unknown as { type: string }).type =
      'shape';
    const cropOperation: PptxRoundTripOperation = {
      expectedCrop: null,
      id: 'set-image-crop-1',
      kind: 'set-image-crop',
      targetKey: 'slide-1-element-1',
      value: { bottom: 0, left: 10, right: 0, top: 0 },
    };
    await expect(
      patchPptxOperations(
        data,
        document,
        [cropOperation],
        resolvePptxResourceLimits(),
      ),
    ).rejects.toThrow(
      'PowerPoint image crop target is not a native image element',
    );

    transformOperation.targetKey = `slide-${'9'.repeat(20)}-element-1`;
    await expect(
      patchPptxOperations(
        data,
        document,
        [transformOperation],
        resolvePptxResourceLimits(),
      ),
    ).rejects.toThrow('PowerPoint transform target index is unsafe');

    for (const malformed of [
      { ...replaceOperation, targetKey: 'bad-run-key' },
      {
        ...replaceOperation,
        targetKey: 'xslide-1-element-1-run-1',
      },
      {
        ...replaceOperation,
        targetKey: 'slide-1-element-1-run-1x',
      },
      { ...transformOperation, targetKey: 'bad-element-key' },
      { ...transformOperation, targetKey: 'xslide-1-element-1' },
      { ...transformOperation, targetKey: 'slide-1-element-1x' },
    ]) {
      await expect(
        patchPptxOperations(
          data,
          document,
          [malformed],
          resolvePptxResourceLimits(),
        ),
      ).rejects.toThrow(/target is not a supported (?:native|slide) text/);
    }

    for (const targetKey of [
      'xslide-1-element-1-row-1-cell-1-run-1',
      'slide-1-element-1-row-1-cell-1-run-1x',
      'slide-a-element-1-row-1-cell-1-run-1',
      'slide-1-element-a-row-1-cell-1-run-1',
      'slide-1-element-1-row-a-cell-1-run-1',
      'slide-1-element-1-row-1-cell-a-run-1',
      'slide-1-element-1-row-0-cell-1-run-1',
      'slide-1-element-1-row-1-cell-0-run-1',
    ]) {
      await expect(
        patchPptxOperations(
          data,
          document,
          [{ ...replaceOperation, targetKey }],
          resolvePptxResourceLimits(),
        ),
      ).rejects.toThrow(
        'PowerPoint text edit target is not a supported native text run key',
      );
    }

    for (const targetKey of [
      'slide-10-element-1-row-1-cell-1-run-1',
      'slide-1-element-10-row-1-cell-1-run-1',
      'slide-1-element-1-row-10-cell-1-run-1',
      'slide-1-element-1-row-1-cell-10-run-1',
    ]) {
      await expect(
        patchPptxOperations(
          data,
          document,
          [{ ...replaceOperation, targetKey }],
          resolvePptxResourceLimits(),
        ),
      ).rejects.toThrow(
        'PowerPoint table text edit target is not a native table element',
      );
    }

    for (const targetKey of [
      `slide-${'9'.repeat(20)}-element-1-run-1`,
      `slide-1-element-${'9'.repeat(20)}-run-1`,
      `slide-1-element-1-row-${'9'.repeat(20)}-cell-1-run-1`,
      `slide-1-element-1-row-1-cell-${'9'.repeat(20)}-run-1`,
    ]) {
      await expect(
        patchPptxOperations(
          data,
          document,
          [{ ...replaceOperation, targetKey }],
          resolvePptxResourceLimits(),
        ),
      ).rejects.toThrow('PowerPoint text edit target index is unsafe');
    }
  });

  it('rejects invalid nested transform ownership paths', async () => {
    const data = await createMinimalPptx();
    const document = await parse(data, { imageMode: 'none' });
    const operation: PptxRoundTripOperation = {
      expectedTransform: {
        flipHorizontal: false,
        flipVertical: false,
        height: 10,
        rotation: 0,
        width: 10,
        x: 1,
        y: 2,
      },
      id: 'set-transform-1',
      kind: 'set-transform',
      targetKey: 'slide-1-element-1-element-1',
      value: {
        flipHorizontal: false,
        flipVertical: false,
        height: 20,
        rotation: 0,
        width: 20,
        x: 3,
        y: 4,
      },
    };
    const child = {
      height: 10,
      id: '3',
      isFlipH: false,
      isFlipV: false,
      left: 1,
      order: 0,
      rotate: 0,
      top: 2,
      type: 'shape',
      width: 10,
    } as const;
    const group = {
      elements: [child],
      height: 20,
      id: '2',
      isFlipH: false,
      isFlipV: false,
      left: 0,
      order: 0,
      rotate: 0,
      top: 0,
      type: 'group',
      width: 20,
    } as const;
    const slide = document.slides[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements = [group as never];

    await expect(
      patchPptxOperations(
        data,
        document,
        [operation],
        resolvePptxResourceLimits(),
      ),
    ).rejects.toThrow(
      'PowerPoint nested transform ancestor has no child coordinate space',
    );

    const groupWithSpace = group as typeof group & {
      childSpace: { height: number; width: number; x: number; y: number };
    };
    groupWithSpace.childSpace = { height: 20, width: 20, x: 0, y: 0 };
    operation.targetKey = 'slide-1-element-1-element-1-element-1';
    await expect(
      patchPptxOperations(
        data,
        document,
        [operation],
        resolvePptxResourceLimits(),
      ),
    ).rejects.toThrow(
      'PowerPoint transform target path crosses a non-group element',
    );

    operation.targetKey = `slide-1-element-1-element-${'9'.repeat(20)}`;
    await expect(
      patchPptxOperations(
        data,
        document,
        [operation],
        resolvePptxResourceLimits(),
      ),
    ).rejects.toThrow('PowerPoint transform target index is unsafe');

    const textOperation: PptxRoundTripOperation = {
      expectedText: 'Before',
      id: 'replace-text-1',
      kind: 'replace-text',
      targetKey: `slide-1-element-1-element-${'9'.repeat(20)}-run-1`,
      value: 'After',
    };
    await expect(
      patchPptxOperations(
        data,
        document,
        [textOperation],
        resolvePptxResourceLimits(),
      ),
    ).rejects.toThrow('PowerPoint text edit target index is unsafe');

    textOperation.targetKey = 'slide-1-element-1-element-1-element-1-run-1';
    await expect(
      patchPptxOperations(
        data,
        document,
        [textOperation],
        resolvePptxResourceLimits(),
      ),
    ).rejects.toThrow(
      'PowerPoint text edit target path crosses a non-group element',
    );
  });
});
