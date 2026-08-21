import { RATIO_EMUs_Points } from '../../../common/ooxml/units';
import { encodeBase64 } from '../../../common/binary/base64';
import { renderPptxDocumentToSvg } from '../render-svg';
import type { PptxSvgRenderResult } from '../render-types';
import type {
  PptxSceneDocument,
  PptxSceneChartElement,
  PptxSceneElement,
  PptxSceneImageElement,
  PptxSceneMedia,
  PptxSceneShapeElement,
  PptxSceneTextElement,
  PptxSceneTextNode,
} from '../scene-types';
import { parse } from '../parser';
import type {
  Chart,
  Image,
  Group,
  PptxDocument,
  PptxParseOptions,
  Shape,
  Table,
  Text,
} from '../types';
import { plainTextFromPowerPointHtml } from '../roundtrip/preview';
import { verifyPowerPointTableElement } from './table-verify';
import { verifyPowerPointGroupElement } from './group-verify';
import { pointsToEmu } from './units';

type PptxCreationParser = (
  data: Uint8Array,
  options: PptxParseOptions,
) => Promise<PptxDocument>;

type PptxCreationRenderer = (document: PptxDocument) => PptxSvgRenderResult;

function countSceneCharts(elements: readonly PptxSceneElement[]): number {
  let count = 0;
  for (const element of elements) {
    if (element.type === 'chart') count += 1;
    if (element.type === 'group') count += countSceneCharts(element.elements);
  }
  return count;
}

function expectedPointValue(value: number): number {
  return pointsToEmu(value) * RATIO_EMUs_Points;
}

function textNodeValue(node: PptxSceneTextNode): string {
  return node.type === 'break' ? '\n' : node.text;
}

function expectedPlainText(element: PptxSceneTextElement): string {
  return element.text.paragraphs
    .map((paragraph) => paragraph.children.map(textNodeValue).join(''))
    .join('\n');
}

function generatedTextElement(
  generated: PptxDocument['slides'][number]['elements'][number] | undefined,
  geometry: NonNullable<PptxSceneTextElement['authored']['geometry']>,
  location: string,
): Shape | Text {
  if (generated === undefined) {
    throw new Error(`Generated PowerPoint text element missing at ${location}`);
  }
  if (geometry === 'rect') {
    if (generated.type === 'text') return generated;
  } else if (generated.type === 'shape' && generated.shapType === geometry) {
    return generated;
  }
  throw new Error(`Generated PowerPoint text element missing at ${location}`);
}

function verifyTextElement(
  generated: PptxDocument['slides'][number]['elements'][number] | undefined,
  expected: PptxSceneTextElement,
  slideIndex: number,
  elementIndex: number,
): void {
  const location = `slide ${slideIndex + 1}, element ${elementIndex + 1}`;
  const geometry = expected.authored.geometry ?? 'rect';
  const textElement = generatedTextElement(generated, geometry, location);
  verifyElementTransform(textElement, expected, location);
  const actualText = plainTextFromPowerPointHtml(textElement.content);
  if (actualText !== expectedPlainText(expected)) {
    throw new Error(`Generated PowerPoint text mismatch at ${location}`);
  }
}

function verifyElementTransform(
  generated: Group | Image | Shape | Table | Text,
  expected: PptxSceneElement,
  location: string,
): void {
  const transform = expected.authored.transform;
  if (transform === undefined) {
    throw new Error(
      `Expected PowerPoint authored transform missing at ${location}`,
    );
  }
  const generatedTransform = {
    flipHorizontal: generated.isFlipH,
    flipVertical: generated.isFlipV,
    height: generated.height,
    rotation: generated.rotate,
    width: generated.width,
    x: generated.left,
    y: generated.top,
  };
  const expectedTransform = {
    flipHorizontal: transform.flipHorizontal ?? false,
    flipVertical: transform.flipVertical ?? false,
    height: expectedPointValue(transform.height),
    rotation: transform.rotation ?? 0,
    width: expectedPointValue(transform.width),
    x: expectedPointValue(transform.x),
    y: expectedPointValue(transform.y),
  };
  if (
    JSON.stringify(generatedTransform) !== JSON.stringify(expectedTransform)
  ) {
    throw new Error(`Generated PowerPoint transform mismatch at ${location}`);
  }
}

const TABLE_VERIFICATION_DEPENDENCIES = {
  expectedPointValue,
  plainText: plainTextFromPowerPointHtml,
  textNodeValue,
  verifyTransform: verifyElementTransform,
};

function verifySceneElement(
  generated: PptxDocument['slides'][number]['elements'][number] | undefined,
  expected: PptxSceneElement,
  scene: PptxSceneDocument,
  slideIndex: number,
  elementIndex: number,
): void {
  if (expected.type === 'text') {
    verifyTextElement(generated, expected, slideIndex, elementIndex);
  } else if (expected.type === 'shape') {
    verifyShapeElement(generated, expected, slideIndex, elementIndex);
  } else if (expected.type === 'chart') {
    verifyChartElement(generated, expected, slideIndex, elementIndex);
  } else if (expected.type === 'image') {
    verifyImageElement(
      generated,
      expected,
      scene.media.find((media) => media.key === expected.mediaKey),
      slideIndex,
      elementIndex,
    );
  } else if (expected.type === 'table') {
    verifyPowerPointTableElement(
      generated,
      expected,
      slideIndex,
      elementIndex,
      TABLE_VERIFICATION_DEPENDENCIES,
    );
  } else if (expected.type === 'group') {
    const location = `slide ${slideIndex + 1}, element ${elementIndex + 1}`;
    verifyPowerPointGroupElement(generated, expected, location, {
      expectedPointValue,
      verifyChild: (generatedChild, expectedChild, childIndex) =>
        verifySceneElement(
          generatedChild,
          expectedChild,
          scene,
          slideIndex,
          childIndex,
        ),
      verifyTransform: verifyElementTransform,
    });
  } else {
    throw new Error(
      `Expected PowerPoint text element missing at slide ${slideIndex + 1}, element ${elementIndex + 1}`,
    );
  }
}

function verifyChartElement(
  generated: PptxDocument['slides'][number]['elements'][number] | undefined,
  expected: PptxSceneChartElement,
  slideIndex: number,
  elementIndex: number,
): void {
  const location = `slide ${slideIndex + 1}, element ${elementIndex + 1}`;
  if (generated?.type !== 'chart') {
    throw new Error(`Generated PowerPoint chart missing at ${location}`);
  }
  const transform = expected.authored.transform;
  if (transform === undefined) {
    throw new Error(
      `Expected PowerPoint authored transform missing at ${location}`,
    );
  }
  if (
    generated.left !== expectedPointValue(transform.x) ||
    generated.top !== expectedPointValue(transform.y) ||
    generated.width !== expectedPointValue(transform.width) ||
    generated.height !== expectedPointValue(transform.height)
  ) {
    throw new Error(
      `Generated PowerPoint chart transform mismatch at ${location}`,
    );
  }
  if (generated.chartType !== expected.chartType) {
    throw new Error(`Generated PowerPoint chart type mismatch at ${location}`);
  }
  const expectedData: Chart['data'] = expected.series.map((series) => ({
    key: series.name,
    values: series.values.map((value, index) => ({
      x: String(index),
      y: value,
    })),
    xlabels: Object.fromEntries(
      series.categories.map((category, index) => [String(index), category]),
    ),
  }));
  if (JSON.stringify(generated.data) !== JSON.stringify(expectedData)) {
    throw new Error(`Generated PowerPoint chart data mismatch at ${location}`);
  }
  const expectedColors = expected.series.map((series) => series.color ?? '');
  if (
    (expected.chartType === 'barChart' || expected.chartType === 'lineChart') &&
    JSON.stringify(generated.colors) !== JSON.stringify(expectedColors)
  ) {
    throw new Error(`Generated PowerPoint chart color mismatch at ${location}`);
  }
  if (
    generated.barDir !== expected.barDirection ||
    generated.grouping !== expected.grouping ||
    generated.holeSize !==
      (expected.holeSize === undefined
        ? undefined
        : String(expected.holeSize)) ||
    generated.marker !== expected.marker
  ) {
    throw new Error(
      `Generated PowerPoint chart option mismatch at ${location}`,
    );
  }
}

function verifyImageElement(
  generated: PptxDocument['slides'][number]['elements'][number] | undefined,
  expected: PptxSceneImageElement,
  media: PptxSceneMedia | undefined,
  slideIndex: number,
  elementIndex: number,
): void {
  const location = `slide ${slideIndex + 1}, element ${elementIndex + 1}`;
  if (generated?.type !== 'image') {
    throw new Error(`Generated PowerPoint image missing at ${location}`);
  }
  if (media === undefined) {
    throw new Error(`Expected PowerPoint image media missing at ${location}`);
  }
  verifyElementTransform(generated, expected, location);
  const expectedBase64 = `data:${media.mimeType};base64,${encodeBase64(media.data)}`;
  if (generated.base64 !== expectedBase64) {
    throw new Error(`Generated PowerPoint image data mismatch at ${location}`);
  }
  if (generated.geom !== 'rect') {
    throw new Error(
      `Generated PowerPoint image geometry mismatch at ${location}`,
    );
  }
  const crop = expected.crop;
  if (
    (crop === undefined && generated.rect !== undefined) ||
    (crop !== undefined &&
      (generated.rect === undefined ||
        generated.rect.b !== crop.bottom ||
        generated.rect.l !== crop.left ||
        generated.rect.r !== crop.right ||
        generated.rect.t !== crop.top))
  ) {
    throw new Error(`Generated PowerPoint image crop mismatch at ${location}`);
  }
}

function verifyShapeElement(
  generated: PptxDocument['slides'][number]['elements'][number] | undefined,
  expected: PptxSceneShapeElement,
  slideIndex: number,
  elementIndex: number,
): void {
  const location = `slide ${slideIndex + 1}, element ${elementIndex + 1}`;
  const geometry = expected.authored.geometry ?? 'rect';
  if (generated?.type !== 'shape' || generated.shapType !== geometry) {
    throw new Error(`Generated PowerPoint shape missing at ${location}`);
  }
  verifyElementTransform(generated, expected, location);
  if (
    expected.authored.fillColor !== undefined &&
    (generated.fill?.type !== 'color' ||
      generated.fill.value !== expected.authored.fillColor)
  ) {
    throw new Error(`Generated PowerPoint shape fill mismatch at ${location}`);
  }
  if (
    expected.authored.lineColor !== undefined &&
    generated.borderColor !== expected.authored.lineColor
  ) {
    throw new Error(`Generated PowerPoint shape line mismatch at ${location}`);
  }
  if (
    expected.authored.lineWidth !== undefined &&
    generated.borderWidth !== expected.authored.lineWidth
  ) {
    throw new Error(
      `Generated PowerPoint shape line width mismatch at ${location}`,
    );
  }
}

function verifyRenderedSlides(
  document: PptxDocument,
  rendered: PptxSvgRenderResult,
): void {
  if (rendered.slides.length !== document.slides.length) {
    throw new Error(
      `Generated PowerPoint render count mismatch: expected ${document.slides.length}, received ${rendered.slides.length}`,
    );
  }
  for (const [index, slide] of rendered.slides.entries()) {
    if (
      slide.format !== 'svg' ||
      slide.mimeType !== 'image/svg+xml' ||
      slide.slideNumber !== index + 1 ||
      slide.width !== document.size.width ||
      slide.height !== document.size.height ||
      slide.data.byteLength === 0
    ) {
      throw new Error(
        `Generated PowerPoint visual invariant mismatch on slide ${index + 1}`,
      );
    }
    const source = new TextDecoder().decode(slide.data);
    if (
      !/^(?:<\?xml[^>]*\?>)?<svg\b/.test(source) ||
      /<(?:foreignObject|script)\b/i.test(source) ||
      /(?:href|src)=["'](?:blob|file|https?):/i.test(source)
    ) {
      throw new Error(
        `Generated PowerPoint unsafe SVG output on slide ${index + 1}`,
      );
    }
  }
}

export async function verifyPowerPointCreationWithParser(
  data: Uint8Array,
  scene: PptxSceneDocument,
  parseDocument: PptxCreationParser,
  renderDocument: PptxCreationRenderer = renderPptxDocumentToSvg,
): Promise<void> {
  const chartCount = scene.slides.reduce(
    (count, slide) => count + countSceneCharts(slide.elements),
    0,
  );
  const document = await parseDocument(data, {
    audioMode: 'none',
    errorMode: 'strict',
    imageMode: scene.media.length === 0 ? 'none' : 'base64',
    limits: {
      maxEntries: scene.slides.length * 2 + scene.media.length + chartCount + 9,
      maxSlides: Math.max(1, scene.slides.length),
    },
    videoMode: 'none',
  });
  if (document.slides.length !== scene.slides.length) {
    throw new Error(
      `Generated PowerPoint slide count mismatch: expected ${scene.slides.length}, received ${document.slides.length}`,
    );
  }
  const expectedWidth = expectedPointValue(scene.size.width);
  const expectedHeight = expectedPointValue(scene.size.height);
  if (
    document.size.width !== expectedWidth ||
    document.size.height !== expectedHeight
  ) {
    throw new Error(
      `Generated PowerPoint size mismatch: expected ${expectedWidth}x${expectedHeight}, received ${document.size.width}x${document.size.height}`,
    );
  }
  scene.slides.forEach((slide, index) => {
    const generated = document.slides[index];
    if (!generated || generated.elements.length !== slide.elements.length) {
      throw new Error(
        `Generated PowerPoint element count mismatch on slide ${index + 1}: expected ${slide.elements.length}, received ${generated?.elements.length ?? 0}`,
      );
    }
    slide.elements.forEach((element, elementIndex) =>
      verifySceneElement(
        generated.elements[elementIndex],
        element,
        scene,
        index,
        elementIndex,
      ),
    );
  });
  verifyRenderedSlides(document, renderDocument(document));
}

export function verifyPowerPointCreation(
  data: Uint8Array,
  scene: PptxSceneDocument,
): Promise<void> {
  return verifyPowerPointCreationWithParser(data, scene, parse);
}
