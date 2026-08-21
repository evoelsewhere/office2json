import { decodeXmlEntities } from '../../../common/text/html';
import type {
  PptxSceneDocument,
  PptxSceneChartElement,
  PptxSceneElement,
  PptxSceneImageElement,
  PptxSceneShapeElement,
  PptxSceneSlide,
  PptxSceneTextBodyProperties,
  PptxSceneTextElement,
  PptxSceneTransform,
  PptxSceneUnsupportedElement,
} from '../scene-types';
import type {
  Chart,
  Image,
  PptxDocument,
  PptxElement,
  Shape,
  Text,
} from '../types';
import { createPptxRoundTripTablePreview } from './table-preview';
import { createPptxRoundTripGroupPreview } from './group-preview';

function resolvedTransform(
  element: PptxElement,
): PptxSceneTransform | undefined {
  if (
    !Number.isFinite(element.left) ||
    !Number.isFinite(element.top) ||
    !Number.isFinite(element.width) ||
    element.width <= 0 ||
    !Number.isFinite(element.height) ||
    element.height <= 0
  ) {
    return undefined;
  }
  return {
    height: element.height,
    width: element.width,
    x: element.left,
    y: element.top,
    ...('isFlipH' in element ? { flipHorizontal: element.isFlipH } : {}),
    ...('isFlipV' in element ? { flipVertical: element.isFlipV } : {}),
    ...('rotate' in element ? { rotation: element.rotate } : {}),
  };
}

function previewText(element: PptxElement): string | undefined {
  return 'content' in element ? element.content : undefined;
}

export function plainTextFromPowerPointHtml(html: string): string {
  const withLineBreaks = html
    .replace(/<br\s*\/?\s*>/gi, '\n')
    .replace(/<\/(?:li|p)\s*>/gi, '\n');
  const withoutTags = withLineBreaks.replace(/<[^>]*>/g, '');
  return decodeXmlEntities(withoutTags.replace(/&nbsp;/gi, ' ')).replace(
    /\n+$/,
    '',
  );
}

function textBodyProperties(element: Text): PptxSceneTextBodyProperties {
  const anchor =
    element.vAlign === 'down'
      ? 'bottom'
      : element.vAlign === 'mid'
        ? 'center'
        : element.vAlign === 'dist'
          ? 'distributed'
          : element.vAlign === 'just'
            ? 'justified'
            : 'top';
  return {
    anchor,
    ...(element.autoFit === undefined ? {} : { autoFit: element.autoFit.type }),
    vertical: element.isVertical,
    wrap: element.wrap,
  };
}

function sceneTextElement(
  element: Text,
  slideIndex: number,
  elementIndex: number,
  keyOverride?: string,
): PptxSceneTextElement {
  const key =
    keyOverride ?? `slide-${slideIndex + 1}-element-${elementIndex + 1}`;
  const transform = resolvedTransform(element) as PptxSceneTransform;
  return {
    authored: {},
    key,
    name: element.name,
    resolved: {
      hidden: false,
      transform,
    },
    text: {
      body: textBodyProperties(element),
      paragraphs: [
        {
          children: [
            {
              key: `${key}-run-1`,
              text: plainTextFromPowerPointHtml(element.content),
              type: 'run',
            },
          ],
          key: `${key}-paragraph-1`,
        },
      ],
    },
    type: 'text',
  };
}

function sceneShapeElement(
  element: Shape,
  slideIndex: number,
  elementIndex: number,
  keyOverride?: string,
): PptxSceneShapeElement {
  const transform = resolvedTransform(element);
  return {
    authored: {},
    key: keyOverride ?? `slide-${slideIndex + 1}-element-${elementIndex + 1}`,
    name: element.name,
    resolved: {
      hidden: false,
      ...(transform === undefined ? {} : { transform }),
    },
    type: 'shape',
  };
}

function sceneImageElement(
  element: Image,
  slideIndex: number,
  elementIndex: number,
  keyOverride?: string,
): PptxSceneImageElement {
  const transform = resolvedTransform(element);
  return {
    authored: {},
    ...(element.rect === undefined
      ? {}
      : {
          crop: {
            bottom: element.rect.b ?? 0,
            left: element.rect.l ?? 0,
            right: element.rect.r ?? 0,
            top: element.rect.t ?? 0,
          },
        }),
    key: keyOverride ?? `slide-${slideIndex + 1}-element-${elementIndex + 1}`,
    resolved: {
      hidden: false,
      ...(transform === undefined ? {} : { transform }),
    },
    type: 'image',
  };
}

function sceneChartElement(
  element: Chart,
  slideIndex: number,
  elementIndex: number,
  keyOverride?: string,
): PptxSceneChartElement | undefined {
  if (
    element.chartType !== 'barChart' &&
    element.chartType !== 'doughnutChart' &&
    element.chartType !== 'lineChart' &&
    element.chartType !== 'pieChart'
  ) {
    return undefined;
  }
  const transform = resolvedTransform(element);
  if (transform === undefined || !Array.isArray(element.data)) return undefined;
  const key =
    keyOverride ?? `slide-${slideIndex + 1}-element-${elementIndex + 1}`;
  const series: PptxSceneChartElement['series'] = [];
  for (const [seriesIndex, item] of element.data.entries()) {
    if (
      Array.isArray(item) ||
      typeof item.key !== 'string' ||
      item.key.trim() === '' ||
      !Array.isArray(item.values) ||
      item.values.length === 0
    ) {
      return undefined;
    }
    const categories: string[] = [];
    const values: number[] = [];
    for (const point of item.values) {
      if (typeof point.x !== 'string' || !Number.isFinite(point.y)) {
        return undefined;
      }
      categories.push(item.xlabels[point.x] ?? point.x);
      values.push(point.y);
    }
    const color = element.colors[seriesIndex];
    series.push({
      categories,
      ...(typeof color === 'string' && /^#[0-9a-f]{6}$/i.test(color)
        ? { color }
        : {}),
      key: `${key}-series-${seriesIndex + 1}`,
      name: item.key,
      values,
    });
  }
  if (
    (element.chartType === 'pieChart' ||
      element.chartType === 'doughnutChart') &&
    series.length !== 1
  ) {
    return undefined;
  }
  if (
    element.grouping !== undefined &&
    element.grouping !== 'clustered' &&
    element.grouping !== 'percentStacked' &&
    element.grouping !== 'stacked' &&
    element.grouping !== 'standard'
  ) {
    return undefined;
  }
  const holeSize =
    element.holeSize === undefined ? undefined : Number(element.holeSize);
  if (
    holeSize !== undefined &&
    (!Number.isSafeInteger(holeSize) || holeSize < 10 || holeSize > 90)
  ) {
    return undefined;
  }
  return {
    authored: {},
    ...(element.barDir === undefined ? {} : { barDirection: element.barDir }),
    chartType: element.chartType,
    ...(element.grouping === undefined ? {} : { grouping: element.grouping }),
    ...(holeSize === undefined ? {} : { holeSize }),
    key,
    ...(element.marker === undefined ? {} : { marker: element.marker }),
    resolved: {
      hidden: false,
      transform: {
        ...transform,
        flipHorizontal: false,
        flipVertical: false,
        rotation: 0,
      },
    },
    series,
    type: 'chart',
  };
}

function sceneUnsupportedElement(
  element: PptxElement,
  slideIndex: number,
  elementIndex: number,
  keyOverride?: string,
): PptxSceneUnsupportedElement {
  const text = previewText(element);
  const transform = resolvedTransform(element);
  return {
    authored: {},
    feature: element.type,
    key: keyOverride ?? `slide-${slideIndex + 1}-element-${elementIndex + 1}`,
    ...(text === undefined ? {} : { previewText: text }),
    resolved: {
      hidden: false,
      ...(transform === undefined ? {} : { transform }),
    },
    type: 'unsupported',
  };
}

function sceneElement(
  element: PptxElement,
  slideIndex: number,
  elementIndex: number,
  keyOverride?: string,
): PptxSceneElement {
  if (element.type === 'text') {
    return sceneTextElement(element, slideIndex, elementIndex, keyOverride);
  }
  if (element.type === 'shape') {
    return plainTextFromPowerPointHtml(element.content) === ''
      ? sceneShapeElement(element, slideIndex, elementIndex, keyOverride)
      : sceneUnsupportedElement(element, slideIndex, elementIndex, keyOverride);
  }
  if (element.type === 'image') {
    return sceneImageElement(element, slideIndex, elementIndex, keyOverride);
  }
  if (element.type === 'chart') {
    return (
      sceneChartElement(element, slideIndex, elementIndex, keyOverride) ??
      sceneUnsupportedElement(element, slideIndex, elementIndex, keyOverride)
    );
  }
  if (element.type === 'table') {
    return (
      createPptxRoundTripTablePreview(
        element,
        slideIndex,
        elementIndex,
        plainTextFromPowerPointHtml,
        resolvedTransform,
        keyOverride,
      ) ??
      sceneUnsupportedElement(element, slideIndex, elementIndex, keyOverride)
    );
  }
  if (element.type === 'group') {
    return (
      createPptxRoundTripGroupPreview(
        element,
        slideIndex,
        elementIndex,
        {
          mapChild: (child, childIndex, key) =>
            sceneElement(child, slideIndex, childIndex, key),
          resolveTransform: resolvedTransform,
        },
        keyOverride,
      ) ??
      sceneUnsupportedElement(element, slideIndex, elementIndex, keyOverride)
    );
  }
  return sceneUnsupportedElement(
    element,
    slideIndex,
    elementIndex,
    keyOverride,
  );
}

function sceneSlide(slide: PptxDocument['slides'][number], index: number) {
  const result: PptxSceneSlide = {
    elements: slide.elements.map((element, elementIndex) =>
      sceneElement(element, index, elementIndex),
    ),
    key: `slide-${index + 1}`,
  };
  return result;
}

export function createPowerPointRoundTripPreview(
  document: PptxDocument,
): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { ...document.size },
    slides: document.slides.map(sceneSlide),
    themes: [],
  };
}
