import type { PptxSceneImageCrop } from '../scene-types';
import { unsupportedPptxEdit } from './patch-error';
import {
  escapePptxXmlPattern,
  pptxShapeHasElement,
  qualifiedPptxName,
  resolvePptxEditablePictureXml,
} from './shape-range';
import type { PptxRoundTripSetImageCropOperation } from './types';

interface XmlRange {
  start: number;
  xml: string;
}

function elementRanges(xml: string, qualifiedName: string): XmlRange[] {
  const name = escapePptxXmlPattern(qualifiedName);
  const pattern = new RegExp(
    `<${name}(?:\\s+[A-Za-z_][\\w.:-]*\\s*=\\s*(?:"[^"]*"|'[^']*'))*\\s*>[\\s\\S]*?<\\/${name}\\s*>`,
    'g',
  );
  return [...xml.matchAll(pattern)].map((match) => ({
    start: match.index ?? 0,
    xml: match[0],
  }));
}

function singleElementRange(
  xml: string,
  qualifiedName: string,
  description: string,
): XmlRange {
  const ranges = elementRanges(xml, qualifiedName);
  if (ranges.length !== 1) {
    unsupportedPptxEdit(
      `PowerPoint image crop requires exactly one ${description}`,
    );
  }
  return ranges[0] as XmlRange;
}

function emptyElementMatches(
  xml: string,
  qualifiedName: string,
): RegExpMatchArray[] {
  const name = escapePptxXmlPattern(qualifiedName);
  return [
    ...xml.matchAll(
      new RegExp(
        `<${name}((?:\\s+[A-Za-z_][\\w.:-]*\\s*=\\s*(?:"[^"]*"|'[^']*'))*)\\s*(?:\\/>|>\\s*<\\/${name}\\s*>)`,
        'g',
      ),
    ),
  ];
}

function completeElementMatches(
  xml: string,
  qualifiedName: string,
): RegExpMatchArray[] {
  const name = escapePptxXmlPattern(qualifiedName);
  return [
    ...xml.matchAll(
      new RegExp(
        `<${name}((?:\\s+[A-Za-z_][\\w.:-]*\\s*=\\s*(?:"[^"]*"|'[^']*'))*)\\s*(?:\\/>|>[\\s\\S]*?<\\/${name}\\s*>)`,
        'g',
      ),
    ),
  ];
}

function cropAttributeValue(value: string | undefined): number {
  if (value === undefined) return 0;
  if (!/^[+-]?\d+$/.test(value)) {
    unsupportedPptxEdit('PowerPoint image crop source percentage is invalid');
  }
  const integer = Number(value);
  if (integer < -100_000 || integer > 100_000) {
    unsupportedPptxEdit('PowerPoint image crop source percentage is unsafe');
  }
  return integer / 1_000;
}

function parseCropAttributes(value: string): PptxSceneImageCrop {
  const attributes = new Map<string, string>();
  const pattern = /\s([A-Za-z_][\w.:-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  for (const match of value.matchAll(pattern)) {
    const name = match[1] as string;
    if (!['b', 'l', 'r', 't'].includes(name) || attributes.has(name)) {
      unsupportedPptxEdit('PowerPoint image crop source attributes are unsafe');
    }
    attributes.set(name, (match[2] ?? match[3]) as string);
  }
  const crop = {
    bottom: cropAttributeValue(attributes.get('b')),
    left: cropAttributeValue(attributes.get('l')),
    right: cropAttributeValue(attributes.get('r')),
    top: cropAttributeValue(attributes.get('t')),
  };
  if (crop.left + crop.right >= 100 || crop.top + crop.bottom >= 100) {
    unsupportedPptxEdit(
      'PowerPoint image crop source has no positive visible region',
    );
  }
  return crop;
}

function sameCrop(
  left: PptxSceneImageCrop | null,
  right: PptxSceneImageCrop | null,
): boolean {
  return (
    left === right ||
    (left !== null &&
      right !== null &&
      left.bottom === right.bottom &&
      left.left === right.left &&
      left.right === right.right &&
      left.top === right.top)
  );
}

function serializeCrop(crop: PptxSceneImageCrop, prefix: string): string {
  const name = qualifiedPptxName(prefix, 'srcRect');
  return `<${name} l="${Math.round(crop.left * 1_000)}" t="${Math.round(crop.top * 1_000)}" r="${Math.round(crop.right * 1_000)}" b="${Math.round(crop.bottom * 1_000)}"/>`;
}

export function patchPptxImageCropXml(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripSetImageCropOperation,
): string {
  const picture = resolvePptxEditablePictureXml(xml, shapeId);
  if (
    (picture.markupPrefix !== undefined &&
      pptxShapeHasElement(
        picture.shape,
        qualifiedPptxName(picture.markupPrefix, 'AlternateContent'),
      )) ||
    pptxShapeHasElement(
      picture.shape,
      qualifiedPptxName(picture.presentationPrefix, 'extLst'),
    ) ||
    pptxShapeHasElement(
      picture.shape,
      qualifiedPptxName(picture.drawingPrefix, 'extLst'),
    )
  ) {
    unsupportedPptxEdit(
      'PowerPoint image crop target contains unsupported compatibility markup',
    );
  }
  const fill = singleElementRange(
    picture.shape,
    qualifiedPptxName(picture.presentationPrefix, 'blipFill'),
    'picture fill',
  );
  const sourceRectName = qualifiedPptxName(picture.drawingPrefix, 'srcRect');
  const sourceRects = emptyElementMatches(fill.xml, sourceRectName);
  if (
    sourceRects.length === 0 &&
    pptxShapeHasElement(fill.xml, sourceRectName)
  ) {
    unsupportedPptxEdit('PowerPoint image crop source attributes are invalid');
  }
  if (sourceRects.length > 1) {
    unsupportedPptxEdit(
      'PowerPoint image crop requires at most one source rectangle',
    );
  }
  const sourceRect = sourceRects[0];
  const sourceCrop =
    sourceRect === undefined
      ? null
      : parseCropAttributes(sourceRect[1] as string);
  if (!sameCrop(sourceCrop, operation.expectedCrop)) {
    unsupportedPptxEdit(
      'PowerPoint image crop source XML does not match its preview precondition',
    );
  }
  if (sourceRect !== undefined) {
    const start = picture.range.start + fill.start + (sourceRect.index ?? 0);
    const replacement =
      operation.value === null
        ? ''
        : serializeCrop(operation.value, picture.drawingPrefix);
    return `${xml.slice(0, start)}${replacement}${xml.slice(start + sourceRect[0].length)}`;
  }
  if (operation.value === null) {
    unsupportedPptxEdit('PowerPoint image crop removal has no source crop');
  }
  const blips = completeElementMatches(
    fill.xml,
    qualifiedPptxName(picture.drawingPrefix, 'blip'),
  );
  if (blips.length !== 1) {
    unsupportedPptxEdit('PowerPoint image crop requires one picture blip');
  }
  const blip = blips[0] as RegExpMatchArray;
  const insert =
    picture.range.start + fill.start + (blip.index ?? 0) + blip[0].length;
  return `${xml.slice(0, insert)}${serializeCrop(operation.value, picture.drawingPrefix)}${xml.slice(insert)}`;
}
