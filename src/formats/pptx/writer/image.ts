import type { PptxSceneImageElement, PptxSceneTransform } from '../scene-types';
import { serializeShapeTransform } from './shape';
import { escapeXmlAttribute } from './xml';

function booleanAttribute(value: boolean): string {
  return value ? '1' : '0';
}

function serializePictureCrop(element: PptxSceneImageElement): string {
  const crop = element.crop;
  if (crop === undefined) return '';
  return `<a:srcRect l="${Math.round(crop.left * 1_000)}" t="${Math.round(crop.top * 1_000)}" r="${Math.round(crop.right * 1_000)}" b="${Math.round(crop.bottom * 1_000)}"/>`;
}

function serializePictureNonVisualProperties(
  element: PptxSceneImageElement,
  shapeId: number,
): string {
  const attributes = [
    `id="${shapeId}"`,
    `name="${escapeXmlAttribute(element.name ?? `Picture ${shapeId}`)}"`,
  ];
  if (element.description !== undefined) {
    attributes.push(`descr="${escapeXmlAttribute(element.description)}"`);
  }
  if (element.title !== undefined) {
    attributes.push(`title="${escapeXmlAttribute(element.title)}"`);
  }
  if (element.authored.hidden !== undefined) {
    attributes.push(`hidden="${booleanAttribute(element.authored.hidden)}"`);
  }
  return `<p:nvPicPr><p:cNvPr ${attributes.join(' ')}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>`;
}

export function serializePicture(
  element: PptxSceneImageElement,
  transform: PptxSceneTransform,
  shapeId: number,
  relationshipId: string,
): string {
  const nonVisual = serializePictureNonVisualProperties(element, shapeId);
  const blipFill = `<p:blipFill><a:blip r:embed="${escapeXmlAttribute(relationshipId)}"/>${serializePictureCrop(element)}<a:stretch><a:fillRect/></a:stretch></p:blipFill>`;
  const shapeProperties = `<p:spPr>${serializeShapeTransform(transform)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`;
  return `<p:pic>${nonVisual}${blipFill}${shapeProperties}</p:pic>`;
}
