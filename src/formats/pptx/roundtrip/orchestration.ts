import {
  assertPptxArchiveWithinLimits,
  assertPptxInputWithinLimits,
  type ResolvedPptxResourceLimits,
} from '../internal/resource-limits';
import type {
  PptxSceneGroupTransform,
  PptxSceneTransform,
} from '../scene-types';
import type { Element, Group, PptxDocument } from '../types';
import {
  assertSafeEditablePptxPackage,
  decodeEditablePptxXml,
  generatePptxPatchedArchive,
  readPptxPartPayloads,
  verifyPptxPatchedPayloads,
} from './package-preservation';
import { patchPptxImageCropXml } from './image-crop-xml';
import { unsupportedPptxEdit } from './patch-error';
import { resolvePptxSlideParts } from './relationships';
import { patchPptxShapeTextXml, patchPptxTableCellTextXml } from './text-xml';
import {
  patchPptxGraphicFrameTransformXml,
  patchPptxChartFrameTransformXml,
  patchPptxGroupTransformXml,
  patchPptxPictureTransformXml,
  patchPptxShapeTransformXml,
} from './transform-xml';
import type {
  PptxRoundTripOperation,
  PptxRoundTripReplaceTextOperation,
  PptxRoundTripSetImageCropOperation,
  PptxRoundTripSetTransformOperation,
} from './types';
import JSZip from 'jszip';

const TEXT_TARGET_KEY_PATTERN =
  /^slide-([1-9]\d*)((?:-element-[1-9]\d*)+)-run-1$/;
const TABLE_TEXT_TARGET_KEY_PATTERN =
  /^slide-([1-9]\d*)((?:-element-[1-9]\d*)+)-row-([1-9]\d*)-cell-([1-9]\d*)-run-1$/;
const TRANSFORM_TARGET_KEY_PATTERN =
  /^slide-([1-9]\d*)((?:-element-[1-9]\d*)+)$/;
const ELEMENT_INDEX_PATTERN = /-element-([1-9]\d*)/g;

interface TextTarget {
  columnIndex?: number;
  elementType: 'chart' | 'group' | 'image' | 'shape' | 'table' | 'text';
  rowIndex?: number;
  shapeId: string;
  slideIndex: number;
  transformOperation?: PptxRoundTripSetTransformOperation;
}

function parsedGroupTransform(group: Group): PptxSceneGroupTransform {
  if (group.childSpace === undefined) {
    unsupportedPptxEdit(
      'PowerPoint nested transform ancestor has no child coordinate space',
    );
  }
  return {
    childSpace: { ...group.childSpace },
    flipHorizontal: group.isFlipH,
    flipVertical: group.isFlipV,
    height: group.height,
    rotation: group.rotate,
    width: group.width,
    x: group.left,
    y: group.top,
  };
}

function localizeNestedTransform(
  transform: PptxSceneTransform,
  ancestors: readonly PptxSceneGroupTransform[],
): PptxSceneTransform {
  let result = { ...transform };
  ancestors.forEach((ancestor, index) => {
    const widthScale = ancestor.width / ancestor.childSpace.width;
    const heightScale = ancestor.height / ancestor.childSpace.height;
    const directChild = index === ancestors.length - 1;
    const offsetX = directChild ? ancestor.childSpace.x : 0;
    const offsetY = directChild ? ancestor.childSpace.y : 0;
    const rotation = (((result.rotation ?? 0) % 360) + 360) % 360;
    const swapped = rotation === 90 || rotation === 270;
    const width = result.width / (swapped ? heightScale : widthScale);
    const height = result.height / (swapped ? widthScale : heightScale);
    result = {
      ...result,
      height,
      width,
      x: (result.x + result.width / 2) / widthScale + offsetX - width / 2,
      y: (result.y + result.height / 2) / heightScale + offsetY - height / 2,
    };
  });
  return result;
}

function localizedOperation(
  operation: PptxRoundTripSetTransformOperation,
  ancestors: readonly PptxSceneGroupTransform[],
): PptxRoundTripSetTransformOperation {
  return {
    ...operation,
    expectedTransform: {
      ...operation.expectedTransform,
      ...localizeNestedTransform(operation.expectedTransform, ancestors),
    },
    value: {
      ...operation.value,
      ...localizeNestedTransform(operation.value, ancestors),
    },
  };
}

export interface PptxPatchedPackage {
  copiedPartCount: number;
  data: Uint8Array;
  patchedPartCount: number;
}

function textTarget(
  operation: PptxRoundTripReplaceTextOperation,
  document: PptxDocument,
): TextTarget {
  const tableMatch = TABLE_TEXT_TARGET_KEY_PATTERN.exec(operation.targetKey);
  const match = tableMatch ?? TEXT_TARGET_KEY_PATTERN.exec(operation.targetKey);
  if (match === null) {
    unsupportedPptxEdit(
      'PowerPoint text edit target is not a supported native text run key',
    );
  }
  const slideIndex = Number(match[1]) - 1;
  const elementIndexes = [
    ...(match[2] as string).matchAll(ELEMENT_INDEX_PATTERN),
  ].map((indexMatch) => Number(indexMatch[1]) - 1);
  const rowIndex = tableMatch === null ? undefined : Number(match[3]) - 1;
  const columnIndex = tableMatch === null ? undefined : Number(match[4]) - 1;
  if (
    !Number.isSafeInteger(slideIndex) ||
    elementIndexes.some((index) => !Number.isSafeInteger(index)) ||
    (rowIndex !== undefined && !Number.isSafeInteger(rowIndex)) ||
    (columnIndex !== undefined && !Number.isSafeInteger(columnIndex))
  ) {
    unsupportedPptxEdit('PowerPoint text edit target index is unsafe');
  }
  let elements = document.slides[slideIndex]?.elements;
  let element: Element | undefined;
  for (const [depth, elementIndex] of elementIndexes.entries()) {
    element = elements?.[elementIndex];
    if (depth === elementIndexes.length - 1) break;
    if (element?.type !== 'group') {
      unsupportedPptxEdit(
        'PowerPoint text edit target path crosses a non-group element',
      );
    }
    elements = element.elements;
  }
  if (tableMatch !== null) {
    if (element?.type !== 'table') {
      unsupportedPptxEdit(
        'PowerPoint table text edit target is not a native table element',
      );
    }
    return {
      columnIndex: columnIndex as number,
      elementType: 'table',
      rowIndex: rowIndex as number,
      shapeId: element.id,
      slideIndex,
    };
  }
  if (element?.type !== 'text') {
    unsupportedPptxEdit(
      'PowerPoint text edit target is not a native text element',
    );
  }
  return {
    elementType: 'text',
    shapeId: element.id,
    slideIndex,
  };
}

function transformTarget(
  operation:
    PptxRoundTripSetImageCropOperation | PptxRoundTripSetTransformOperation,
  document: PptxDocument,
): TextTarget {
  const match = TRANSFORM_TARGET_KEY_PATTERN.exec(operation.targetKey);
  if (match === null) {
    unsupportedPptxEdit(
      'PowerPoint transform target is not a supported slide text element key',
    );
  }
  const slideIndex = Number(match[1]) - 1;
  const elementIndexes = [
    ...(match[2] as string).matchAll(ELEMENT_INDEX_PATTERN),
  ].map((indexMatch) => Number(indexMatch[1]) - 1);
  if (
    !Number.isSafeInteger(slideIndex) ||
    elementIndexes.some((index) => !Number.isSafeInteger(index))
  ) {
    unsupportedPptxEdit('PowerPoint transform target index is unsafe');
  }
  let elements = document.slides[slideIndex]?.elements;
  let element: Element | undefined;
  const ancestors: PptxSceneGroupTransform[] = [];
  for (const [depth, elementIndex] of elementIndexes.entries()) {
    element = elements?.[elementIndex];
    if (depth === elementIndexes.length - 1) break;
    if (element?.type !== 'group') {
      unsupportedPptxEdit(
        'PowerPoint transform target path crosses a non-group element',
      );
    }
    const resolvedAncestor = parsedGroupTransform(element);
    const localAncestor = localizeNestedTransform(
      resolvedAncestor,
      ancestors,
    ) as PptxSceneGroupTransform;
    ancestors.push(localAncestor);
    elements = element.elements;
  }
  if (
    element?.type !== 'chart' &&
    element?.type !== 'image' &&
    element?.type !== 'group' &&
    element?.type !== 'shape' &&
    element?.type !== 'table' &&
    element?.type !== 'text'
  ) {
    unsupportedPptxEdit(
      'PowerPoint transform target is not a slide-owned text, shape, image, table, chart, or group element',
    );
  }
  return {
    elementType: element.type,
    shapeId: element.id,
    slideIndex,
    ...(operation.kind === 'set-transform'
      ? { transformOperation: localizedOperation(operation, ancestors) }
      : {}),
  };
}

export async function patchPptxOperations(
  bytes: Uint8Array,
  document: PptxDocument,
  operations: readonly PptxRoundTripOperation[],
  limits: ResolvedPptxResourceLimits,
): Promise<PptxPatchedPackage> {
  const archive = await JSZip.loadAsync(bytes);
  assertPptxArchiveWithinLimits(archive, limits);
  assertSafeEditablePptxPackage(archive);
  const [slides, sourcePayloads] = await Promise.all([
    resolvePptxSlideParts(archive, limits),
    readPptxPartPayloads(archive, limits),
  ]);
  if (slides.length !== document.slides.length) {
    unsupportedPptxEdit(
      'PowerPoint text edit slide order does not match the parsed document',
    );
  }

  const patchedParts = new Set<string>();
  const editedXml = new Map<string, string>();
  for (const operation of operations) {
    const target =
      operation.kind === 'replace-text'
        ? textTarget(operation, document)
        : transformTarget(operation, document);
    if (operation.kind === 'set-image-crop' && target.elementType !== 'image') {
      unsupportedPptxEdit(
        'PowerPoint image crop target is not a native image element',
      );
    }
    const slidePart = slides[target.slideIndex] as string;
    const sourceBytes = sourcePayloads.get(slidePart) as Uint8Array;
    const current =
      editedXml.get(slidePart) ?? decodeEditablePptxXml(sourceBytes, limits);
    const patched =
      operation.kind === 'set-image-crop'
        ? patchPptxImageCropXml(current, target.shapeId, operation)
        : operation.kind === 'replace-text'
          ? target.elementType === 'table'
            ? patchPptxTableCellTextXml(
                current,
                target.shapeId,
                target.rowIndex as number,
                target.columnIndex as number,
                operation,
              )
            : patchPptxShapeTextXml(current, target.shapeId, operation)
          : target.elementType === 'chart'
            ? patchPptxChartFrameTransformXml(
                current,
                target.shapeId,
                target.transformOperation as PptxRoundTripSetTransformOperation,
              )
            : target.elementType === 'image'
              ? patchPptxPictureTransformXml(
                  current,
                  target.shapeId,
                  target.transformOperation as PptxRoundTripSetTransformOperation,
                )
              : target.elementType === 'group'
                ? patchPptxGroupTransformXml(
                    current,
                    target.shapeId,
                    target.transformOperation as PptxRoundTripSetTransformOperation,
                  )
                : target.elementType === 'table'
                  ? patchPptxGraphicFrameTransformXml(
                      current,
                      target.shapeId,
                      target.transformOperation as PptxRoundTripSetTransformOperation,
                    )
                  : patchPptxShapeTransformXml(
                      current,
                      target.shapeId,
                      target.transformOperation as PptxRoundTripSetTransformOperation,
                    );
    editedXml.set(slidePart, patched);
    patchedParts.add(slidePart);
  }

  for (const [part, xml] of editedXml) {
    const entry = archive.file(part) as JSZip.JSZipObject;
    archive.file(part, xml, {
      date: entry.date,
    });
  }
  const output = await generatePptxPatchedArchive(archive);
  assertPptxInputWithinLimits(output, limits);

  const outputArchive = await JSZip.loadAsync(output);
  assertPptxArchiveWithinLimits(outputArchive, limits);
  const outputPayloads = await readPptxPartPayloads(outputArchive, limits);
  verifyPptxPatchedPayloads(sourcePayloads, outputPayloads, patchedParts);

  return {
    copiedPartCount: sourcePayloads.size - patchedParts.size,
    data: output,
    patchedPartCount: patchedParts.size,
  };
}
