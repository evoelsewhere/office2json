import { resolvePptxResourceLimits } from '../internal/resource-limits';
import { RATIO_EMUs_Points } from '../../../common/ooxml/units';
import {
  isRepresentablePptxCropPercentage,
  isValidXmlText,
} from '../scene-validation';
import { PptxWriteError } from '../write-error';
import { degreesToAngle, pointsToEmu } from '../writer/units';
import type {
  PptxSceneChartElement,
  PptxSceneImageElement,
  PptxSceneGroupElement,
  PptxSceneGroupTransform,
  PptxSceneElement,
  PptxSceneImageCrop,
  PptxSceneShapeElement,
  PptxSceneTableElement,
  PptxSceneTextElement,
  PptxSceneTransform,
} from '../scene-types';
import { canonicalJson } from './canonical-json';
import {
  createPptxRoundTripNativeEditSupportProfile,
  createPptxRoundTripTextEditSupportProfile,
  createPptxSnapshotConsistency,
} from './consistency';
import type {
  PptxRoundTripReplaceTextOperation,
  PptxRoundTripSetImageCropOperation,
  PptxRoundTripSetTransformOperation,
  PptxRoundTripSnapshot,
} from './types';
import { validatePptxRoundTripSnapshot } from './validate';
import { scalePptxTableIntegerSizes } from './transform-xml';

export interface PptxRoundTripReplaceTextRequest {
  targetKey: string;
  value: string;
}

export interface PptxRoundTripSetTransformRequest {
  targetKey: string;
  value: PptxSceneTransform;
}

export interface PptxRoundTripSetGroupTransformRequest {
  targetKey: string;
  value: PptxSceneGroupTransform;
}

export interface PptxRoundTripSetImageCropRequest {
  targetKey: string;
  value: PptxSceneImageCrop | null;
}

function scaledTableSizes(
  values: readonly number[],
  replacementTotal: number,
): number[] {
  const source = values.map(pointsToEmu);
  const replacement = pointsToEmu(replacementTotal);
  return scalePptxTableIntegerSizes(source, replacement).map(
    (value) => value * RATIO_EMUs_Points,
  );
}

function scaleGroupElementTransform(
  element: PptxSceneElement,
  expected: PptxSceneGroupTransform,
  replacement: PptxSceneGroupTransform,
  directChild: boolean,
): void {
  const transform = element.resolved.transform;
  if (transform !== undefined) {
    const oldWidthScale = expected.width / expected.childSpace.width;
    const oldHeightScale = expected.height / expected.childSpace.height;
    const newWidthScale = replacement.width / replacement.childSpace.width;
    const newHeightScale = replacement.height / replacement.childSpace.height;
    const oldOffsetX = directChild ? expected.childSpace.x : 0;
    const oldOffsetY = directChild ? expected.childSpace.y : 0;
    const newOffsetX = directChild ? replacement.childSpace.x : 0;
    const newOffsetY = directChild ? replacement.childSpace.y : 0;
    const centerX = transform.x + transform.width / 2;
    const centerY = transform.y + transform.height / 2;
    const rawCenterX = centerX / oldWidthScale + oldOffsetX;
    const rawCenterY = centerY / oldHeightScale + oldOffsetY;
    const rotation = (((transform.rotation ?? 0) % 360) + 360) % 360;
    const swapped = rotation === 90 || rotation === 270;
    const oldElementWidthScale = swapped ? oldHeightScale : oldWidthScale;
    const oldElementHeightScale = swapped ? oldWidthScale : oldHeightScale;
    const newElementWidthScale = swapped ? newHeightScale : newWidthScale;
    const newElementHeightScale = swapped ? newWidthScale : newHeightScale;
    const width =
      (transform.width / oldElementWidthScale) * newElementWidthScale;
    const height =
      (transform.height / oldElementHeightScale) * newElementHeightScale;
    element.resolved.transform = {
      ...transform,
      height,
      width,
      x: (rawCenterX - newOffsetX) * newWidthScale - width / 2,
      y: (rawCenterY - newOffsetY) * newHeightScale - height / 2,
    };
  }
  if (element.type === 'group') {
    element.elements.forEach((child) =>
      scaleGroupElementTransform(child, expected, replacement, false),
    );
  }
}

export function applyPptxRoundTripOperationsToPreview(
  snapshot: PptxRoundTripSnapshot,
): PptxRoundTripSnapshot['document'] {
  const document = structuredClone(snapshot.document);
  for (const operation of snapshot.operations) {
    if (operation.kind === 'set-transform') {
      let applied = false;
      for (const slide of document.slides) {
        visitTransformElements(slide.elements, (element) => {
          if (element.key === operation.targetKey) {
            if (element.type === 'table') {
              element.columns = scaledTableSizes(
                element.columns,
                operation.value.width,
              );
              const rowHeights = scaledTableSizes(
                element.rows.map((row) => row.height),
                operation.value.height,
              );
              element.rows.forEach((row, index) => {
                row.height = rowHeights[index] as number;
              });
            }
            if (
              element.type === 'group' &&
              'childSpace' in operation.expectedTransform &&
              'childSpace' in operation.value
            ) {
              element.elements.forEach((child) =>
                scaleGroupElementTransform(
                  child,
                  operation.expectedTransform as PptxSceneGroupTransform,
                  operation.value as PptxSceneGroupTransform,
                  true,
                ),
              );
            }
            element.resolved.transform = structuredClone(operation.value);
            applied = true;
          }
        });
      }
      if (!applied) {
        throw new PptxWriteError(
          'verification-failed',
          `PowerPoint transform verification target disappeared: ${operation.targetKey}`,
        );
      }
      continue;
    }
    if (operation.kind === 'set-image-crop') {
      let applied = false;
      for (const slide of document.slides) {
        visitTransformElements(slide.elements, (element) => {
          if (element.type === 'image' && element.key === operation.targetKey) {
            if (operation.value === null) {
              delete element.crop;
            } else {
              element.crop = structuredClone(operation.value);
            }
            applied = true;
          }
        });
      }
      if (!applied) {
        throw new PptxWriteError(
          'verification-failed',
          `PowerPoint image crop verification target disappeared: ${operation.targetKey}`,
        );
      }
      continue;
    }
    let applied = false;
    for (const slide of document.slides) {
      visitEditableTextRuns(slide.elements, (run) => {
        if (run.key === operation.targetKey) {
          run.text = operation.value;
          applied = true;
        }
      });
    }
    if (!applied) {
      throw new PptxWriteError(
        'verification-failed',
        `PowerPoint text edit verification target disappeared: ${operation.targetKey}`,
      );
    }
  }
  return document;
}

type PptxTransformElement =
  | PptxSceneChartElement
  | PptxSceneGroupElement
  | PptxSceneImageElement
  | PptxSceneShapeElement
  | PptxSceneTableElement
  | PptxSceneTextElement;

type PptxEditableTextRun = Extract<
  PptxSceneTextElement['text']['paragraphs'][number]['children'][number],
  { type: 'run' }
>;

function visitTextBodyRuns(
  text: PptxSceneTextElement['text'],
  visitor: (run: PptxEditableTextRun) => void,
): void {
  for (const paragraph of text.paragraphs) {
    for (const child of paragraph.children) {
      if (child.type === 'run') visitor(child);
    }
  }
}

function visitEditableTextRuns(
  elements: readonly PptxSceneElement[],
  visitor: (run: PptxEditableTextRun, nativeOwner: boolean) => void,
  nestedInGroup = false,
): void {
  for (const element of elements) {
    if (element.type === 'text') {
      visitTextBodyRuns(element.text, (run) => visitor(run, nestedInGroup));
    } else if (element.type === 'table') {
      for (const row of element.rows) {
        for (const cell of row.cells) {
          visitTextBodyRuns(cell.text, (run) => visitor(run, true));
        }
      }
    } else if (element.type === 'group') {
      visitEditableTextRuns(element.elements, visitor, true);
    }
  }
}

function visitTransformElements(
  elements: readonly PptxSceneElement[],
  visitor: (element: PptxTransformElement) => void,
): void {
  for (const element of elements) {
    if (
      element.type === 'chart' ||
      element.type === 'image' ||
      element.type === 'group' ||
      element.type === 'shape' ||
      element.type === 'table' ||
      element.type === 'text'
    ) {
      visitor(element);
    }
    if (element.type === 'group') {
      visitTransformElements(element.elements, visitor);
    }
  }
}

function findTransformElement(
  snapshot: PptxRoundTripSnapshot,
  targetKey: string,
  targetType: PptxTransformElement['type'],
): PptxTransformElement {
  let matched: PptxTransformElement | undefined;
  for (const slide of snapshot.document.slides) {
    visitTransformElements(slide.elements, (element) => {
      if (element.type === targetType && element.key === targetKey) {
        matched = element;
      }
    });
  }
  if (matched === undefined) {
    invalidEdit(
      targetType === 'text'
        ? 'PowerPoint transform target key does not exist'
        : `PowerPoint ${targetType} transform target key does not exist`,
    );
  }
  return matched;
}

export function normalizePptxRoundTripTransform(
  value: PptxSceneTransform,
): PptxSceneTransform {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    invalidEdit('PowerPoint transform value must be an object');
  }
  const allowedKeys = new Set([
    'flipHorizontal',
    'flipVertical',
    'height',
    'rotation',
    'width',
    'x',
    'y',
  ]);
  if (Object.keys(value).some((key) => !allowedKeys.has(key))) {
    invalidEdit('PowerPoint transform value is not a valid scene transform');
  }
  if (value.width <= 0 || value.height <= 0) {
    invalidEdit('PowerPoint transform value is not a valid scene transform');
  }
  for (const key of ['flipHorizontal', 'flipVertical'] as const) {
    if (value[key] !== undefined && typeof value[key] !== 'boolean') {
      invalidEdit('PowerPoint transform value is not a valid scene transform');
    }
  }
  try {
    pointsToEmu(value.x);
    pointsToEmu(value.y);
    pointsToEmu(value.width);
    pointsToEmu(value.height);
    degreesToAngle(value.rotation ?? 0);
  } catch {
    invalidEdit('PowerPoint transform value is not a valid scene transform');
  }
  return {
    flipHorizontal: value.flipHorizontal ?? false,
    flipVertical: value.flipVertical ?? false,
    height: value.height,
    rotation: value.rotation ?? 0,
    width: value.width,
    x: value.x,
    y: value.y,
  };
}

export function normalizePptxRoundTripImageCrop(
  value: PptxSceneImageCrop | null,
): PptxSceneImageCrop | null {
  if (value === null) return null;
  if (typeof value !== 'object' || Array.isArray(value)) {
    invalidEdit('PowerPoint image crop value must be an object or null');
  }
  const keys = ['bottom', 'left', 'right', 'top'] as const;
  if (
    Object.keys(value).length !== keys.length ||
    keys.some((key) => !Object.hasOwn(value, key))
  ) {
    invalidEdit('PowerPoint image crop value has an invalid shape');
  }
  for (const key of keys) {
    const percentage = value[key];
    if (
      !isRepresentablePptxCropPercentage(percentage) ||
      percentage < -100 ||
      percentage > 100
    ) {
      invalidEdit('PowerPoint image crop value has an invalid percentage');
    }
  }
  if (value.left + value.right >= 100 || value.top + value.bottom >= 100) {
    invalidEdit('PowerPoint image crop must leave a positive visible region');
  }
  return {
    bottom: value.bottom,
    left: value.left,
    right: value.right,
    top: value.top,
  };
}

function normalizeCoordinateSpace(
  value: PptxSceneGroupTransform['childSpace'],
): PptxSceneGroupTransform['childSpace'] {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    invalidEdit('PowerPoint group child space must be an object');
  }
  const allowedKeys = new Set(['height', 'width', 'x', 'y']);
  if (
    Object.keys(value).some((key) => !allowedKeys.has(key)) ||
    value.width <= 0 ||
    value.height <= 0
  ) {
    invalidEdit('PowerPoint group child space is not valid');
  }
  try {
    pointsToEmu(value.x);
    pointsToEmu(value.y);
    pointsToEmu(value.width);
    pointsToEmu(value.height);
  } catch {
    invalidEdit('PowerPoint group child space is not valid');
  }
  return { height: value.height, width: value.width, x: value.x, y: value.y };
}

export function normalizePptxRoundTripGroupTransform(
  value: PptxSceneGroupTransform,
): PptxSceneGroupTransform {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    invalidEdit('PowerPoint group transform value must be an object');
  }
  const allowedKeys = new Set([
    'childSpace',
    'flipHorizontal',
    'flipVertical',
    'height',
    'rotation',
    'width',
    'x',
    'y',
  ]);
  if (Object.keys(value).some((key) => !allowedKeys.has(key))) {
    invalidEdit('PowerPoint group transform value is not valid');
  }
  const outer = normalizePptxRoundTripTransform({
    flipHorizontal: value.flipHorizontal ?? false,
    flipVertical: value.flipVertical ?? false,
    height: value.height,
    rotation: value.rotation ?? 0,
    width: value.width,
    x: value.x,
    y: value.y,
  });
  return { ...outer, childSpace: normalizeCoordinateSpace(value.childSpace) };
}

function validateTableTransformSize(
  target: PptxSceneTableElement,
  value: PptxSceneTransform,
): void {
  if (
    pointsToEmu(value.width) < target.columns.length ||
    pointsToEmu(value.height) < target.rows.length
  ) {
    invalidEdit(
      'PowerPoint table transform is too small for its column and row grid',
    );
  }
}

function invalidEdit(message: string): never {
  throw new PptxWriteError('invalid-edit-operation', message);
}

function findRunText(
  snapshot: PptxRoundTripSnapshot,
  targetKey: string,
): { nativeOwner: boolean; text: string } {
  let matched: { nativeOwner: boolean; text: string } | undefined;
  for (const slide of snapshot.document.slides) {
    visitEditableTextRuns(slide.elements, (run, nativeOwner) => {
      if (run.key === targetKey) {
        matched = { nativeOwner, text: run.text };
      }
    });
  }
  if (matched === undefined) {
    invalidEdit('PowerPoint text edit target key does not exist');
  }
  return matched;
}

export function validatePptxRoundTripReplaceTextRequest(
  request: PptxRoundTripReplaceTextRequest,
  maxXmlBytes: number,
): void {
  if (typeof request.targetKey !== 'string' || request.targetKey.length === 0) {
    invalidEdit('PowerPoint text edit target key must be a non-empty string');
  }
  if (typeof request.value !== 'string') {
    invalidEdit('PowerPoint text edit value must be a string');
  }
  if (!isValidXmlText(request.value)) {
    invalidEdit('PowerPoint text edit value is not safe XML text');
  }
  if (new TextEncoder().encode(request.value).byteLength > maxXmlBytes) {
    invalidEdit('PowerPoint text edit value exceeds the XML part byte limit');
  }
}

export async function replacePptxRoundTripText(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripReplaceTextRequest,
): Promise<PptxRoundTripSnapshot> {
  const limits = resolvePptxResourceLimits();
  const validated = validatePptxRoundTripSnapshot(value, limits);
  validatePptxRoundTripReplaceTextRequest(request, limits.maxXmlBytes);
  const snapshot = structuredClone(validated);
  const target = findRunText(snapshot, request.targetKey);
  const expectedText = target.text;
  if (expectedText === request.value) {
    invalidEdit('PowerPoint text edit must change the target value');
  }
  if (
    snapshot.operations.some(
      (operation) => operation.targetKey === request.targetKey,
    )
  ) {
    invalidEdit('PowerPoint text edit target is already scheduled');
  }

  const operation: PptxRoundTripReplaceTextOperation = {
    expectedText,
    id: `replace-text-${snapshot.operations.length + 1}`,
    kind: 'replace-text',
    targetKey: request.targetKey,
    value: request.value,
  };
  snapshot.operations.push(operation);
  snapshot.supportProfile =
    target.nativeOwner ||
    snapshot.supportProfile.id === 'pptx-roundtrip-native-v1'
      ? createPptxRoundTripNativeEditSupportProfile()
      : createPptxRoundTripTextEditSupportProfile();
  snapshot.consistency = await createPptxSnapshotConsistency({
    document: snapshot.document,
    operations: snapshot.operations,
    source: snapshot.source,
    supportProfile: snapshot.supportProfile,
  });
  return validatePptxRoundTripSnapshot(snapshot, limits);
}

async function setPptxRoundTripTransform(
  value: PptxRoundTripSnapshot,
  request:
    PptxRoundTripSetGroupTransformRequest | PptxRoundTripSetTransformRequest,
  targetType: PptxTransformElement['type'],
): Promise<PptxRoundTripSnapshot> {
  const limits = resolvePptxResourceLimits();
  const validated = validatePptxRoundTripSnapshot(value, limits);
  const snapshot = structuredClone(validated);
  if (typeof request.targetKey !== 'string' || request.targetKey.length === 0) {
    invalidEdit('PowerPoint transform target key must be a non-empty string');
  }
  const target = findTransformElement(snapshot, request.targetKey, targetType);
  const expectedTransform = target.resolved.transform;
  if (expectedTransform === undefined) {
    invalidEdit('PowerPoint transform target has no resolved transform');
  }
  const transform =
    target.type === 'group'
      ? normalizePptxRoundTripGroupTransform(
          request.value as PptxSceneGroupTransform,
        )
      : normalizePptxRoundTripTransform(request.value);
  if (target.type === 'table') {
    validateTableTransformSize(target, transform);
  }
  if (canonicalJson(expectedTransform) === canonicalJson(transform)) {
    invalidEdit('PowerPoint transform edit must change the target value');
  }
  if (
    snapshot.operations.some(
      (operation) =>
        operation.kind === 'set-transform' &&
        operation.targetKey === request.targetKey,
    )
  ) {
    invalidEdit('PowerPoint transform target is already scheduled');
  }
  const operation: PptxRoundTripSetTransformOperation = {
    expectedTransform: structuredClone(expectedTransform),
    id: `set-transform-${snapshot.operations.length + 1}`,
    kind: 'set-transform',
    targetKey: request.targetKey,
    value: transform,
  };
  snapshot.operations.push(operation);
  snapshot.supportProfile =
    targetType === 'image' ||
    targetType === 'chart' ||
    targetType === 'group' ||
    targetType === 'shape' ||
    targetType === 'table' ||
    snapshot.supportProfile.id === 'pptx-roundtrip-native-v1'
      ? createPptxRoundTripNativeEditSupportProfile()
      : createPptxRoundTripTextEditSupportProfile();
  snapshot.consistency = await createPptxSnapshotConsistency({
    document: snapshot.document,
    operations: snapshot.operations,
    source: snapshot.source,
    supportProfile: snapshot.supportProfile,
  });
  return validatePptxRoundTripSnapshot(snapshot, limits);
}

export function setPptxRoundTripTextTransform(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetTransformRequest,
): Promise<PptxRoundTripSnapshot> {
  return setPptxRoundTripTransform(value, request, 'text');
}

export function setPptxRoundTripShapeTransform(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetTransformRequest,
): Promise<PptxRoundTripSnapshot> {
  return setPptxRoundTripTransform(value, request, 'shape');
}

export function setPptxRoundTripImageTransform(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetTransformRequest,
): Promise<PptxRoundTripSnapshot> {
  return setPptxRoundTripTransform(value, request, 'image');
}

export async function setPptxRoundTripImageCrop(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetImageCropRequest,
): Promise<PptxRoundTripSnapshot> {
  const limits = resolvePptxResourceLimits();
  const validated = validatePptxRoundTripSnapshot(value, limits);
  const snapshot = structuredClone(validated);
  if (typeof request.targetKey !== 'string' || request.targetKey.length === 0) {
    invalidEdit('PowerPoint image crop target key must be a non-empty string');
  }
  const target = findTransformElement(
    snapshot,
    request.targetKey,
    'image',
  ) as PptxSceneImageElement;
  const expectedCrop = target.crop ?? null;
  const crop = normalizePptxRoundTripImageCrop(request.value);
  if (canonicalJson(expectedCrop) === canonicalJson(crop)) {
    invalidEdit('PowerPoint image crop edit must change the target value');
  }
  if (
    snapshot.operations.some(
      (operation) =>
        operation.kind === 'set-image-crop' &&
        operation.targetKey === request.targetKey,
    )
  ) {
    invalidEdit('PowerPoint image crop target is already scheduled');
  }
  const operation: PptxRoundTripSetImageCropOperation = {
    expectedCrop: structuredClone(expectedCrop),
    id: `set-image-crop-${snapshot.operations.length + 1}`,
    kind: 'set-image-crop',
    targetKey: request.targetKey,
    value: structuredClone(crop),
  };
  snapshot.operations.push(operation);
  snapshot.supportProfile = createPptxRoundTripNativeEditSupportProfile();
  snapshot.consistency = await createPptxSnapshotConsistency({
    document: snapshot.document,
    operations: snapshot.operations,
    source: snapshot.source,
    supportProfile: snapshot.supportProfile,
  });
  return validatePptxRoundTripSnapshot(snapshot, limits);
}

export function setPptxRoundTripTableTransform(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetTransformRequest,
): Promise<PptxRoundTripSnapshot> {
  return setPptxRoundTripTransform(value, request, 'table');
}

export function setPptxRoundTripChartTransform(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetTransformRequest,
): Promise<PptxRoundTripSnapshot> {
  return setPptxRoundTripTransform(value, request, 'chart');
}

export function setPptxRoundTripGroupTransform(
  value: PptxRoundTripSnapshot,
  request: PptxRoundTripSetGroupTransformRequest,
): Promise<PptxRoundTripSnapshot> {
  return setPptxRoundTripTransform(value, request, 'group');
}
