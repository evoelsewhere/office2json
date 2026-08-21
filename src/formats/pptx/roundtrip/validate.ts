import {
  assertPptxInputWithinLimits,
  type ResolvedPptxResourceLimits,
} from '../internal/resource-limits';
import { isValidXmlText, validatePptxScene } from '../scene-validation';
import type { PptxSceneElement, PptxSceneImageCrop } from '../scene-types';
import { PptxWriteError } from '../write-error';
import { canonicalJson } from './canonical-json';
import {
  PPTX_ROUND_TRIP_CANONICALIZATION_VERSION,
  PPTX_ROUND_TRIP_CAPABILITY_PROFILE_VERSION,
  PPTX_ROUND_TRIP_CONTRACT_VERSION,
  PPTX_ROUND_TRIP_KEY_ALGORITHM_VERSION,
  PPTX_ROUND_TRIP_NATIVE_CAPABILITY_PROFILE_VERSION,
} from './consistency';
import { assertPptxRoundTripDataTree } from './data-tree';
import type { PptxRoundTripOperation, PptxRoundTripSnapshot } from './types';

const ROOT_KEYS = [
  'consistency',
  'document',
  'format',
  'operations',
  'schemaVersion',
  'source',
  'supportProfile',
] as const;
const SOURCE_KEYS = [
  'byteLength',
  'conformance',
  'data',
  'kind',
  'sha256',
] as const;
const SUPPORT_PROFILE_KEYS = [
  'effectiveLevel',
  'id',
  'producerMatrix',
  'version',
] as const;
const CONSISTENCY_KEYS = [
  'canonicalizationVersion',
  'capabilityProfileVersion',
  'contractVersion',
  'hashAlgorithm',
  'keyAlgorithmVersion',
  'operationsSha256',
  'semanticPreviewSha256',
  'sourceManifestSha256',
] as const;
const REPLACE_TEXT_OPERATION_KEYS = [
  'expectedText',
  'id',
  'kind',
  'targetKey',
  'value',
] as const;
const SET_TRANSFORM_OPERATION_KEYS = [
  'expectedTransform',
  'id',
  'kind',
  'targetKey',
  'value',
] as const;
const SET_IMAGE_CROP_OPERATION_KEYS = [
  'expectedCrop',
  'id',
  'kind',
  'targetKey',
  'value',
] as const;
const TRANSFORM_KEYS = [
  'flipHorizontal',
  'flipVertical',
  'height',
  'rotation',
  'width',
  'x',
  'y',
] as const;
const GROUP_TRANSFORM_KEYS = ['childSpace', ...TRANSFORM_KEYS] as const;
const COORDINATE_SPACE_KEYS = ['height', 'width', 'x', 'y'] as const;
const SHA256_PATTERN = /^[0-9a-f]{64}$/;

function invalidSnapshot(message: string): never {
  throw new PptxWriteError('invalid-snapshot', message);
}

function exactRecord(
  value: unknown,
  keys: readonly string[],
  message: string,
): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    invalidSnapshot(message);
  }
  const names = Object.getOwnPropertyNames(value);
  if (
    names.length !== keys.length ||
    !keys.every((key) => Object.hasOwn(value, key))
  ) {
    invalidSnapshot(message);
  }
  return value as Record<string, unknown>;
}

function assertSha256(
  value: unknown,
  message: string,
): asserts value is string {
  if (typeof value !== 'string' || !SHA256_PATTERN.test(value)) {
    invalidSnapshot(message);
  }
}

function assertLiteral(
  value: unknown,
  expected: string | number,
  message: string,
): void {
  if (value !== expected) invalidSnapshot(message);
}

function validateSource(
  value: unknown,
  limits: ResolvedPptxResourceLimits,
): void {
  const source = exactRecord(
    value,
    SOURCE_KEYS,
    'PowerPoint round-trip snapshot source has an invalid shape',
  );
  assertLiteral(
    source.kind,
    'bytes',
    'PowerPoint round-trip snapshot source kind must be bytes',
  );
  if (!(source.data instanceof Uint8Array) && !(source.data instanceof Blob)) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot source data must be Uint8Array or Blob',
    );
  }
  assertPptxInputWithinLimits(source.data, limits);
  if (
    !Number.isSafeInteger(source.byteLength) ||
    Number(source.byteLength) <= 0
  ) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot source byteLength must be a positive safe integer',
    );
  }
  const actualByteLength =
    source.data instanceof Blob ? source.data.size : source.data.byteLength;
  if (source.byteLength !== actualByteLength) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot source byteLength does not match its data',
    );
  }
  if (
    source.conformance !== 'strict' &&
    source.conformance !== 'transitional' &&
    source.conformance !== 'unknown'
  ) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot source conformance is invalid',
    );
  }
  assertSha256(
    source.sha256,
    'PowerPoint round-trip snapshot source SHA-256 is invalid',
  );
}

type PptxRoundTripSupportProfileId =
  'pptx-roundtrip-native-v1' | 'pptx-roundtrip-r0' | 'pptx-roundtrip-text-v1';

function validateSupportProfile(
  value: unknown,
  expectedId: PptxRoundTripSupportProfileId,
): void {
  const hasOperations = expectedId !== 'pptx-roundtrip-r0';
  const profile = exactRecord(
    value,
    SUPPORT_PROFILE_KEYS,
    'PowerPoint round-trip snapshot support profile has an invalid shape',
  );
  assertLiteral(
    profile.effectiveLevel,
    hasOperations ? 'R2' : 'R0',
    hasOperations
      ? expectedId === 'pptx-roundtrip-text-v1'
        ? 'PowerPoint text edit snapshot support level must be R2'
        : 'PowerPoint native edit snapshot support level must be R2'
      : 'PowerPoint round-trip snapshot support level must be R0',
  );
  assertLiteral(
    profile.id,
    expectedId,
    'PowerPoint round-trip snapshot support profile id is unsupported',
  );
  assertLiteral(
    profile.version,
    '1',
    'PowerPoint round-trip snapshot support profile version is unsupported',
  );
  if (!Array.isArray(profile.producerMatrix)) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot producer matrix must be an array',
    );
  }
  if (profile.producerMatrix.length !== 0) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot cannot claim unverified producer evidence',
    );
  }
}

interface EditableRun {
  nativeOwner: boolean;
  text: string;
}

function editableRuns(
  value: PptxRoundTripSnapshot['document'],
): Map<string, EditableRun> {
  const runs = new Map<string, EditableRun>();
  const collectText = (
    text: Extract<PptxSceneElement, { type: 'text' }>['text'],
    nativeOwner: boolean,
  ): void => {
    for (const paragraph of text.paragraphs) {
      for (const child of paragraph.children) {
        if (child.type === 'run') {
          runs.set(child.key, { nativeOwner, text: child.text });
        }
      }
    }
  };
  const collect = (
    elements: readonly PptxSceneElement[],
    nestedInGroup = false,
  ): void => {
    for (const element of elements) {
      if (element.type === 'text') {
        collectText(element.text, nestedInGroup);
      } else if (element.type === 'table') {
        for (const row of element.rows) {
          for (const cell of row.cells) collectText(cell.text, true);
        }
      } else if (element.type === 'group') {
        collect(element.elements, true);
      }
    }
  };
  for (const slide of value.slides) {
    collect(slide.elements);
  }
  return runs;
}

function editableTransforms(
  value: PptxRoundTripSnapshot['document'],
): Map<string, unknown> {
  const transforms = new Map<string, unknown>();
  const collect = (elements: readonly PptxSceneElement[]): void => {
    for (const element of elements) {
      if (
        element.type === 'chart' ||
        element.type === 'image' ||
        element.type === 'group' ||
        element.type === 'shape' ||
        element.type === 'table' ||
        element.type === 'text'
      ) {
        transforms.set(element.key, element.resolved.transform);
      }
      if (element.type === 'group') collect(element.elements);
    }
  };
  for (const slide of value.slides) {
    collect(slide.elements);
  }
  return transforms;
}

function editableImageCrops(
  value: PptxRoundTripSnapshot['document'],
): Map<string, PptxSceneImageCrop | null> {
  const crops = new Map<string, PptxSceneImageCrop | null>();
  const collect = (elements: readonly PptxSceneElement[]): void => {
    for (const element of elements) {
      if (element.type === 'image')
        crops.set(element.key, element.crop ?? null);
      if (element.type === 'group') collect(element.elements);
    }
  };
  for (const slide of value.slides) collect(slide.elements);
  return crops;
}

function validateImageCrop(
  value: unknown,
  message: string,
): PptxSceneImageCrop | null {
  if (value === null) return null;
  const keys = ['bottom', 'left', 'right', 'top'] as const;
  const crop = exactRecord(value, keys, message);
  for (const key of keys) {
    const percentage = crop[key];
    if (
      typeof percentage !== 'number' ||
      !Number.isFinite(percentage) ||
      percentage < -100 ||
      percentage > 100 ||
      !Number.isSafeInteger(percentage * 1_000)
    ) {
      invalidSnapshot(message);
    }
  }
  if (
    Number(crop.left) + Number(crop.right) >= 100 ||
    Number(crop.top) + Number(crop.bottom) >= 100
  ) {
    invalidSnapshot(message);
  }
  return crop as unknown as PptxSceneImageCrop;
}

function validateTransform(
  value: unknown,
  message: string,
  groupTransform: boolean,
): Record<string, unknown> {
  const transform = exactRecord(
    value,
    groupTransform ? GROUP_TRANSFORM_KEYS : TRANSFORM_KEYS,
    message,
  );
  for (const key of ['x', 'y', 'width', 'height', 'rotation'] as const) {
    if (!Number.isFinite(transform[key])) {
      invalidSnapshot(message);
    }
  }
  if (Number(transform.width) <= 0 || Number(transform.height) <= 0) {
    invalidSnapshot(message);
  }
  for (const key of ['flipHorizontal', 'flipVertical'] as const) {
    if (typeof transform[key] !== 'boolean') invalidSnapshot(message);
  }
  if (groupTransform) {
    const childSpace = exactRecord(
      transform.childSpace,
      COORDINATE_SPACE_KEYS,
      message,
    );
    for (const key of COORDINATE_SPACE_KEYS) {
      if (!Number.isFinite(childSpace[key])) invalidSnapshot(message);
    }
    if (Number(childSpace.width) <= 0 || Number(childSpace.height) <= 0) {
      invalidSnapshot(message);
    }
  }
  return transform;
}

function validateOperations(
  values: unknown[],
  document: PptxRoundTripSnapshot['document'],
  limits: ResolvedPptxResourceLimits,
): void {
  const runs = editableRuns(document);
  const transforms = editableTransforms(document);
  const imageCrops = editableImageCrops(document);
  const ids = new Set<string>();
  const targets = new Set<string>();
  for (const [index, value] of values.entries()) {
    if (value === null || typeof value !== 'object' || Array.isArray(value)) {
      invalidSnapshot(
        `PowerPoint round-trip operation ${index + 1} has an invalid shape`,
      );
    }
    const kind = (value as Record<string, unknown>).kind;
    if (
      kind !== 'replace-text' &&
      kind !== 'set-image-crop' &&
      kind !== 'set-transform'
    ) {
      invalidSnapshot(
        `PowerPoint round-trip operation ${index + 1} kind is unsupported`,
      );
    }
    const operation = exactRecord(
      value,
      kind === 'replace-text'
        ? REPLACE_TEXT_OPERATION_KEYS
        : kind === 'set-image-crop'
          ? SET_IMAGE_CROP_OPERATION_KEYS
          : SET_TRANSFORM_OPERATION_KEYS,
      `PowerPoint round-trip operation ${index + 1} has an invalid shape`,
    );
    for (const key of ['id', 'targetKey'] as const) {
      if (typeof operation[key] !== 'string') {
        invalidSnapshot(
          `PowerPoint round-trip operation ${index + 1} ${key} must be a string`,
        );
      }
    }
    const id = operation.id as string;
    const targetKey = operation.targetKey as string;
    const targetIdentity = `${kind}\u0000${targetKey}`;
    if (id.length === 0 || ids.has(id)) {
      invalidSnapshot('PowerPoint round-trip operation ids must be unique');
    }
    if (targetKey.length === 0 || targets.has(targetIdentity)) {
      invalidSnapshot(
        'PowerPoint round-trip text edit targets must be non-empty and unique',
      );
    }
    ids.add(id);
    targets.add(targetIdentity);
    if (kind === 'set-image-crop') {
      if (!imageCrops.has(targetKey)) {
        invalidSnapshot(
          'PowerPoint round-trip image crop target does not exist',
        );
      }
      const sourceCrop = imageCrops.get(targetKey) as PptxSceneImageCrop | null;
      const expectedCrop = validateImageCrop(
        operation.expectedCrop,
        `PowerPoint round-trip operation ${index + 1} expectedCrop is invalid`,
      );
      const replacement = validateImageCrop(
        operation.value,
        `PowerPoint round-trip operation ${index + 1} value is not a valid image crop`,
      );
      if (canonicalJson(sourceCrop) !== canonicalJson(expectedCrop)) {
        invalidSnapshot(
          'PowerPoint round-trip image crop precondition does not match the preview',
        );
      }
      if (canonicalJson(replacement) === canonicalJson(expectedCrop)) {
        invalidSnapshot(
          'PowerPoint round-trip image crop must change the value',
        );
      }
      continue;
    }
    if (kind === 'set-transform') {
      const sourceTransform = transforms.get(targetKey);
      if (sourceTransform === undefined) {
        invalidSnapshot(
          'PowerPoint round-trip transform target does not exist',
        );
      }
      const groupTransform = Object.hasOwn(
        sourceTransform as object,
        'childSpace',
      );
      const expectedTransform = validateTransform(
        operation.expectedTransform,
        `PowerPoint round-trip operation ${index + 1} expectedTransform is invalid`,
        groupTransform,
      );
      const replacement = validateTransform(
        operation.value,
        `PowerPoint round-trip operation ${index + 1} value is not a valid transform`,
        groupTransform,
      );
      if (canonicalJson(sourceTransform) !== canonicalJson(expectedTransform)) {
        invalidSnapshot(
          'PowerPoint round-trip transform precondition does not match the preview',
        );
      }
      if (canonicalJson(replacement) === canonicalJson(expectedTransform)) {
        invalidSnapshot(
          'PowerPoint round-trip transform must change the value',
        );
      }
      continue;
    }
    for (const key of ['expectedText', 'value'] as const) {
      if (typeof operation[key] !== 'string') {
        invalidSnapshot(
          `PowerPoint round-trip operation ${index + 1} ${key} must be a string`,
        );
      }
    }
    const expectedText = operation.expectedText as string;
    const replacement = operation.value as string;
    const sourceRun = runs.get(targetKey);
    if (sourceRun === undefined) {
      invalidSnapshot('PowerPoint round-trip text edit target does not exist');
    }
    if (sourceRun.text !== expectedText) {
      invalidSnapshot(
        'PowerPoint round-trip text edit precondition does not match the preview',
      );
    }
    if (replacement === expectedText) {
      invalidSnapshot('PowerPoint round-trip text edit must change the value');
    }
    if (!isValidXmlText(replacement)) {
      invalidSnapshot(
        'PowerPoint round-trip text edit value is not safe XML text',
      );
    }
    if (new TextEncoder().encode(replacement).byteLength > limits.maxXmlBytes) {
      invalidSnapshot(
        'PowerPoint round-trip text edit value exceeds the XML part byte limit',
      );
    }
  }
}

function validateConsistency(
  value: unknown,
  capabilityProfileVersion: string,
): void {
  const consistency = exactRecord(
    value,
    CONSISTENCY_KEYS,
    'PowerPoint round-trip snapshot consistency has an invalid shape',
  );
  const literals: ReadonlyArray<readonly [unknown, string, string]> = [
    [
      consistency.canonicalizationVersion,
      PPTX_ROUND_TRIP_CANONICALIZATION_VERSION,
      'PowerPoint round-trip snapshot canonicalization version is unsupported',
    ],
    [
      consistency.capabilityProfileVersion,
      capabilityProfileVersion,
      'PowerPoint round-trip snapshot capability profile version is unsupported',
    ],
    [
      consistency.contractVersion,
      PPTX_ROUND_TRIP_CONTRACT_VERSION,
      'PowerPoint round-trip snapshot contract version is unsupported',
    ],
    [
      consistency.hashAlgorithm,
      'sha256',
      'PowerPoint round-trip snapshot hash algorithm is unsupported',
    ],
    [
      consistency.keyAlgorithmVersion,
      PPTX_ROUND_TRIP_KEY_ALGORITHM_VERSION,
      'PowerPoint round-trip snapshot key algorithm version is unsupported',
    ],
  ];
  for (const [actual, expected, message] of literals) {
    assertLiteral(actual, expected, message);
  }
  assertSha256(
    consistency.operationsSha256,
    'PowerPoint round-trip snapshot operations SHA-256 is invalid',
  );
  assertSha256(
    consistency.semanticPreviewSha256,
    'PowerPoint round-trip snapshot semantic preview SHA-256 is invalid',
  );
  assertSha256(
    consistency.sourceManifestSha256,
    'PowerPoint round-trip snapshot source manifest SHA-256 is invalid',
  );
}

function expectedSupportProfileId(
  operations: readonly PptxRoundTripOperation[],
  document: PptxRoundTripSnapshot['document'],
): PptxRoundTripSupportProfileId {
  if (operations.length === 0) return 'pptx-roundtrip-r0';
  const nativeKeys = new Set<string>();
  const collectNativeKeys = (elements: readonly PptxSceneElement[]): void => {
    for (const element of elements) {
      if (
        element.type === 'chart' ||
        element.type === 'image' ||
        element.type === 'group' ||
        element.type === 'shape' ||
        element.type === 'table'
      ) {
        nativeKeys.add(element.key);
      }
      if (element.type === 'group') collectNativeKeys(element.elements);
    }
  };
  for (const slide of document.slides) {
    collectNativeKeys(slide.elements);
  }
  const runs = editableRuns(document);
  return operations.some(
    (operation) =>
      nativeKeys.has(operation.targetKey) ||
      runs.get(operation.targetKey)?.nativeOwner === true,
  )
    ? 'pptx-roundtrip-native-v1'
    : 'pptx-roundtrip-text-v1';
}

export function validatePptxRoundTripSnapshot(
  value: unknown,
  limits: ResolvedPptxResourceLimits,
): PptxRoundTripSnapshot {
  assertPptxRoundTripDataTree(value, limits);
  const snapshot = exactRecord(
    value,
    ROOT_KEYS,
    'PowerPoint round-trip snapshot has an invalid root shape',
  );
  assertLiteral(
    snapshot.schemaVersion,
    1,
    'PowerPoint round-trip snapshot schema version is unsupported',
  );
  assertLiteral(
    snapshot.format,
    'pptx',
    'PowerPoint round-trip snapshot format must be pptx',
  );
  if (!Array.isArray(snapshot.operations)) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot operations must be an array',
    );
  }

  validateSource(snapshot.source, limits);
  const sceneValidation = validatePptxScene(snapshot.document);
  if (!sceneValidation.valid) {
    throw new PptxWriteError(
      'invalid-snapshot',
      'PowerPoint round-trip snapshot semantic preview is invalid',
      { issues: sceneValidation.issues },
    );
  }
  validateOperations(
    snapshot.operations as PptxRoundTripOperation[],
    snapshot.document as PptxRoundTripSnapshot['document'],
    limits,
  );
  const supportProfileId = expectedSupportProfileId(
    snapshot.operations,
    snapshot.document as PptxRoundTripSnapshot['document'],
  );
  validateSupportProfile(snapshot.supportProfile, supportProfileId);
  validateConsistency(
    snapshot.consistency,
    supportProfileId === 'pptx-roundtrip-native-v1'
      ? PPTX_ROUND_TRIP_NATIVE_CAPABILITY_PROFILE_VERSION
      : PPTX_ROUND_TRIP_CAPABILITY_PROFILE_VERSION,
  );

  return snapshot as unknown as PptxRoundTripSnapshot;
}
