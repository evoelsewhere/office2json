import type {
  PptxSceneValidationCode,
  PptxSceneValidationIssue,
  PptxSceneValidationOptions,
  PptxSceneValidationResult,
} from './scene-types';
import {
  isSupportedPowerPointCreationSlideCount,
  MAX_POWERPOINT_CREATION_CHART_POINTS,
  MAX_POWERPOINT_CREATION_CHART_SERIES,
  MAX_POWERPOINT_CREATION_SLIDES,
} from './creation-limits';
import { validatePowerPointCreationResources } from './creation-resource-validation';
import {
  validatePptxSceneTable,
  validatePptxSceneTableDimensions,
  type PptxTableValidationDependencies,
} from './scene-table-validation';

type JsonObject = Record<string, unknown>;

const KEY_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/;
const EMUS_PER_POINT = 12_700;
const ANGLE_UNITS_PER_DEGREE = 60_000;
const FONT_SIZE_UNITS_PER_POINT = 100;
const HEX_COLOR_PATTERN = /^#[0-9a-f]{6}$/i;
const CHART_TYPES = [
  'barChart',
  'doughnutChart',
  'lineChart',
  'pieChart',
] as const;

type ValidationProfile = NonNullable<PptxSceneValidationOptions['profile']>;
type CreationValidationProfile = Exclude<ValidationProfile, 'scene'>;

export function isRepresentablePptxCropPercentage(value: unknown): boolean {
  const percentage = value as number;
  const rounded = Math.round(percentage * 1_000);
  return Number.isSafeInteger(rounded) && rounded / 1_000 === percentage;
}

function isCreationProfile(
  profile: ValidationProfile,
): profile is CreationValidationProfile {
  return profile === 'create-native-v1' || profile === 'create-text-v1';
}

function isObject(value: unknown): value is JsonObject {
  if (value === null) return false;
  if (value === undefined) return false;
  return Object.getPrototypeOf(value) === Object.prototype;
}

function addIssue(
  issues: PptxSceneValidationIssue[],
  code: PptxSceneValidationCode,
  path: string,
  message: string,
): void {
  issues.push({ code, message, path });
}

function rejectUnknownKeys(
  value: JsonObject,
  allowed: readonly string[],
  path: string,
  issues: PptxSceneValidationIssue[],
): void {
  const allowedKeys = new Set(allowed);
  for (const key of Object.keys(value)) {
    if (!allowedKeys.has(key)) {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.${key}`,
        'Unknown property',
      );
    }
  }
}

function requireObject(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
): JsonObject | undefined {
  if (isObject(value)) return value;
  addIssue(issues, 'invalid-scene-document', path, 'Expected an object');
  return undefined;
}

function requireArray(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
): unknown[] | undefined {
  if (Array.isArray(value)) return value as unknown[];
  addIssue(issues, 'invalid-scene-document', path, 'Expected an array');
  return undefined;
}

function isOneOf(value: unknown, allowed: readonly string[]): value is string {
  return typeof value === 'string' && allowed.includes(value);
}

function optionalString(
  value: JsonObject,
  key: string,
  path: string,
  issues: PptxSceneValidationIssue[],
): void {
  if (value[key] === undefined || typeof value[key] === 'string') return;
  addIssue(
    issues,
    'invalid-scene-document',
    `${path}.${key}`,
    'Expected a string',
  );
}

function optionalBoolean(
  value: JsonObject,
  key: string,
  path: string,
  issues: PptxSceneValidationIssue[],
): void {
  if (value[key] === undefined || typeof value[key] === 'boolean') return;
  addIssue(
    issues,
    'invalid-scene-document',
    `${path}.${key}`,
    'Expected a boolean',
  );
}

function optionalColor(
  value: JsonObject,
  key: string,
  path: string,
  issues: PptxSceneValidationIssue[],
): void {
  const color = value[key];
  if (color === undefined) return;
  if (typeof color === 'string' && HEX_COLOR_PATTERN.test(color)) return;
  addIssue(
    issues,
    'invalid-scene-document',
    `${path}.${key}`,
    'Expected a #RRGGBB color',
  );
}

function requireFiniteNumber(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
  positive: boolean,
): void {
  if (
    typeof value === 'number' &&
    Number.isFinite(value) &&
    (!positive || value > 0)
  ) {
    return;
  }
  addIssue(
    issues,
    'invalid-numeric-value',
    path,
    positive ? 'Expected a positive finite number' : 'Expected a finite number',
  );
}

function requireSerializableInteger(
  value: unknown,
  multiplier: number,
  path: string,
  issues: PptxSceneValidationIssue[],
  positive: boolean,
): void {
  if (typeof value !== 'number' || !Number.isFinite(value)) return;
  const result = Math.round(value * multiplier);
  if (!Number.isSafeInteger(result)) {
    addIssue(
      issues,
      'invalid-numeric-value',
      path,
      'Value exceeds the safe OOXML integer range',
    );
  } else if (positive && result <= 0) {
    addIssue(
      issues,
      'invalid-numeric-value',
      path,
      'Value must round to a positive OOXML integer',
    );
  }
}

function validateSize(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
  profile: ValidationProfile,
): void {
  const size = requireObject(value, path, issues);
  if (!size) return;
  rejectUnknownKeys(size, ['height', 'width'], path, issues);
  requireFiniteNumber(size.width, `${path}.width`, issues, true);
  requireFiniteNumber(size.height, `${path}.height`, issues, true);
  if (isCreationProfile(profile)) {
    requireSerializableInteger(
      size.width,
      EMUS_PER_POINT,
      `${path}.width`,
      issues,
      true,
    );
    requireSerializableInteger(
      size.height,
      EMUS_PER_POINT,
      `${path}.height`,
      issues,
      true,
    );
  }
}

function validateTransform(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
  profile: ValidationProfile,
  groupTransform: boolean,
  coordinateSpace = false,
): void {
  const transform = requireObject(value, path, issues);
  if (!transform) return;
  rejectUnknownKeys(
    transform,
    coordinateSpace
      ? ['height', 'width', 'x', 'y']
      : [
          ...(groupTransform ? ['childSpace'] : []),
          'flipHorizontal',
          'flipVertical',
          'height',
          'rotation',
          'width',
          'x',
          'y',
        ],
    path,
    issues,
  );
  requireFiniteNumber(transform.x, `${path}.x`, issues, false);
  requireFiniteNumber(transform.y, `${path}.y`, issues, false);
  requireFiniteNumber(transform.width, `${path}.width`, issues, true);
  requireFiniteNumber(transform.height, `${path}.height`, issues, true);
  if (!coordinateSpace && transform.rotation !== undefined) {
    requireFiniteNumber(transform.rotation, `${path}.rotation`, issues, false);
  }
  if (!coordinateSpace) {
    optionalBoolean(transform, 'flipHorizontal', path, issues);
    optionalBoolean(transform, 'flipVertical', path, issues);
  }
  if (isCreationProfile(profile)) {
    for (const key of ['x', 'y'] as const) {
      requireSerializableInteger(
        transform[key],
        EMUS_PER_POINT,
        `${path}.${key}`,
        issues,
        false,
      );
    }
    for (const key of ['height', 'width'] as const) {
      requireSerializableInteger(
        transform[key],
        EMUS_PER_POINT,
        `${path}.${key}`,
        issues,
        true,
      );
    }
    if (!coordinateSpace) {
      requireSerializableInteger(
        transform.rotation,
        ANGLE_UNITS_PER_DEGREE,
        `${path}.rotation`,
        issues,
        false,
      );
    }
  }
  if (groupTransform) {
    validateTransform(
      transform.childSpace,
      `${path}.childSpace`,
      issues,
      profile,
      false,
      true,
    );
  }
}

function registerKey(
  value: unknown,
  path: string,
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
): string | undefined {
  if (typeof value !== 'string' || !KEY_PATTERN.test(value)) {
    addIssue(
      issues,
      'invalid-scene-document',
      path,
      'Expected a non-empty portable key of at most 128 characters',
    );
    return undefined;
  }
  if (keys.has(value)) {
    addIssue(
      issues,
      'duplicate-public-key',
      path,
      `Duplicate public key: ${value}`,
    );
    return value;
  }
  keys.add(value);
  return value;
}

export function isValidXmlText(value: string): boolean {
  for (const character of value) {
    if (character.length === 2) continue;
    const codeUnit = character.charCodeAt(0);
    if (
      (codeUnit < 0x20 &&
        codeUnit !== 0x09 &&
        codeUnit !== 0x0a &&
        codeUnit !== 0x0d) ||
      (codeUnit >= 0xd800 && codeUnit <= 0xdfff) ||
      codeUnit === 0xfffe ||
      codeUnit === 0xffff
    ) {
      return false;
    }
  }
  return true;
}

function validateTextValue(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
): void {
  if (typeof value !== 'string') {
    addIssue(issues, 'invalid-scene-document', path, 'Expected text content');
  } else if (!isValidXmlText(value)) {
    addIssue(
      issues,
      'invalid-office-text-escape',
      path,
      'Text contains a character that cannot be serialized safely',
    );
  }
}

function validateRunProperties(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
  profile: ValidationProfile,
): void {
  const properties = requireObject(value, path, issues);
  if (!properties) return;
  rejectUnknownKeys(
    properties,
    ['bold', 'color', 'fontFamily', 'fontSize', 'italic', 'language'],
    path,
    issues,
  );
  optionalBoolean(properties, 'bold', path, issues);
  optionalBoolean(properties, 'italic', path, issues);
  optionalColor(properties, 'color', path, issues);
  optionalString(properties, 'fontFamily', path, issues);
  optionalString(properties, 'language', path, issues);
  if (properties.fontSize !== undefined) {
    requireFiniteNumber(properties.fontSize, `${path}.fontSize`, issues, true);
    if (isCreationProfile(profile)) {
      requireSerializableInteger(
        properties.fontSize,
        FONT_SIZE_UNITS_PER_POINT,
        `${path}.fontSize`,
        issues,
        true,
      );
    }
  }
}

function validateTextNode(
  value: unknown,
  path: string,
  profile: ValidationProfile,
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
): void {
  const node = requireObject(value, path, issues);
  if (!node) return;
  registerKey(node.key, `${path}.key`, keys, issues);
  if (node.type === 'run') {
    rejectUnknownKeys(
      node,
      ['key', 'preserveSpace', 'properties', 'text', 'type'],
      path,
      issues,
    );
    validateTextValue(node.text, `${path}.text`, issues);
    optionalBoolean(node, 'preserveSpace', path, issues);
  } else if (node.type === 'field') {
    rejectUnknownKeys(
      node,
      ['fieldType', 'key', 'properties', 'text', 'type'],
      path,
      issues,
    );
    validateTextValue(node.text, `${path}.text`, issues);
    if (typeof node.fieldType !== 'string' || node.fieldType.trim() === '') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.fieldType`,
        'Expected a non-empty field type',
      );
    }
  } else if (node.type === 'break') {
    rejectUnknownKeys(node, ['key', 'properties', 'type'], path, issues);
  } else {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.type`,
      'Unknown text node type',
    );
  }
  if (node.properties !== undefined) {
    validateRunProperties(
      node.properties,
      `${path}.properties`,
      issues,
      profile,
    );
  }
}

function validateParagraph(
  value: unknown,
  path: string,
  profile: ValidationProfile,
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
): void {
  const paragraph = requireObject(value, path, issues);
  if (!paragraph) return;
  rejectUnknownKeys(
    paragraph,
    ['children', 'endProperties', 'key', 'properties'],
    path,
    issues,
  );
  registerKey(paragraph.key, `${path}.key`, keys, issues);
  const children = requireArray(paragraph.children, `${path}.children`, issues);
  children?.forEach((child, index) =>
    validateTextNode(
      child,
      `${path}.children[${index}]`,
      profile,
      keys,
      issues,
    ),
  );
  if (paragraph.endProperties !== undefined) {
    validateRunProperties(
      paragraph.endProperties,
      `${path}.endProperties`,
      issues,
      profile,
    );
  }
  if (paragraph.properties !== undefined) {
    const properties = requireObject(
      paragraph.properties,
      `${path}.properties`,
      issues,
    );
    if (properties) {
      rejectUnknownKeys(
        properties,
        ['alignment', 'level'],
        `${path}.properties`,
        issues,
      );
      if (
        properties.alignment !== undefined &&
        !isOneOf(properties.alignment, [
          'center',
          'distributed',
          'justify',
          'left',
          'right',
        ])
      ) {
        addIssue(
          issues,
          'invalid-scene-document',
          `${path}.properties.alignment`,
          'Unknown paragraph alignment',
        );
      }
      if (
        properties.level !== undefined &&
        (!Number.isSafeInteger(properties.level) ||
          Number(properties.level) < 0 ||
          Number(properties.level) > 8)
      ) {
        addIssue(
          issues,
          'invalid-numeric-value',
          `${path}.properties.level`,
          'Paragraph level must be an integer from 0 through 8',
        );
      }
    }
  }
}

function validateTextBody(
  value: unknown,
  path: string,
  profile: ValidationProfile,
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
): void {
  const text = requireObject(value, path, issues);
  if (!text) return;
  rejectUnknownKeys(text, ['body', 'paragraphs'], path, issues);
  const body = requireObject(text.body, `${path}.body`, issues);
  if (body) {
    rejectUnknownKeys(
      body,
      ['anchor', 'autoFit', 'vertical', 'wrap'],
      `${path}.body`,
      issues,
    );
    optionalBoolean(body, 'vertical', `${path}.body`, issues);
    optionalBoolean(body, 'wrap', `${path}.body`, issues);
    if (
      body.anchor !== undefined &&
      !isOneOf(body.anchor, [
        'bottom',
        'center',
        'distributed',
        'justified',
        'top',
      ])
    ) {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.body.anchor`,
        'Unknown text anchor',
      );
    }
    if (
      body.autoFit !== undefined &&
      !isOneOf(body.autoFit, ['none', 'shape', 'text'])
    ) {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.body.autoFit`,
        'Unknown text auto-fit mode',
      );
    }
  }
  const paragraphs = requireArray(
    text.paragraphs,
    `${path}.paragraphs`,
    issues,
  );
  if (paragraphs?.length === 0) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.paragraphs`,
      'A text body needs at least one paragraph',
    );
  }
  paragraphs?.forEach((paragraph, index) =>
    validateParagraph(
      paragraph,
      `${path}.paragraphs[${index}]`,
      profile,
      keys,
      issues,
    ),
  );
}

const TABLE_VALIDATION_DEPENDENCIES: PptxTableValidationDependencies = {
  addIssue,
  isCreationProfile,
  isObject,
  optionalBoolean,
  optionalColor,
  rejectUnknownKeys,
  requireArray,
  requireFiniteNumber,
  requireObject,
  requireSerializableInteger,
  validateTextBody,
};

function validateChartSeries(
  value: unknown,
  path: string,
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
): void {
  const series = requireArray(value, path, issues);
  if (!series) return;
  if (series.length === 0) {
    addIssue(
      issues,
      'invalid-scene-document',
      path,
      'A chart needs at least one series',
    );
  }
  if (series.length > MAX_POWERPOINT_CREATION_CHART_SERIES) {
    addIssue(
      issues,
      'resource-limit-exceeded',
      path,
      `A chart supports at most ${MAX_POWERPOINT_CREATION_CHART_SERIES} series`,
    );
  }
  series.forEach((item, seriesIndex) => {
    const itemPath = `${path}[${seriesIndex}]`;
    const record = requireObject(item, itemPath, issues);
    if (!record) return;
    rejectUnknownKeys(
      record,
      ['categories', 'color', 'key', 'name', 'values'],
      itemPath,
      issues,
    );
    registerKey(record.key, `${itemPath}.key`, keys, issues);
    if (typeof record.name !== 'string' || record.name.trim() === '') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${itemPath}.name`,
        'Expected a non-empty chart series name',
      );
    } else if (!isValidXmlText(record.name)) {
      addIssue(
        issues,
        'invalid-office-text-escape',
        `${itemPath}.name`,
        'Chart series name cannot be serialized safely',
      );
    }
    optionalColor(record, 'color', itemPath, issues);
    const categories = requireArray(
      record.categories,
      `${itemPath}.categories`,
      issues,
    );
    const values = requireArray(record.values, `${itemPath}.values`, issues);
    if (categories && values && categories.length !== values.length) {
      addIssue(
        issues,
        'invalid-scene-document',
        itemPath,
        'Chart categories and values must have equal lengths',
      );
    }
    if (
      (categories?.length ?? 0) > MAX_POWERPOINT_CREATION_CHART_POINTS ||
      (values?.length ?? 0) > MAX_POWERPOINT_CREATION_CHART_POINTS
    ) {
      addIssue(
        issues,
        'resource-limit-exceeded',
        itemPath,
        `A chart series supports at most ${MAX_POWERPOINT_CREATION_CHART_POINTS} points`,
      );
    }
    categories?.forEach((category, pointIndex) =>
      validateTextValue(
        category,
        `${itemPath}.categories[${pointIndex}]`,
        issues,
      ),
    );
    values?.forEach((point, pointIndex) =>
      requireFiniteNumber(
        point,
        `${itemPath}.values[${pointIndex}]`,
        issues,
        false,
      ),
    );
  });
}

function validateChartElement(
  element: JsonObject,
  path: string,
  baseKeys: readonly string[],
  profile: ValidationProfile,
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
): void {
  rejectUnknownKeys(
    element,
    [
      ...baseKeys,
      'barDirection',
      'chartType',
      'grouping',
      'holeSize',
      'marker',
      'series',
    ],
    path,
    issues,
  );
  if (!isOneOf(element.chartType, CHART_TYPES)) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.chartType`,
      'Unknown native chart type',
    );
  }
  validateChartSeries(element.series, `${path}.series`, keys, issues);
  if (
    element.barDirection !== undefined &&
    !isOneOf(element.barDirection, ['bar', 'col'])
  ) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.barDirection`,
      'Unknown chart bar direction',
    );
  }
  if (
    element.grouping !== undefined &&
    !isOneOf(element.grouping, [
      'clustered',
      'percentStacked',
      'stacked',
      'standard',
    ])
  ) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.grouping`,
      'Unknown chart grouping',
    );
  }
  optionalBoolean(element, 'marker', path, issues);
  if (
    element.holeSize !== undefined &&
    (!Number.isSafeInteger(element.holeSize) ||
      Number(element.holeSize) < 10 ||
      Number(element.holeSize) > 90)
  ) {
    addIssue(
      issues,
      'invalid-numeric-value',
      `${path}.holeSize`,
      'Doughnut hole size must be an integer from 10 through 90',
    );
  }
  if (
    (element.chartType === 'pieChart' ||
      element.chartType === 'doughnutChart') &&
    Array.isArray(element.series) &&
    element.series.length !== 1
  ) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.series`,
      'Pie and doughnut charts require exactly one series',
    );
  }
  const incompatible =
    (element.barDirection !== undefined && element.chartType !== 'barChart') ||
    (element.holeSize !== undefined && element.chartType !== 'doughnutChart') ||
    (element.marker !== undefined && element.chartType !== 'lineChart') ||
    (element.grouping !== undefined &&
      element.chartType !== 'barChart' &&
      element.chartType !== 'lineChart');
  if (incompatible) {
    addIssue(
      issues,
      'unsupported-feature',
      path,
      'Chart options must match the selected native chart type',
    );
  }
  if (profile === 'create-text-v1') {
    addIssue(
      issues,
      'unsupported-feature',
      path,
      'Creation profile create-text-v1 supports text elements only',
    );
  }
  const authored = isObject(element.authored) ? element.authored : undefined;
  const transform = isObject(authored?.transform)
    ? authored.transform
    : undefined;
  if (
    profile === 'create-native-v1' &&
    transform !== undefined &&
    (transform.flipHorizontal === true ||
      transform.flipVertical === true ||
      (transform.rotation !== undefined && transform.rotation !== 0))
  ) {
    addIssue(
      issues,
      'unsupported-feature',
      `${path}.authored.transform`,
      'Native chart creation supports unrotated, unflipped graphic frames only',
    );
  }
}

function validateImageCrop(
  value: unknown,
  path: string,
  profile: ValidationProfile,
  issues: PptxSceneValidationIssue[],
): void {
  const crop = requireObject(value, path, issues);
  if (!crop) return;
  const keys = ['bottom', 'left', 'right', 'top'] as const;
  rejectUnknownKeys(crop, keys, path, issues);
  const issueCount = issues.length;
  for (const key of keys) {
    const percentage = crop[key];
    if (!isRepresentablePptxCropPercentage(percentage)) {
      addIssue(
        issues,
        'invalid-numeric-value',
        `${path}.${key}`,
        'Image crop must be a finite percentage with at most three decimal places',
      );
    } else if (
      profile !== 'scene' &&
      ((percentage as number) < -100 || (percentage as number) > 100)
    ) {
      addIssue(
        issues,
        'invalid-numeric-value',
        `${path}.${key}`,
        'Native image crop must be from -100 through 100',
      );
    }
  }
  if (issues.length !== issueCount || profile === 'scene') return;
  if (Number(crop.left) + Number(crop.right) >= 100) {
    addIssue(
      issues,
      'invalid-numeric-value',
      path,
      'Horizontal image crop must leave a positive visible region',
    );
  }
  if (Number(crop.top) + Number(crop.bottom) >= 100) {
    addIssue(
      issues,
      'invalid-numeric-value',
      path,
      'Vertical image crop must leave a positive visible region',
    );
  }
}

function validatePlaceholder(
  value: unknown,
  path: string,
  referenceKeys: Array<{ path: string; value: string }>,
  issues: PptxSceneValidationIssue[],
): void {
  const placeholder = requireObject(value, path, issues);
  if (!placeholder) return;
  rejectUnknownKeys(
    placeholder,
    [
      'hasCustomPrompt',
      'index',
      'orientation',
      'prompt',
      'role',
      'size',
      'sourceKey',
      'type',
    ],
    path,
    issues,
  );
  optionalBoolean(placeholder, 'hasCustomPrompt', path, issues);
  optionalString(placeholder, 'prompt', path, issues);
  optionalString(placeholder, 'type', path, issues);
  if (
    placeholder.index !== undefined &&
    (!Number.isSafeInteger(placeholder.index) || Number(placeholder.index) < 0)
  ) {
    addIssue(
      issues,
      'invalid-numeric-value',
      `${path}.index`,
      'Placeholder index must be a non-negative integer',
    );
  }
  if (
    !isOneOf(placeholder.role, [
      'layout-definition',
      'master-definition',
      'slide-instance',
    ])
  ) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.role`,
      'Unknown placeholder role',
    );
  }
  if (
    placeholder.orientation !== undefined &&
    !isOneOf(placeholder.orientation, ['horizontal', 'vertical'])
  ) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.orientation`,
      'Unknown placeholder orientation',
    );
  }
  if (
    placeholder.size !== undefined &&
    !isOneOf(placeholder.size, ['full', 'half', 'quarter'])
  ) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.size`,
      'Unknown placeholder size',
    );
  }
  if (placeholder.sourceKey !== undefined) {
    if (typeof placeholder.sourceKey === 'string') {
      referenceKeys.push({
        path: `${path}.sourceKey`,
        value: placeholder.sourceKey,
      });
    } else {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.sourceKey`,
        'Expected a public key',
      );
    }
  }
}

function validateElement(
  value: unknown,
  path: string,
  profile: ValidationProfile,
  keys: Set<string>,
  referenceKeys: Array<{ path: string; value: string }>,
  issues: PptxSceneValidationIssue[],
  elementObjects: WeakSet<object>,
): void {
  const element = requireObject(value, path, issues);
  if (!element) return;
  if (elementObjects.has(element)) {
    addIssue(
      issues,
      'invalid-scene-document',
      path,
      'Scene elements must not contain repeated or cyclic object references',
    );
    return;
  }
  elementObjects.add(element);
  registerKey(element.key, `${path}.key`, keys, issues);
  const baseKeys = [
    'authored',
    'description',
    'key',
    'name',
    'placeholder',
    'resolved',
    'title',
    'type',
  ];
  if (element.type === 'text') {
    rejectUnknownKeys(element, [...baseKeys, 'text'], path, issues);
    validateTextBody(element.text, `${path}.text`, profile, keys, issues);
  } else if (element.type === 'shape') {
    rejectUnknownKeys(element, baseKeys, path, issues);
    if (profile === 'create-text-v1') {
      addIssue(
        issues,
        'unsupported-feature',
        path,
        'Creation profile create-text-v1 supports text elements only',
      );
    }
  } else if (element.type === 'image') {
    rejectUnknownKeys(element, [...baseKeys, 'crop', 'mediaKey'], path, issues);
    if (element.crop !== undefined) {
      validateImageCrop(element.crop, `${path}.crop`, profile, issues);
    }
    if (typeof element.mediaKey === 'string') {
      referenceKeys.push({
        path: `${path}.mediaKey`,
        value: element.mediaKey,
      });
    } else if (profile === 'create-native-v1') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.mediaKey`,
        'Expected a media key',
      );
    }
    if (profile === 'create-text-v1') {
      addIssue(
        issues,
        'unsupported-feature',
        path,
        'Creation profile create-text-v1 supports text elements only',
      );
    }
  } else if (element.type === 'chart') {
    validateChartElement(element, path, baseKeys, profile, keys, issues);
  } else if (element.type === 'group') {
    rejectUnknownKeys(element, [...baseKeys, 'elements'], path, issues);
    validateElementArray(
      element.elements,
      `${path}.elements`,
      profile,
      keys,
      referenceKeys,
      issues,
      elementObjects,
    );
    if (profile === 'create-text-v1') {
      addIssue(
        issues,
        'unsupported-feature',
        path,
        'Creation profile create-text-v1 supports text elements only',
      );
    }
  } else if (element.type === 'table') {
    rejectUnknownKeys(element, [...baseKeys, 'columns', 'rows'], path, issues);
    validatePptxSceneTable(
      element,
      path,
      profile,
      keys,
      issues,
      EMUS_PER_POINT,
      TABLE_VALIDATION_DEPENDENCIES,
    );
    if (profile === 'create-text-v1') {
      addIssue(
        issues,
        'unsupported-feature',
        path,
        'Creation profile create-text-v1 supports text elements only',
      );
    }
  } else if (element.type === 'unsupported') {
    rejectUnknownKeys(
      element,
      [...baseKeys, 'feature', 'previewText'],
      path,
      issues,
    );
    if (typeof element.feature !== 'string' || element.feature.trim() === '') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.feature`,
        'Expected a non-empty unsupported feature name',
      );
    }
    optionalString(element, 'previewText', path, issues);
    if (isCreationProfile(profile)) {
      addIssue(
        issues,
        'unsupported-feature',
        path,
        profile === 'create-text-v1'
          ? 'Creation profile create-text-v1 supports text elements only'
          : 'Creation profile create-native-v1 does not support opaque elements',
      );
    }
  } else {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.type`,
      'Unknown scene element type',
    );
  }
  optionalString(element, 'description', path, issues);
  optionalString(element, 'name', path, issues);
  optionalString(element, 'title', path, issues);

  const authored = requireObject(element.authored, `${path}.authored`, issues);
  if (authored) {
    rejectUnknownKeys(
      authored,
      [
        'fillColor',
        'geometry',
        'hidden',
        'lineColor',
        'lineWidth',
        'transform',
      ],
      `${path}.authored`,
      issues,
    );
    optionalBoolean(authored, 'hidden', `${path}.authored`, issues);
    optionalColor(authored, 'fillColor', `${path}.authored`, issues);
    optionalColor(authored, 'lineColor', `${path}.authored`, issues);
    if (
      authored.geometry !== undefined &&
      !isOneOf(authored.geometry, ['ellipse', 'rect', 'roundRect'])
    ) {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.authored.geometry`,
        'Unknown text shape geometry',
      );
    }
    if (authored.lineWidth !== undefined) {
      requireFiniteNumber(
        authored.lineWidth,
        `${path}.authored.lineWidth`,
        issues,
        true,
      );
      if (isCreationProfile(profile)) {
        requireSerializableInteger(
          authored.lineWidth,
          EMUS_PER_POINT,
          `${path}.authored.lineWidth`,
          issues,
          true,
        );
      }
    }
    if (
      profile === 'create-native-v1' &&
      (element.type === 'chart' ||
        element.type === 'group' ||
        element.type === 'image' ||
        element.type === 'table') &&
      (authored.fillColor !== undefined ||
        authored.geometry !== undefined ||
        authored.lineColor !== undefined ||
        authored.lineWidth !== undefined)
    ) {
      addIssue(
        issues,
        'unsupported-feature',
        `${path}.authored`,
        `Creation profile create-native-v1 does not apply shape styling to ${element.type}s`,
      );
    }
    if (authored.transform !== undefined) {
      validateTransform(
        authored.transform,
        `${path}.authored.transform`,
        issues,
        profile,
        element.type === 'group',
      );
      if (isObject(authored.transform)) {
        validatePptxSceneTableDimensions(
          element,
          authored.transform,
          path,
          issues,
          EMUS_PER_POINT,
          TABLE_VALIDATION_DEPENDENCIES,
        );
      }
    } else if (
      isCreationProfile(profile) &&
      (element.type === 'image' ||
        element.type === 'chart' ||
        element.type === 'group' ||
        element.type === 'shape' ||
        element.type === 'table' ||
        element.type === 'text')
    ) {
      addIssue(
        issues,
        'unsupported-feature',
        `${path}.authored.transform`,
        `Creation profile ${profile} requires an authored ${element.type} transform`,
      );
    }
  }
  const resolved = requireObject(element.resolved, `${path}.resolved`, issues);
  if (resolved) {
    rejectUnknownKeys(
      resolved,
      ['hidden', 'transform'],
      `${path}.resolved`,
      issues,
    );
    if (typeof resolved.hidden !== 'boolean') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.resolved.hidden`,
        'Expected a boolean',
      );
    }
    if (resolved.transform !== undefined) {
      validateTransform(
        resolved.transform,
        `${path}.resolved.transform`,
        issues,
        profile,
        element.type === 'group',
      );
    }
  }
  if (element.placeholder !== undefined) {
    validatePlaceholder(
      element.placeholder,
      `${path}.placeholder`,
      referenceKeys,
      issues,
    );
    if (isCreationProfile(profile)) {
      addIssue(
        issues,
        'unsupported-feature',
        `${path}.placeholder`,
        `Creation profile ${profile} does not support placeholders yet`,
      );
    }
  }
}

function startsWithBytes(
  data: Uint8Array,
  expected: readonly number[],
): boolean {
  return expected.every((value, index) => data[index] === value);
}

function validateMedia(
  value: unknown,
  path: string,
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
): void {
  const media = requireObject(value, path, issues);
  if (!media) return;
  rejectUnknownKeys(media, ['data', 'key', 'mimeType'], path, issues);
  registerKey(media.key, `${path}.key`, keys, issues);
  if (!(media.data instanceof Uint8Array) || media.data.byteLength === 0) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.data`,
      'Expected non-empty Uint8Array media data',
    );
    return;
  }
  if (media.mimeType === 'image/png') {
    if (
      !startsWithBytes(
        media.data,
        [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a],
      )
    ) {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.data`,
        'PNG media data has an invalid signature',
      );
    }
  } else if (media.mimeType === 'image/jpeg') {
    if (
      !startsWithBytes(media.data, [0xff, 0xd8, 0xff]) ||
      media.data.at(-2) !== 0xff ||
      media.data.at(-1) !== 0xd9
    ) {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.data`,
        'JPEG media data has an invalid signature',
      );
    }
  } else {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.mimeType`,
      'Expected image/png or image/jpeg media type',
    );
  }
}

function validateElementArray(
  value: unknown,
  path: string,
  profile: ValidationProfile,
  keys: Set<string>,
  referenceKeys: Array<{ path: string; value: string }>,
  issues: PptxSceneValidationIssue[],
  elementObjects: WeakSet<object>,
): void {
  const elements = requireArray(value, path, issues);
  elements?.forEach((element, index) =>
    validateElement(
      element,
      `${path}[${index}]`,
      profile,
      keys,
      referenceKeys,
      issues,
      elementObjects,
    ),
  );
}

export function validatePptxScene(
  value: unknown,
  options: PptxSceneValidationOptions = {},
): PptxSceneValidationResult {
  const issues: PptxSceneValidationIssue[] = [];
  const document = requireObject(value, '$', issues);
  if (!document) return { issues, valid: false };
  rejectUnknownKeys(
    document,
    [
      'layouts',
      'masters',
      'media',
      'schemaVersion',
      'size',
      'slides',
      'themes',
    ],
    '$',
    issues,
  );
  if (document.schemaVersion !== 2) {
    addIssue(
      issues,
      'unsupported-schema-version',
      '$.schemaVersion',
      'Only PowerPoint scene schema version 2 is supported',
    );
  }
  const profile = options.profile ?? 'scene';
  validateSize(document.size, '$.size', issues, profile);
  const keys = new Set<string>();
  const references: Array<{ path: string; value: string }> = [];
  const themeKeys = new Set<string>();
  const masterKeys = new Set<string>();
  const layoutKeys = new Set<string>();
  const elementObjects = new WeakSet<object>();

  const themes = requireArray(document.themes, '$.themes', issues);
  themes?.forEach((value, index) => {
    const path = `$.themes[${index}]`;
    const theme = requireObject(value, path, issues);
    if (!theme) return;
    rejectUnknownKeys(theme, ['key', 'name'], path, issues);
    const key = registerKey(theme.key, `${path}.key`, keys, issues);
    if (key) themeKeys.add(key);
    optionalString(theme, 'name', path, issues);
  });

  const masters = requireArray(document.masters, '$.masters', issues);
  masters?.forEach((value, index) => {
    const path = `$.masters[${index}]`;
    const master = requireObject(value, path, issues);
    if (!master) return;
    rejectUnknownKeys(
      master,
      ['elements', 'key', 'name', 'themeKey'],
      path,
      issues,
    );
    const key = registerKey(master.key, `${path}.key`, keys, issues);
    if (key) masterKeys.add(key);
    optionalString(master, 'name', path, issues);
    if (typeof master.themeKey !== 'string') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.themeKey`,
        'Expected a theme key',
      );
    }
    validateElementArray(
      master.elements,
      `${path}.elements`,
      profile,
      keys,
      references,
      issues,
      elementObjects,
    );
  });

  const layouts = requireArray(document.layouts, '$.layouts', issues);
  layouts?.forEach((value, index) => {
    const path = `$.layouts[${index}]`;
    const layout = requireObject(value, path, issues);
    if (!layout) return;
    rejectUnknownKeys(
      layout,
      ['elements', 'key', 'masterKey', 'name'],
      path,
      issues,
    );
    const key = registerKey(layout.key, `${path}.key`, keys, issues);
    if (key) layoutKeys.add(key);
    optionalString(layout, 'name', path, issues);
    if (typeof layout.masterKey !== 'string') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.masterKey`,
        'Expected a master key',
      );
    }
    validateElementArray(
      layout.elements,
      `${path}.elements`,
      profile,
      keys,
      references,
      issues,
      elementObjects,
    );
  });

  const slides = requireArray(document.slides, '$.slides', issues);
  if (
    isCreationProfile(profile) &&
    slides &&
    !isSupportedPowerPointCreationSlideCount(slides.length)
  ) {
    addIssue(
      issues,
      'unsupported-feature',
      '$.slides',
      `Creation profile ${profile} supports at most ${MAX_POWERPOINT_CREATION_SLIDES} slides`,
    );
  }
  slides?.forEach((value, index) => {
    const path = `$.slides[${index}]`;
    const slide = requireObject(value, path, issues);
    if (!slide) return;
    rejectUnknownKeys(
      slide,
      ['backgroundColor', 'elements', 'hidden', 'key', 'layoutKey', 'name'],
      path,
      issues,
    );
    registerKey(slide.key, `${path}.key`, keys, issues);
    optionalColor(slide, 'backgroundColor', path, issues);
    optionalBoolean(slide, 'hidden', path, issues);
    optionalString(slide, 'name', path, issues);
    if (slide.layoutKey !== undefined && typeof slide.layoutKey !== 'string') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.layoutKey`,
        'Expected a layout key',
      );
    }
    validateElementArray(
      slide.elements,
      `${path}.elements`,
      profile,
      keys,
      references,
      issues,
      elementObjects,
    );
  });

  const media = requireArray(document.media, '$.media', issues);
  media?.forEach((value, index) =>
    validateMedia(value, `$.media[${index}]`, keys, issues),
  );
  if (profile === 'create-text-v1' && media && media.length > 0) {
    addIssue(
      issues,
      'unsupported-feature',
      '$.media',
      'Creation profile create-text-v1 does not support media resources',
    );
  }

  let hierarchyEmpty: boolean | undefined;
  if (themes && masters && layouts) {
    const emptyHierarchyCollections = [themes, masters, layouts].filter(
      (collection) => collection.length === 0,
    ).length;
    hierarchyEmpty = emptyHierarchyCollections === 3;
    if (isCreationProfile(profile) && !hierarchyEmpty) {
      addIssue(
        issues,
        'unsupported-feature',
        '$',
        `Creation profile ${profile} generates its own minimal hierarchy`,
      );
    }
    if (emptyHierarchyCollections > 0 && emptyHierarchyCollections < 3) {
      addIssue(
        issues,
        'invalid-hierarchy-reference',
        '$',
        'A declared hierarchy needs themes, masters, and layouts',
      );
    }
    masters.forEach((value, index) => {
      if (
        isObject(value) &&
        typeof value.themeKey === 'string' &&
        !themeKeys.has(value.themeKey)
      ) {
        addIssue(
          issues,
          'invalid-hierarchy-reference',
          `$.masters[${index}].themeKey`,
          'Master references an unknown theme',
        );
      }
    });
    layouts.forEach((value, index) => {
      if (
        isObject(value) &&
        typeof value.masterKey === 'string' &&
        !masterKeys.has(value.masterKey)
      ) {
        addIssue(
          issues,
          'invalid-hierarchy-reference',
          `$.layouts[${index}].masterKey`,
          'Layout references an unknown master',
        );
      }
    });
  }
  slides?.forEach((value, index) => {
    if (!isObject(value) || hierarchyEmpty === undefined) return;
    if (hierarchyEmpty) {
      if (value.layoutKey !== undefined) {
        addIssue(
          issues,
          'invalid-hierarchy-reference',
          `$.slides[${index}].layoutKey`,
          'A generated minimal hierarchy must not name a layout',
        );
      }
    } else if (
      typeof value.layoutKey !== 'string' ||
      !layoutKeys.has(value.layoutKey)
    ) {
      addIssue(
        issues,
        'invalid-hierarchy-reference',
        `$.slides[${index}].layoutKey`,
        'Slide references an unknown layout',
      );
    }
  });
  for (const reference of references) {
    if (!keys.has(reference.value)) {
      addIssue(
        issues,
        'invalid-hierarchy-reference',
        reference.path,
        'Reference points to an unknown public key',
      );
    }
  }
  if (isCreationProfile(profile) && issues.length === 0) {
    issues.push(...validatePowerPointCreationResources(document, profile));
  }

  return { issues, valid: issues.length === 0 };
}
