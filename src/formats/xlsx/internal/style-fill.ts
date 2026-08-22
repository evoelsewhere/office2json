import { XlsxParseError } from '../errors';
import type {
  XlsxColor,
  XlsxFill,
  XlsxGradientStop,
  XlsxPatternType,
} from '../types';
import { parseXlsxStyleColor } from './style-color';

type XmlRecord = Record<string, unknown>;

function patternTypes(): readonly XlsxPatternType[] {
  return [
    'none',
    'solid',
    'mediumGray',
    'darkGray',
    'lightGray',
    'darkHorizontal',
    'darkVertical',
    'darkDown',
    'darkUp',
    'darkGrid',
    'darkTrellis',
    'lightHorizontal',
    'lightVertical',
    'lightDown',
    'lightUp',
    'lightGrid',
    'lightTrellis',
    'gray125',
    'gray0625',
  ];
}

function structureFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-structure',
    message,
    part,
    severity: 'error',
  });
}

function valueFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function child(
  root: XmlRecord,
  prefix: string,
  localName: string,
  context: string,
  part: string,
): XmlRecord | undefined {
  const value = root[prefix ? `${prefix}:${localName}` : localName];
  if (value === undefined) return undefined;
  const values: unknown[] = Array.isArray(value) ? value : [value];
  if (values.length !== 1) {
    structureFailure(`${context} ${localName} element is duplicated`, part);
  }
  const item = record(values[0]);
  if (!item) {
    structureFailure(`${context} ${localName} element is invalid`, part);
  }
  return item;
}

function attributes(node: XmlRecord, context: string, part: string): XmlRecord {
  if (node.attrs === undefined) return {};
  const attrs = record(node.attrs);
  if (!attrs) structureFailure(`${context} attributes are invalid`, part);
  return attrs;
}

function decimal(
  value: unknown,
  minimum: number,
  maximum: number,
  message: string,
  part: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (
    typeof value !== 'string' ||
    !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/u.test(value)
  ) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < minimum || parsed > maximum) {
    valueFailure(message, part);
  }
  return parsed;
}

function normalizedOptional(value: number | undefined): number | undefined {
  return value === 0 ? undefined : value;
}

function patternFill(node: XmlRecord, prefix: string, part: string): XlsxFill {
  const attrs = attributes(node, 'Pattern fill', part);
  const sourcePattern = attrs.patternType;
  const pattern = sourcePattern === undefined ? 'none' : sourcePattern;
  if (!patternTypes().includes(pattern as XlsxPatternType)) {
    valueFailure('Pattern fill type is invalid', part);
  }
  const foregroundColor = parseXlsxStyleColor(
    child(node, prefix, 'fgColor', 'Pattern fill', part),
    part,
    'Pattern foreground',
  );
  const backgroundColor = parseXlsxStyleColor(
    child(node, prefix, 'bgColor', 'Pattern fill', part),
    part,
    'Pattern background',
  );
  return Object.freeze({
    ...(backgroundColor === undefined ? {} : { backgroundColor }),
    ...(foregroundColor === undefined ? {} : { foregroundColor }),
    kind: 'pattern',
    pattern: pattern as XlsxPatternType,
  });
}

function stopValues(
  node: XmlRecord,
  prefix: string,
  part: string,
): XmlRecord[] {
  const value = node[prefix ? `${prefix}:stop` : 'stop'];
  if (!Array.isArray(value) || value.length < 2) {
    structureFailure('Gradient fill requires at least two stops', part);
  }
  return value.map((item) => {
    const parsed = record(item);
    if (!parsed) structureFailure('Gradient stop element is invalid', part);
    return parsed;
  });
}

function gradientColor(
  node: XmlRecord,
  prefix: string,
  part: string,
): XlsxColor {
  const color = parseXlsxStyleColor(
    child(node, prefix, 'color', 'Gradient stop', part),
    part,
    'Gradient stop',
  );
  if (!color) valueFailure('Gradient stop color is missing', part);
  return color;
}

function gradientStops(
  node: XmlRecord,
  prefix: string,
  part: string,
): XlsxGradientStop[] {
  const stops: XlsxGradientStop[] = [];
  let previous = -1;
  for (const stop of stopValues(node, prefix, part)) {
    const position = decimal(
      attributes(stop, 'Gradient stop', part).position,
      0,
      1,
      'Gradient stop position is invalid',
      part,
    );
    if (position === undefined) {
      valueFailure('Gradient stop position is missing', part);
    }
    if (position <= previous) {
      valueFailure('Gradient stop positions are out of order', part);
    }
    previous = position;
    stops.push(
      Object.freeze({
        color: gradientColor(stop, prefix, part),
        position,
      }),
    );
  }
  return Object.freeze(stops) as XlsxGradientStop[];
}

function gradientFill(node: XmlRecord, prefix: string, part: string): XlsxFill {
  const attrs = attributes(node, 'Gradient fill', part);
  const sourceType = attrs.type;
  const type = sourceType === undefined ? 'linear' : sourceType;
  if (type !== 'linear' && type !== 'path') {
    valueFailure('Gradient fill type is invalid', part);
  }
  const degree = decimal(
    attrs.degree,
    0,
    360,
    'Gradient fill degree is invalid',
    part,
  );
  const left = decimal(
    attrs.left,
    0,
    1,
    'Gradient fill left bound is invalid',
    part,
  );
  const right = decimal(
    attrs.right,
    0,
    1,
    'Gradient fill right bound is invalid',
    part,
  );
  const top = decimal(
    attrs.top,
    0,
    1,
    'Gradient fill top bound is invalid',
    part,
  );
  const bottom = decimal(
    attrs.bottom,
    0,
    1,
    'Gradient fill bottom bound is invalid',
    part,
  );
  if (
    type === 'linear' &&
    (left !== undefined ||
      right !== undefined ||
      top !== undefined ||
      bottom !== undefined)
  ) {
    valueFailure('Linear gradient fill has path bounds', part);
  }
  if (type === 'path' && degree !== undefined) {
    valueFailure('Path gradient fill has a degree', part);
  }
  const angle = normalizedOptional(degree);
  const normalizedLeft = normalizedOptional(left);
  const normalizedRight = normalizedOptional(right);
  const normalizedTop = normalizedOptional(top);
  const normalizedBottom = normalizedOptional(bottom);
  return Object.freeze({
    ...(angle === undefined ? {} : { angle }),
    ...(normalizedBottom === undefined ? {} : { bottom: normalizedBottom }),
    kind: 'gradient',
    ...(normalizedLeft === undefined ? {} : { left: normalizedLeft }),
    ...(normalizedRight === undefined ? {} : { right: normalizedRight }),
    stops: gradientStops(node, prefix, part),
    ...(normalizedTop === undefined ? {} : { top: normalizedTop }),
    type,
  });
}

export function parseXlsxStyleFill(
  value: unknown,
  prefix: string,
  part: string,
): XlsxFill {
  const root = record(value);
  if (!root) structureFailure('Fill element is invalid', part);
  const pattern = child(root, prefix, 'patternFill', 'Fill', part);
  const gradient = child(root, prefix, 'gradientFill', 'Fill', part);
  if (pattern && gradient) {
    structureFailure('Fill has both pattern and gradient definitions', part);
  }
  if (gradient) return gradientFill(gradient, prefix, part);
  return patternFill(pattern ?? {}, prefix, part);
}
