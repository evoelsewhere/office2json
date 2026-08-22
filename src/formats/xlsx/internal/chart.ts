import { decodeXmlEntities } from '../../../common/text/html';
import type { XmlLookupValue } from '../../../common/xml/tree';
import { getXmlNodeOrder } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type {
  XlsxChart,
  XlsxChartAxis,
  XlsxChartDataLabels,
  XlsxChartDataSource,
  XlsxChartLegend,
  XlsxChartPlot,
  XlsxChartSeries,
  XlsxChartText,
  XlsxChartType,
  XlsxDrawingColor,
} from '../types';
import { XlsxPartReader } from './part-reader';
import type { ResolvedXlsxResourceLimits } from './resource-limits';
import {
  consumeXlsxWorksheetBudget,
  consumeXlsxWorksheetFormulaCharacters,
  type XlsxWorksheetBudget,
} from './worksheet';
import type { XlsxWorkbookDiscovery } from './workbook-discovery';

type XmlRecord = Record<string, unknown>;

const CHART_NAMESPACES = {
  strict: 'http://purl.oclc.org/ooxml/drawingml/chart',
  transitional: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
} as const;

const CHART_TYPES: Readonly<Record<string, XlsxChartType>> = {
  area3DChart: 'area-3d',
  areaChart: 'area',
  bar3DChart: 'bar-3d',
  barChart: 'bar',
  bubbleChart: 'bubble',
  doughnutChart: 'doughnut',
  line3DChart: 'line-3d',
  lineChart: 'line',
  ofPieChart: 'of-pie',
  pie3DChart: 'pie-3d',
  pieChart: 'pie',
  radarChart: 'radar',
  scatterChart: 'scatter',
  stockChart: 'stock',
  surface3DChart: 'surface-3d',
  surfaceChart: 'surface',
};

const AXIS_TYPES = {
  catAx: 'category',
  dateAx: 'date',
  serAx: 'series',
  valAx: 'value',
} as const;

export type XlsxLoadedChart = Omit<
  XlsxChart,
  'description' | 'hidden' | 'id' | 'kind' | 'name' | 'part' | 'transform'
>;

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function records(value: unknown): XmlRecord[] | undefined {
  if (value === undefined) return [];
  const values: unknown[] = Array.isArray(value)
    ? [...(value as unknown[])]
    : [value];
  const output: XmlRecord[] = [];
  for (const item of values) {
    const node = record(item);
    if (!node) return undefined;
    output.push(node);
  }
  return output;
}

function optionalRecord(
  value: unknown,
  message: string,
  part: string,
): XmlRecord | undefined {
  if (value === undefined) return undefined;
  const node = record(value);
  if (!node) fail('invalid-document-structure', message, part);
  return node;
}

function attributes(node: XmlRecord): Record<string, string> {
  return (record(node.attrs) ?? {}) as Record<string, string>;
}

function localName(name: string): string {
  return name.slice(name.lastIndexOf(':') + 1);
}

function childByLocal(node: XmlRecord, name: string): unknown {
  return Object.entries(node).find(([key]) => localName(key) === name)?.[1];
}

function fail(
  code:
    | 'invalid-document-structure'
    | 'invalid-document-value'
    | 'unsupported-feature',
  message: string,
  part: string,
): never {
  throw new XlsxParseError({ code, message, part, severity: 'error' });
}

function scalar(value: unknown, message: string, part: string): string {
  const node = record(value);
  const text =
    typeof value === 'string'
      ? value
      : typeof node?.value === 'string'
        ? node.value
        : undefined;
  if (text === undefined) fail('invalid-document-structure', message, part);
  return decodeXmlEntities(text);
}

function optionalScalar(
  value: unknown,
  message: string,
  part: string,
): string | undefined {
  return value === undefined ? undefined : scalar(value, message, part);
}

function root(
  value: XmlLookupValue,
  part: string,
  dialect: 'strict' | 'transitional',
): XmlRecord {
  const entry = Object.entries(value).find(
    ([name]) => localName(name) === 'chartSpace',
  );
  const node = record(entry?.[1]);
  if (!entry || !node) {
    fail('invalid-document-structure', 'Chart root is missing', part);
  }
  const declaredChartNamespaces = Object.values(attributes(node)).filter(
    (namespace) =>
      namespace === CHART_NAMESPACES.strict ||
      namespace === CHART_NAMESPACES.transitional,
  );
  if (
    entry[0] !== 'c:chartSpace' ||
    declaredChartNamespaces.length !== 1 ||
    declaredChartNamespaces[0] !== CHART_NAMESPACES[dialect]
  ) {
    fail(
      'invalid-document-structure',
      'Chart root has the wrong namespace',
      part,
    );
  }
  return node;
}

export function parseXlsxChartUnsignedInteger(
  value: unknown,
  message: string,
  part: string,
): number {
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed)) {
    fail('invalid-document-value', message, part);
  }
  return parsed;
}

export function parseXlsxChartFiniteNumber(
  value: unknown,
  message: string,
  part: string,
): number {
  if (
    typeof value !== 'string' ||
    !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/u.test(value)
  ) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    fail('invalid-document-value', message, part);
  }
  return Object.is(parsed, -0) ? 0 : parsed;
}

const unsignedInteger = parseXlsxChartUnsignedInteger;
const finiteNumber = parseXlsxChartFiniteNumber;

function elementValue(
  node: XmlRecord,
  name: string,
  message: string,
  part: string,
): string | undefined {
  const value = childByLocal(node, name);
  if (value === undefined) return undefined;
  const element = record(value);
  if (!element) fail('invalid-document-structure', message, part);
  const result = attributes(element).val;
  if (result === undefined) fail('invalid-document-value', message, part);
  return result;
}

function booleanValue(
  node: XmlRecord,
  name: string,
  defaultValue: boolean,
  message: string,
  part: string,
): boolean {
  const value = childByLocal(node, name);
  if (value === undefined) return defaultValue;
  const element = record(value);
  if (!element) fail('invalid-document-structure', message, part);
  const authored = attributes(element).val;
  if (authored === undefined || authored === '1' || authored === 'true') {
    return true;
  }
  if (authored === '0' || authored === 'false') return false;
  fail('invalid-document-value', message, part);
}

function authoredBooleanValue(
  node: XmlRecord,
  name: string,
  message: string,
  part: string,
): boolean {
  const value = childByLocal(node, name);
  const element = record(value);
  if (!element) fail('invalid-document-structure', message, part);
  const authored = attributes(element).val;
  if (authored === undefined || authored === '1' || authored === 'true') {
    return true;
  }
  if (authored === '0' || authored === 'false') return false;
  fail('invalid-document-value', message, part);
}

function consumeText(
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  text: string,
  part: string,
): void {
  consumeXlsxWorksheetBudget(
    budget,
    'textCharacters',
    text.length,
    'maxTextCharacters',
    limits,
    part,
  );
}

function orderedChildren(
  node: XmlRecord,
): Array<{ name: string; value: unknown }> {
  return Object.entries(node)
    .flatMap(([name, value]) => {
      const children: unknown[] = Array.isArray(value)
        ? [...(value as unknown[])]
        : [value];
      return children.map((child) => ({
        name: localName(name),
        value: child,
      }));
    })
    .sort(
      (left, right) =>
        (getXmlNodeOrder(left.value) ?? Number.MAX_SAFE_INTEGER) -
        (getXmlNodeOrder(right.value) ?? Number.MAX_SAFE_INTEGER),
    );
}

function drawingColor(
  value: unknown,
  part: string,
): XlsxDrawingColor | undefined {
  const fill = optionalRecord(value, 'Chart color is invalid', part);
  if (!fill) return undefined;
  const children = { ...fill };
  delete children.attrs;
  delete children.value;
  const entries = Object.entries(children);
  if (entries.length !== 1) {
    fail('invalid-document-structure', 'Chart color is invalid', part);
  }
  const [name, colorValue] = entries[0]!;
  const color = record(colorValue);
  if (!color)
    fail('invalid-document-structure', 'Chart color is invalid', part);
  const attrs = attributes(color);
  if (
    localName(name) === 'srgbClr' &&
    typeof attrs.val === 'string' &&
    /^[0-9A-Fa-f]{6}$/u.test(attrs.val)
  ) {
    return { kind: 'rgb', value: attrs.val.toUpperCase() };
  }
  if (localName(name) === 'schemeClr' && attrs.val) {
    return { kind: 'scheme', value: attrs.val };
  }
  if (
    localName(name) === 'sysClr' &&
    attrs.val &&
    (attrs.lastClr === undefined || /^[0-9A-Fa-f]{6}$/u.test(attrs.lastClr))
  ) {
    return {
      kind: 'system',
      ...(attrs.lastClr === undefined
        ? {}
        : { lastColor: attrs.lastClr.toUpperCase() }),
      value: attrs.val,
    };
  }
  fail('invalid-document-value', 'Chart color is invalid', part);
}

function shapeColor(
  node: XmlRecord,
  part: string,
): XlsxDrawingColor | undefined {
  const properties = optionalRecord(
    childByLocal(node, 'spPr'),
    'Chart shape properties are invalid',
    part,
  );
  return properties
    ? drawingColor(childByLocal(properties, 'solidFill'), part)
    : undefined;
}

function pointCount(node: XmlRecord, part: string): number | undefined {
  const value = elementValue(
    node,
    'ptCount',
    'Chart cache point count is invalid',
    part,
  );
  return value === undefined
    ? undefined
    : unsignedInteger(value, 'Chart cache point count is invalid', part);
}

function cachePoints(
  node: XmlRecord,
  numeric: boolean,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): Array<{ index: number; value: number | string }> {
  const points = records(childByLocal(node, 'pt'));
  if (!points)
    fail('invalid-document-structure', 'Chart cache points are invalid', part);
  const output: Array<{ index: number; value: number | string }> = [];
  const indexes = new Set<number>();
  for (const point of points) {
    const index = unsignedInteger(
      attributes(point).idx,
      'Chart cache point index is invalid',
      part,
    );
    if (indexes.has(index)) {
      fail(
        'invalid-document-value',
        'Chart cache contains a duplicate point index',
        part,
      );
    }
    indexes.add(index);
    const text = scalar(
      childByLocal(point, 'v'),
      'Chart cache point value is invalid',
      part,
    );
    consumeText(budget, limits, text, part);
    output.push({
      index,
      value: numeric
        ? finiteNumber(text, 'Chart numeric cache value is invalid', part)
        : text,
    });
  }
  return output;
}

function validatePointCount(
  count: number | undefined,
  points: ReadonlyArray<{ index: number }>,
  part: string,
): void {
  if (count === undefined) return;
  if (points.some((point) => point.index >= count)) {
    fail(
      'invalid-document-value',
      'Chart cache point exceeds its declared count',
      part,
    );
  }
}

function dataSource(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxChartDataSource | undefined {
  const container = optionalRecord(value, 'Chart data source is invalid', part);
  if (!container) return undefined;
  const candidates = ['numRef', 'numLit', 'strRef', 'strLit', 'multiLvlStrRef']
    .map((name) => ({ name, node: record(childByLocal(container, name)) }))
    .filter(
      (candidate): candidate is { name: string; node: XmlRecord } =>
        candidate.node !== undefined,
    );
  if (candidates.length !== 1) {
    fail('invalid-document-structure', 'Chart data source is invalid', part);
  }
  const candidate = candidates[0]!;
  const referenced = candidate.name.endsWith('Ref');
  const formula = referenced
    ? scalar(
        childByLocal(candidate.node, 'f'),
        'Chart data source formula is invalid',
        part,
      )
    : undefined;
  if (formula !== undefined) {
    consumeXlsxWorksheetFormulaCharacters(budget, formula, limits, part);
  }
  if (candidate.name === 'multiLvlStrRef') {
    const cache = record(childByLocal(candidate.node, 'multiLvlStrCache'));
    if (!cache) {
      fail(
        'invalid-document-structure',
        'Chart multi-level cache is missing',
        part,
      );
    }
    const levels = records(childByLocal(cache, 'lvl'));
    if (!levels) {
      fail(
        'invalid-document-structure',
        'Chart multi-level cache is invalid',
        part,
      );
    }
    const count = pointCount(cache, part);
    const parsedLevels = levels.map((level) => {
      const points = cachePoints(level, false, budget, limits, part) as Array<{
        index: number;
        value: string;
      }>;
      validatePointCount(count, points, part);
      return points;
    });
    return {
      ...(formula === undefined ? {} : { formula }),
      kind: 'multi-level-string',
      levels: parsedLevels,
      ...(count === undefined ? {} : { pointCount: count }),
    };
  }
  const numeric = candidate.name.startsWith('num');
  const cache = referenced
    ? record(childByLocal(candidate.node, numeric ? 'numCache' : 'strCache'))
    : candidate.node;
  if (!cache)
    fail('invalid-document-structure', 'Chart data cache is missing', part);
  const count = pointCount(cache, part);
  const points = cachePoints(cache, numeric, budget, limits, part);
  validatePointCount(count, points, part);
  if (numeric) {
    const formatCode = optionalScalar(
      childByLocal(cache, 'formatCode'),
      'Chart cache format code is invalid',
      part,
    );
    if (formatCode !== undefined) consumeText(budget, limits, formatCode, part);
    return {
      ...(formatCode === undefined ? {} : { formatCode }),
      ...(formula === undefined ? {} : { formula }),
      kind: 'number',
      ...(count === undefined ? {} : { pointCount: count }),
      points: points as Array<{ index: number; value: number }>,
    };
  }
  return {
    ...(formula === undefined ? {} : { formula }),
    kind: 'string',
    ...(count === undefined ? {} : { pointCount: count }),
    points: points as Array<{ index: number; value: string }>,
  };
}

function richText(
  value: XmlRecord,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): string {
  const paragraphs = records(childByLocal(value, 'p'));
  if (!paragraphs)
    fail('invalid-document-structure', 'Chart text is invalid', part);
  const text = paragraphs
    .map((paragraph) => {
      const output: string[] = [];
      const visit = (node: XmlRecord): void => {
        for (const child of orderedChildren(node)) {
          if (child.name === 't') {
            output.push(scalar(child.value, 'Chart text is invalid', part));
          } else if (child.name === 'br') {
            output.push('\n');
          } else {
            const childNode = record(child.value);
            if (childNode) visit(childNode);
          }
        }
      };
      visit(paragraph);
      return output.join('');
    })
    .join('\n');
  consumeText(budget, limits, text, part);
  return text;
}

function chartText(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxChartText | undefined {
  const container = optionalRecord(value, 'Chart text is invalid', part);
  if (!container) return undefined;
  const tx = record(childByLocal(container, 'tx')) ?? container;
  const rich = record(childByLocal(tx, 'rich'));
  if (rich) return { text: richText(rich, budget, limits, part) };
  const direct = childByLocal(tx, 'v');
  if (direct !== undefined) {
    const text = scalar(direct, 'Chart text is invalid', part);
    consumeText(budget, limits, text, part);
    return { text };
  }
  const source = dataSource(tx, budget, limits, part);
  if (!source || source.kind !== 'string') {
    fail('invalid-document-structure', 'Chart text source is invalid', part);
  }
  return {
    ...(source.formula === undefined ? {} : { formula: source.formula }),
    text: source.points[0]?.value ?? '',
  };
}

function dataLabels(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxChartDataLabels | undefined {
  const node = optionalRecord(value, 'Chart data labels are invalid', part);
  if (!node) return undefined;
  const separator = optionalScalar(
    childByLocal(node, 'separator'),
    'Chart data-label separator is invalid',
    part,
  );
  if (separator !== undefined) consumeText(budget, limits, separator, part);
  const position = elementValue(
    node,
    'dLblPos',
    'Chart data-label position is invalid',
    part,
  );
  return {
    ...(position === undefined ? {} : { position }),
    ...(separator === undefined ? {} : { separator }),
    showBubbleSize: booleanValue(
      node,
      'showBubbleSize',
      false,
      'Chart data-label bubble-size flag is invalid',
      part,
    ),
    showCategoryName: booleanValue(
      node,
      'showCatName',
      false,
      'Chart data-label category-name flag is invalid',
      part,
    ),
    showLegendKey: booleanValue(
      node,
      'showLegendKey',
      false,
      'Chart data-label legend-key flag is invalid',
      part,
    ),
    showPercent: booleanValue(
      node,
      'showPercent',
      false,
      'Chart data-label percent flag is invalid',
      part,
    ),
    showSeriesName: booleanValue(
      node,
      'showSerName',
      false,
      'Chart data-label series-name flag is invalid',
      part,
    ),
    showValue: booleanValue(
      node,
      'showVal',
      false,
      'Chart data-label value flag is invalid',
      part,
    ),
  };
}

function series(
  node: XmlRecord,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxChartSeries {
  const indexValue = elementValue(
    node,
    'idx',
    'Chart series index is invalid',
    part,
  );
  const orderValue = elementValue(
    node,
    'order',
    'Chart series order is invalid',
    part,
  );
  if (indexValue === undefined || orderValue === undefined) {
    fail(
      'invalid-document-structure',
      'Chart series identity is missing',
      part,
    );
  }
  const markerNode = optionalRecord(
    childByLocal(node, 'marker'),
    'Chart marker is invalid',
    part,
  );
  const markerSize = markerNode
    ? elementValue(markerNode, 'size', 'Chart marker size is invalid', part)
    : undefined;
  const markerSymbol = markerNode
    ? elementValue(markerNode, 'symbol', 'Chart marker symbol is invalid', part)
    : undefined;
  const labels = dataLabels(childByLocal(node, 'dLbls'), budget, limits, part);
  const color = shapeColor(node, part);
  const bubbleSizes = dataSource(
    childByLocal(node, 'bubbleSize'),
    budget,
    limits,
    part,
  );
  const categories = dataSource(
    childByLocal(node, 'cat'),
    budget,
    limits,
    part,
  );
  const name = chartText(childByLocal(node, 'tx'), budget, limits, part);
  const values = dataSource(childByLocal(node, 'val'), budget, limits, part);
  const xValues = dataSource(childByLocal(node, 'xVal'), budget, limits, part);
  const yValues = dataSource(childByLocal(node, 'yVal'), budget, limits, part);
  const smooth =
    childByLocal(node, 'smooth') === undefined
      ? undefined
      : authoredBooleanValue(
          node,
          'smooth',
          'Chart series smooth flag is invalid',
          part,
        );
  return {
    ...(bubbleSizes === undefined ? {} : { bubbleSizes }),
    ...(categories === undefined ? {} : { categories }),
    ...(color === undefined ? {} : { color }),
    ...(labels === undefined ? {} : { dataLabels: labels }),
    index: unsignedInteger(indexValue, 'Chart series index is invalid', part),
    ...(markerNode === undefined
      ? {}
      : {
          marker: {
            ...(markerSize === undefined
              ? {}
              : {
                  size: unsignedInteger(
                    markerSize,
                    'Chart marker size is invalid',
                    part,
                  ),
                }),
            ...(markerSymbol === undefined ? {} : { symbol: markerSymbol }),
          },
        }),
    ...(name === undefined ? {} : { name }),
    order: unsignedInteger(orderValue, 'Chart series order is invalid', part),
    ...(smooth === undefined ? {} : { smooth }),
    ...(values === undefined ? {} : { values }),
    ...(xValues === undefined ? {} : { xValues }),
    ...(yValues === undefined ? {} : { yValues }),
  };
}

function optionalNumberElement(
  node: XmlRecord,
  name: string,
  message: string,
  part: string,
): number | undefined {
  const value = elementValue(node, name, message, part);
  return value === undefined ? undefined : finiteNumber(value, message, part);
}

function plot(
  type: XlsxChartType,
  node: XmlRecord,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxChartPlot {
  const seriesNodes = records(childByLocal(node, 'ser'));
  const axisNodes = records(childByLocal(node, 'axId'));
  if (!seriesNodes || !axisNodes) {
    fail('invalid-document-structure', 'Chart plot structure is invalid', part);
  }
  const parsedSeries = seriesNodes.map((item) =>
    series(item, budget, limits, part),
  );
  const indexes = new Set<number>();
  const orders = new Set<number>();
  for (const item of parsedSeries) {
    if (indexes.has(item.index) || orders.has(item.order)) {
      fail(
        'invalid-document-value',
        'Chart plot contains duplicate series identity',
        part,
      );
    }
    indexes.add(item.index);
    orders.add(item.order);
  }
  const direction = elementValue(
    node,
    'barDir',
    'Chart bar direction is invalid',
    part,
  );
  if (direction !== undefined && direction !== 'bar' && direction !== 'col') {
    fail('invalid-document-value', 'Chart bar direction is invalid', part);
  }
  const labels = dataLabels(childByLocal(node, 'dLbls'), budget, limits, part);
  const bubbleScale = optionalNumberElement(
    node,
    'bubbleScale',
    'Chart bubble scale is invalid',
    part,
  );
  const firstSliceAngle = optionalNumberElement(
    node,
    'firstSliceAng',
    'Chart first-slice angle is invalid',
    part,
  );
  const gapDepth = optionalNumberElement(
    node,
    'gapDepth',
    'Chart gap depth is invalid',
    part,
  );
  const gapWidth = optionalNumberElement(
    node,
    'gapWidth',
    'Chart gap width is invalid',
    part,
  );
  const holeSize = optionalNumberElement(
    node,
    'holeSize',
    'Chart hole size is invalid',
    part,
  );
  const overlap = optionalNumberElement(
    node,
    'overlap',
    'Chart overlap is invalid',
    part,
  );
  const grouping = elementValue(
    node,
    'grouping',
    'Chart grouping is invalid',
    part,
  );
  const radarStyle = elementValue(
    node,
    'radarStyle',
    'Chart radar style is invalid',
    part,
  );
  const scatterStyle = elementValue(
    node,
    'scatterStyle',
    'Chart scatter style is invalid',
    part,
  );
  return {
    axisIds: axisNodes.map((axis) =>
      unsignedInteger(
        attributes(axis).val,
        'Chart plot axis reference is invalid',
        part,
      ),
    ),
    ...(direction === undefined
      ? {}
      : { barDirection: direction === 'bar' ? 'bar' : 'column' }),
    ...(bubbleScale === undefined ? {} : { bubbleScale }),
    ...(labels === undefined ? {} : { dataLabels: labels }),
    ...(firstSliceAngle === undefined ? {} : { firstSliceAngle }),
    ...(gapDepth === undefined ? {} : { gapDepth }),
    ...(gapWidth === undefined ? {} : { gapWidth }),
    ...(grouping === undefined ? {} : { grouping }),
    ...(holeSize === undefined ? {} : { holeSize }),
    ...(overlap === undefined ? {} : { overlap }),
    ...(radarStyle === undefined ? {} : { radarStyle }),
    ...(scatterStyle === undefined ? {} : { scatterStyle }),
    series: parsedSeries,
    type,
    varyColors: booleanValue(
      node,
      'varyColors',
      false,
      'Chart vary-colors flag is invalid',
      part,
    ),
  };
}

function axis(
  kind: XlsxChartAxis['kind'],
  node: XmlRecord,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxChartAxis {
  const id = elementValue(node, 'axId', 'Chart axis ID is invalid', part);
  if (id === undefined)
    fail('invalid-document-structure', 'Chart axis ID is missing', part);
  const scaling = record(childByLocal(node, 'scaling'));
  if (!scaling)
    fail('invalid-document-structure', 'Chart axis scaling is missing', part);
  const orientation = elementValue(
    scaling,
    'orientation',
    'Chart axis orientation is invalid',
    part,
  );
  if (
    orientation !== undefined &&
    orientation !== 'minMax' &&
    orientation !== 'maxMin'
  ) {
    fail('invalid-document-value', 'Chart axis orientation is invalid', part);
  }
  const position = elementValue(
    node,
    'axPos',
    'Chart axis position is invalid',
    part,
  );
  const positions = { b: 'bottom', l: 'left', r: 'right', t: 'top' } as const;
  if (position !== undefined && !(position in positions)) {
    fail('invalid-document-value', 'Chart axis position is invalid', part);
  }
  const numberFormat = optionalRecord(
    childByLocal(node, 'numFmt'),
    'Chart axis number format is invalid',
    part,
  );
  const numberFormatAttrs = numberFormat ? attributes(numberFormat) : undefined;
  const formatCode = numberFormatAttrs?.formatCode;
  if (numberFormatAttrs && !formatCode) {
    fail('invalid-document-value', 'Chart axis number format is invalid', part);
  }
  if (formatCode !== undefined) consumeText(budget, limits, formatCode, part);
  let sourceLinked = true;
  if (
    numberFormatAttrs?.sourceLinked === '0' ||
    numberFormatAttrs?.sourceLinked === 'false'
  ) {
    sourceLinked = false;
  } else if (
    numberFormatAttrs?.sourceLinked !== undefined &&
    numberFormatAttrs.sourceLinked !== '1' &&
    numberFormatAttrs.sourceLinked !== 'true'
  ) {
    fail(
      'invalid-document-value',
      'Chart axis number-format source flag is invalid',
      part,
    );
  }
  const crossAxisValue = elementValue(
    node,
    'crossAx',
    'Chart cross-axis reference is invalid',
    part,
  );
  const crosses = elementValue(
    node,
    'crosses',
    'Chart axis crossing is invalid',
    part,
  );
  const crossesAt = optionalNumberElement(
    node,
    'crossesAt',
    'Chart axis crossing value is invalid',
    part,
  );
  const logBase = optionalNumberElement(
    scaling,
    'logBase',
    'Chart axis logarithm base is invalid',
    part,
  );
  const majorUnit = optionalNumberElement(
    node,
    'majorUnit',
    'Chart axis major unit is invalid',
    part,
  );
  const maximum = optionalNumberElement(
    scaling,
    'max',
    'Chart axis maximum is invalid',
    part,
  );
  const minimum = optionalNumberElement(
    scaling,
    'min',
    'Chart axis minimum is invalid',
    part,
  );
  const minorUnit = optionalNumberElement(
    node,
    'minorUnit',
    'Chart axis minor unit is invalid',
    part,
  );
  const title = chartText(childByLocal(node, 'title'), budget, limits, part);
  return {
    ...(crossAxisValue === undefined
      ? {}
      : {
          crossAxis: unsignedInteger(
            crossAxisValue,
            'Chart cross-axis reference is invalid',
            part,
          ),
        }),
    ...(crosses === undefined ? {} : { crosses }),
    ...(crossesAt === undefined ? {} : { crossesAt }),
    deleted: booleanValue(
      node,
      'delete',
      false,
      'Chart axis delete flag is invalid',
      part,
    ),
    id: unsignedInteger(id, 'Chart axis ID is invalid', part),
    kind,
    ...(logBase === undefined ? {} : { logBase }),
    majorGridlines: childByLocal(node, 'majorGridlines') !== undefined,
    ...(majorUnit === undefined ? {} : { majorUnit }),
    ...(maximum === undefined ? {} : { maximum }),
    ...(minimum === undefined ? {} : { minimum }),
    minorGridlines: childByLocal(node, 'minorGridlines') !== undefined,
    ...(minorUnit === undefined ? {} : { minorUnit }),
    ...(formatCode === undefined
      ? {}
      : {
          numberFormat: {
            code: formatCode,
            sourceLinked,
          },
        }),
    orientation: orientation === 'maxMin' ? 'max-min' : 'min-max',
    ...(position === undefined
      ? {}
      : { position: positions[position as keyof typeof positions] }),
    ...(title === undefined ? {} : { title }),
  };
}

function legend(value: unknown, part: string): XlsxChartLegend | undefined {
  const node = optionalRecord(value, 'Chart legend is invalid', part);
  if (!node) return undefined;
  const position = elementValue(
    node,
    'legendPos',
    'Chart legend position is invalid',
    part,
  );
  const positions = {
    b: 'bottom',
    l: 'left',
    r: 'right',
    t: 'top',
    tr: 'top-right',
  } as const;
  if (position !== undefined && !(position in positions)) {
    fail('invalid-document-value', 'Chart legend position is invalid', part);
  }
  const entryNodes = records(childByLocal(node, 'legendEntry'));
  if (!entryNodes)
    fail(
      'invalid-document-structure',
      'Chart legend entries are invalid',
      part,
    );
  const entries = entryNodes.map((entry) => {
    const index = elementValue(
      entry,
      'idx',
      'Chart legend entry index is invalid',
      part,
    );
    if (index === undefined) {
      fail(
        'invalid-document-structure',
        'Chart legend entry index is missing',
        part,
      );
    }
    return {
      deleted: booleanValue(
        entry,
        'delete',
        false,
        'Chart legend entry delete flag is invalid',
        part,
      ),
      index: unsignedInteger(
        index,
        'Chart legend entry index is invalid',
        part,
      ),
    };
  });
  return {
    entries,
    overlay: booleanValue(
      node,
      'overlay',
      false,
      'Chart legend overlay flag is invalid',
      part,
    ),
    ...(position === undefined
      ? {}
      : { position: positions[position as keyof typeof positions] }),
  };
}

function parseChart(
  value: XmlLookupValue,
  discovery: XlsxWorkbookDiscovery,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxLoadedChart {
  const chartSpace = root(value, part, discovery.dialect);
  const chart = record(childByLocal(chartSpace, 'chart'));
  if (!chart)
    fail('invalid-document-structure', 'Chart definition is missing', part);
  const plotArea = record(childByLocal(chart, 'plotArea'));
  if (!plotArea)
    fail('invalid-document-structure', 'Chart plot area is missing', part);
  const plots: XlsxChartPlot[] = [];
  const axes: XlsxChartAxis[] = [];
  for (const child of orderedChildren(plotArea)) {
    const chartType = CHART_TYPES[child.name];
    const axisType = AXIS_TYPES[child.name as keyof typeof AXIS_TYPES];
    if (!chartType && !axisType) {
      if (child.name.endsWith('Chart')) {
        fail('unsupported-feature', 'Chart family is not supported', part);
      }
      continue;
    }
    const childNode = record(child.value);
    if (!childNode) {
      fail('invalid-document-structure', 'Chart plot area is invalid', part);
    }
    if (chartType) {
      plots.push(plot(chartType, childNode, budget, limits, part));
      continue;
    }
    axes.push(axis(axisType, childNode, budget, limits, part));
  }
  if (plots.length === 0) {
    fail(
      'invalid-document-structure',
      'Chart contains no supported plot',
      part,
    );
  }
  const axisIds = new Set<number>();
  for (const item of axes) {
    if (axisIds.has(item.id)) {
      fail(
        'invalid-document-value',
        'Chart contains a duplicate axis ID',
        part,
      );
    }
    axisIds.add(item.id);
  }
  for (const item of plots) {
    if (item.axisIds.some((id) => !axisIds.has(id))) {
      fail(
        'invalid-document-value',
        'Chart plot references a missing axis',
        part,
      );
    }
  }
  for (const item of axes) {
    if (item.crossAxis !== undefined && !axisIds.has(item.crossAxis)) {
      fail(
        'invalid-document-value',
        'Chart axis references a missing cross axis',
        part,
      );
    }
  }
  const displayBlanks = elementValue(
    chart,
    'dispBlanksAs',
    'Chart blank-display mode is invalid',
    part,
  );
  if (
    displayBlanks !== undefined &&
    displayBlanks !== 'gap' &&
    displayBlanks !== 'span' &&
    displayBlanks !== 'zero'
  ) {
    fail('invalid-document-value', 'Chart blank-display mode is invalid', part);
  }
  const styleValue = elementValue(
    chartSpace,
    'style',
    'Chart style is invalid',
    part,
  );
  const chartLegend = legend(childByLocal(chart, 'legend'), part);
  const title = chartText(childByLocal(chart, 'title'), budget, limits, part);
  return {
    axes,
    autoTitleDeleted: booleanValue(
      chart,
      'autoTitleDeleted',
      false,
      'Chart auto-title-delete flag is invalid',
      part,
    ),
    displayBlanksAs: displayBlanks ?? 'gap',
    ...(chartLegend === undefined ? {} : { legend: chartLegend }),
    plots,
    plotVisibleOnly: booleanValue(
      chart,
      'plotVisOnly',
      true,
      'Chart visible-only flag is invalid',
      part,
    ),
    roundedCorners: booleanValue(
      chartSpace,
      'roundedCorners',
      false,
      'Chart rounded-corners flag is invalid',
      part,
    ),
    showDataLabelsOverMaximum: booleanValue(
      chart,
      'showDLblsOverMax',
      false,
      'Chart labels-over-maximum flag is invalid',
      part,
    ),
    ...(styleValue === undefined
      ? {}
      : { style: unsignedInteger(styleValue, 'Chart style is invalid', part) }),
    ...(title === undefined ? {} : { title }),
  };
}

export async function loadXlsxChart(
  part: string,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
): Promise<XlsxLoadedChart> {
  const value = await reader.readXml(part, { required: true });
  return parseChart(value, discovery, budget, limits, part);
}
