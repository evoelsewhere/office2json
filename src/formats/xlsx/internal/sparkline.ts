import { XlsxParseError } from '../errors';
import type {
  XlsxColor,
  XlsxSparkline,
  XlsxSparklineColors,
  XlsxSparklineGroup,
} from '../types';
import { parseXlsxRangeReference } from './cell-reference';
import type { ResolvedXlsxResourceLimits } from './resource-limits';
import type { XlsxResolvedSheetSelection } from './selection';
import { xlsxSelectionIncludesCell } from './selection';
import { parseXlsxStyleColor } from './style-color';
import type { XlsxXmlElement, XlsxXmlEventSink } from './streaming-xml';
import {
  consumeXlsxWorksheetBudget,
  consumeXlsxWorksheetFormulaCharacters,
  type XlsxWorksheetBudget,
} from './worksheet';

interface ExtensionNode {
  attributes: Record<string, string>;
  children: ExtensionNode[];
  localName: string;
  namespace: string;
  text: string;
}

function x14Namespace(): string {
  return 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
}

function xmNamespace(): string {
  return 'http://schemas.microsoft.com/office/excel/2006/main';
}

function sparklineExtensionUri(): string {
  return '{05c60535-1f16-4fd2-b633-f4f36f0b64e0}';
}

function fail(
  code:
    | 'invalid-cell-reference'
    | 'invalid-document-structure'
    | 'invalid-document-value',
  message: string,
  part: string,
): never {
  throw new XlsxParseError({ code, message, part, severity: 'error' });
}

function attribute(node: ExtensionNode, name: string): string | undefined {
  return node.attributes[`{}${name}`];
}

function children(
  node: ExtensionNode,
  namespace: string,
  localName: string,
): ExtensionNode[] {
  return node.children.filter(
    (child) => child.namespace === namespace && child.localName === localName,
  );
}

function child(
  node: ExtensionNode,
  namespace: string,
  localName: string,
  part: string,
): ExtensionNode | undefined {
  const values = children(node, namespace, localName);
  if (values.length > 1) {
    fail(
      'invalid-document-structure',
      `Sparkline ${localName} element is duplicated`,
      part,
    );
  }
  return values[0];
}

function booleanAttribute(
  value: string | undefined,
  message: string,
  part: string,
): boolean {
  if (value === undefined || value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  fail('invalid-document-value', message, part);
}

function finiteNumber(
  value: string | undefined,
  message: string,
  part: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/u.test(value)) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) fail('invalid-document-value', message, part);
  return Object.is(parsed, -0) ? 0 : parsed;
}

function enumValue<T extends string>(
  value: string | undefined,
  fallback: T,
  values: readonly T[],
  message: string,
  part: string,
): T {
  if (value === undefined) return fallback;
  const parsed = values.find((candidate) => candidate === value);
  if (!parsed) fail('invalid-document-value', message, part);
  return parsed;
}

function color(
  group: ExtensionNode,
  localName: string,
  part: string,
): XlsxColor | undefined {
  const node = child(group, x14Namespace(), localName, part);
  if (!node) return undefined;
  return parseXlsxStyleColor(
    {
      attrs: Object.fromEntries(
        Object.entries(node.attributes).map(([name, value]) => [
          name.slice(2),
          value,
        ]),
      ),
    },
    part,
    `Sparkline ${localName}`,
  );
}

function sparklineColors(
  group: ExtensionNode,
  part: string,
): XlsxSparklineColors {
  const mappings = {
    axis: 'colorAxis',
    first: 'colorFirst',
    high: 'colorHigh',
    last: 'colorLast',
    low: 'colorLow',
    markers: 'colorMarkers',
    negative: 'colorNegative',
    series: 'colorSeries',
  } as const;
  const output: XlsxSparklineColors = {};
  for (const [key, element] of Object.entries(mappings) as Array<
    [keyof XlsxSparklineColors, string]
  >) {
    const value = color(group, element, part);
    if (value !== undefined) output[key] = value;
  }
  return output;
}

function parseSparkline(
  node: ExtensionNode,
  selection: XlsxResolvedSheetSelection,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxSparkline | null {
  const formulas = children(node, xmNamespace(), 'f');
  const locations = children(node, xmNamespace(), 'sqref');
  if (formulas.length !== 1 || locations.length !== 1) {
    fail(
      'invalid-document-structure',
      'Sparkline formula and location are required exactly once',
      part,
    );
  }
  const dataFormula = formulas[0]!.text;
  if (dataFormula.length === 0 || dataFormula !== dataFormula.trim()) {
    fail('invalid-document-value', 'Sparkline data formula is invalid', part);
  }
  consumeXlsxWorksheetFormulaCharacters(budget, dataFormula, limits, part);
  const location = parseXlsxRangeReference(locations[0]!.text);
  if (
    !location ||
    location.start.row !== location.end.row ||
    location.start.column !== location.end.column
  ) {
    fail('invalid-cell-reference', 'Sparkline location is invalid', part);
  }
  consumeXlsxWorksheetBudget(
    budget,
    'rangeAreas',
    1,
    'maxRangeAreas',
    limits,
    part,
  );
  const selected = xlsxSelectionIncludesCell(
    selection,
    location.start.row,
    location.start.column,
  );
  if (!selected) return null;
  return {
    dataFormula,
    location: location.reference,
    selectionRelation:
      selection.kind === 'full-sheet' ? 'full-sheet' : 'intersects-selection',
  };
}

function parseGroup(
  node: ExtensionNode,
  selection: XlsxResolvedSheetSelection,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxSparklineGroup | null {
  const sparklineContainers = children(node, x14Namespace(), 'sparklines');
  if (sparklineContainers.length !== 1) {
    fail(
      'invalid-document-structure',
      'Sparkline group must contain one sparkline collection',
      part,
    );
  }
  const sparklineNodes = children(
    sparklineContainers[0]!,
    x14Namespace(),
    'sparkline',
  );
  if (sparklineNodes.length === 0) {
    fail('invalid-document-structure', 'Sparkline group is empty', part);
  }
  const sparklines = sparklineNodes.flatMap((sparkline) => {
    const parsed = parseSparkline(sparkline, selection, budget, limits, part);
    return parsed ? [parsed] : [];
  });
  const manualMaximum = finiteNumber(
    attribute(node, 'manualMax'),
    'Sparkline manual maximum is invalid',
    part,
  );
  const manualMinimum = finiteNumber(
    attribute(node, 'manualMin'),
    'Sparkline manual minimum is invalid',
    part,
  );
  const lineWeight = finiteNumber(
    attribute(node, 'lineWeight'),
    'Sparkline line weight is invalid',
    part,
  );
  if (lineWeight !== undefined && lineWeight < 0) {
    fail('invalid-document-value', 'Sparkline line weight is invalid', part);
  }
  const minimumAxisType = enumValue(
    attribute(node, 'minAxisType'),
    'individual',
    ['individual', 'group', 'custom'],
    'Sparkline minimum-axis type is invalid',
    part,
  );
  const maximumAxisType = enumValue(
    attribute(node, 'maxAxisType'),
    'individual',
    ['individual', 'group', 'custom'],
    'Sparkline maximum-axis type is invalid',
    part,
  );
  if (minimumAxisType === 'custom' && manualMinimum === undefined) {
    fail(
      'invalid-document-value',
      'Sparkline custom minimum requires a manual value',
      part,
    );
  }
  if (maximumAxisType === 'custom' && manualMaximum === undefined) {
    fail(
      'invalid-document-value',
      'Sparkline custom maximum requires a manual value',
      part,
    );
  }
  if (sparklines.length === 0) return null;
  return {
    colors: sparklineColors(node, part),
    dateAxis: booleanAttribute(
      attribute(node, 'dateAxis'),
      'Sparkline date-axis flag is invalid',
      part,
    ),
    displayEmptyCellsAs: enumValue(
      attribute(node, 'displayEmptyCellsAs'),
      'zero',
      ['zero', 'span', 'gap'],
      'Sparkline empty-cell display mode is invalid',
      part,
    ),
    displayHidden: booleanAttribute(
      attribute(node, 'displayHidden'),
      'Sparkline hidden-data flag is invalid',
      part,
    ),
    displayXAxis: booleanAttribute(
      attribute(node, 'displayXAxis'),
      'Sparkline X-axis flag is invalid',
      part,
    ),
    first: booleanAttribute(
      attribute(node, 'first'),
      'Sparkline first-point flag is invalid',
      part,
    ),
    high: booleanAttribute(
      attribute(node, 'high'),
      'Sparkline high-point flag is invalid',
      part,
    ),
    last: booleanAttribute(
      attribute(node, 'last'),
      'Sparkline last-point flag is invalid',
      part,
    ),
    ...(lineWeight === undefined ? {} : { lineWeight }),
    low: booleanAttribute(
      attribute(node, 'low'),
      'Sparkline low-point flag is invalid',
      part,
    ),
    ...(manualMaximum === undefined ? {} : { manualMaximum }),
    ...(manualMinimum === undefined ? {} : { manualMinimum }),
    markers: booleanAttribute(
      attribute(node, 'markers'),
      'Sparkline marker flag is invalid',
      part,
    ),
    maximumAxisType,
    minimumAxisType,
    negative: booleanAttribute(
      attribute(node, 'negative'),
      'Sparkline negative-point flag is invalid',
      part,
    ),
    rightToLeft: booleanAttribute(
      attribute(node, 'rightToLeft'),
      'Sparkline right-to-left flag is invalid',
      part,
    ),
    sparklines,
    type: enumValue(
      attribute(node, 'type'),
      'line',
      ['line', 'column', 'stacked'],
      'Sparkline type is invalid',
      part,
    ),
  };
}

export class XlsxWorksheetExtensionsCapture implements XlsxXmlEventSink {
  private root: ExtensionNode | undefined;
  private readonly stack: ExtensionNode[] = [];
  private unsupportedExtensionSeen = false;

  constructor(
    private readonly worksheetNamespace: string,
    private readonly selection: XlsxResolvedSheetSelection,
    private readonly budget: XlsxWorksheetBudget,
    private readonly limits: ResolvedXlsxResourceLimits,
    private readonly part: string,
  ) {}

  openElement(element: XlsxXmlElement): void {
    const node: ExtensionNode = {
      attributes: Object.fromEntries(element.attributes),
      children: [],
      localName: element.localName,
      namespace: element.namespace,
      text: '',
    };
    const parent = this.stack.at(-1);
    if (parent) parent.children.push(node);
    else {
      if (
        this.root ||
        element.localName !== 'extLst' ||
        element.namespace !== this.worksheetNamespace
      ) {
        fail(
          'invalid-document-structure',
          'Worksheet extension capture root is invalid',
          this.part,
        );
      }
      this.root = node;
    }
    this.stack.push(node);
  }

  closeElement(element: XlsxXmlElement): void {
    const node = this.stack.pop();
    if (
      !node ||
      node.localName !== element.localName ||
      node.namespace !== element.namespace
    ) {
      fail(
        'invalid-document-structure',
        'Worksheet extension capture nesting is invalid',
        this.part,
      );
    }
  }

  text(value: string): void {
    const node = this.stack.at(-1);
    if (!node) {
      fail(
        'invalid-document-structure',
        'Worksheet extension text is outside the root',
        this.part,
      );
    }
    node.text += value;
  }

  result(): XlsxSparklineGroup[] {
    if (!this.root || this.stack.length !== 0) {
      fail(
        'invalid-document-structure',
        'Worksheet extension capture is incomplete',
        this.part,
      );
    }
    const groups: XlsxSparklineGroup[] = [];
    let sparklineExtensionSeen = false;
    for (const extension of this.root.children) {
      if (
        extension.localName !== 'ext' ||
        extension.namespace !== this.worksheetNamespace
      ) {
        fail(
          'invalid-document-structure',
          'Worksheet extension entry is invalid',
          this.part,
        );
      }
      const uri = attribute(extension, 'uri');
      if (!uri || uri !== uri.trim()) {
        fail(
          'invalid-document-value',
          'Worksheet extension URI is invalid',
          this.part,
        );
      }
      if (uri.toLowerCase() !== sparklineExtensionUri()) {
        this.unsupportedExtensionSeen = true;
        continue;
      }
      if (sparklineExtensionSeen) {
        fail(
          'invalid-document-structure',
          'Worksheet contains duplicate sparkline extensions',
          this.part,
        );
      }
      sparklineExtensionSeen = true;
      const containers = children(extension, x14Namespace(), 'sparklineGroups');
      if (containers.length !== 1) {
        fail(
          'invalid-document-structure',
          'Sparkline extension payload is invalid',
          this.part,
        );
      }
      for (const group of children(
        containers[0]!,
        x14Namespace(),
        'sparklineGroup',
      )) {
        const parsed = parseGroup(
          group,
          this.selection,
          this.budget,
          this.limits,
          this.part,
        );
        if (parsed) groups.push(parsed);
      }
    }
    return groups;
  }

  hasUnsupportedExtension(): boolean {
    return this.unsupportedExtensionSeen;
  }
}
