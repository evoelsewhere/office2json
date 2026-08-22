import { XlsxParseError } from '../errors';
import type {
  XlsxProtectedRange,
  XlsxRange,
  XlsxWorksheetProtection,
} from '../types';
import { parseXlsxRangeReference } from './cell-reference';
import { parseXlsxProtectionCredential } from './protection-hash';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import type { XlsxResolvedSheetSelection } from './selection';
import type { XlsxXmlElement, XlsxXmlEventSink } from './streaming-xml';
import {
  consumeXlsxWorksheetBudget,
  type XlsxWorksheetBudget,
} from './worksheet';

type XmlRecord = Record<string, unknown>;

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

function records(value: unknown): XmlRecord[] | undefined {
  const values = Array.isArray(value) ? value : [value];
  const output: XmlRecord[] = [];
  for (const item of values) {
    const parsed = record(item);
    if (!parsed) return undefined;
    output.push(parsed);
  }
  return output;
}

function elementAttributes(element: XlsxXmlElement): XmlRecord {
  return Object.fromEntries(
    [...element.attributes].flatMap(([name, value]) =>
      name.startsWith('{}') ? [[name.slice(2), value]] : [],
    ),
  );
}

function attributes(node: XmlRecord): XmlRecord {
  return record(node.attrs) ?? {};
}

function optionalBoolean(
  value: unknown,
  part: string,
  message: string,
): boolean {
  if (value === undefined || value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  valueFailure(message, part);
}

function consumeText(
  amount: number,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): void {
  consumeXlsxWorksheetBudget(
    budget,
    'textCharacters',
    amount,
    'maxTextCharacters',
    limits,
    part,
  );
}

export function parseXlsxWorksheetProtection(
  element: XlsxXmlElement,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxWorksheetProtection {
  const attrs = elementAttributes(element);
  const credential = parseXlsxProtectionCredential(
    attrs,
    {
      algorithmName: 'algorithmName',
      hashValue: 'hashValue',
      legacyHash: 'password',
      saltValue: 'saltValue',
      spinCount: 'spinCount',
    },
    part,
    'Worksheet protection',
  );
  consumeText(credential.textCharacters, budget, limits, part);
  const flag = (name: string, message: string) =>
    optionalBoolean(attrs[name], part, message);
  return {
    ...(credential.credential === undefined
      ? {}
      : { credential: credential.credential }),
    protectAutoFilter: flag(
      'autoFilter',
      'Worksheet protection auto-filter flag is invalid',
    ),
    protectDeleteColumns: flag(
      'deleteColumns',
      'Worksheet protection delete-columns flag is invalid',
    ),
    protectDeleteRows: flag(
      'deleteRows',
      'Worksheet protection delete-rows flag is invalid',
    ),
    protectFormatCells: flag(
      'formatCells',
      'Worksheet protection format-cells flag is invalid',
    ),
    protectFormatColumns: flag(
      'formatColumns',
      'Worksheet protection format-columns flag is invalid',
    ),
    protectFormatRows: flag(
      'formatRows',
      'Worksheet protection format-rows flag is invalid',
    ),
    protectInsertColumns: flag(
      'insertColumns',
      'Worksheet protection insert-columns flag is invalid',
    ),
    protectInsertHyperlinks: flag(
      'insertHyperlinks',
      'Worksheet protection insert-hyperlinks flag is invalid',
    ),
    protectInsertRows: flag(
      'insertRows',
      'Worksheet protection insert-rows flag is invalid',
    ),
    protectObjects: flag(
      'objects',
      'Worksheet protection objects flag is invalid',
    ),
    protectPivotTables: flag(
      'pivotTables',
      'Worksheet protection pivot-tables flag is invalid',
    ),
    protectScenarios: flag(
      'scenarios',
      'Worksheet protection scenarios flag is invalid',
    ),
    protectSelectLockedCells: flag(
      'selectLockedCells',
      'Worksheet protection select-locked-cells flag is invalid',
    ),
    protectSelectUnlockedCells: flag(
      'selectUnlockedCells',
      'Worksheet protection select-unlocked-cells flag is invalid',
    ),
    protectSheet: flag('sheet', 'Worksheet protection sheet flag is invalid'),
    protectSort: flag('sort', 'Worksheet protection sort flag is invalid'),
  };
}

function protectedRanges(
  value: unknown,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxRange[] {
  if (typeof value !== 'string') {
    valueFailure('Protected-range reference list is invalid', part);
  }
  const references = value.trim().split(/[\t\n\r ]+/u);
  if (references[0] === '') {
    valueFailure('Protected-range reference list is invalid', part);
  }
  const ranges = references.map((reference) => {
    const parsed = parseXlsxRangeReference(reference);
    if (!parsed || reference.includes('$')) {
      valueFailure('Protected-range reference is invalid', part);
    }
    if (parsed.end.row > limits.maxRowsPerWorksheet) {
      throw new XlsxResourceLimitError(
        'maxRowsPerWorksheet',
        parsed.end.row,
        limits.maxRowsPerWorksheet,
        part,
      );
    }
    if (parsed.end.column > limits.maxColumnsPerWorksheet) {
      throw new XlsxResourceLimitError(
        'maxColumnsPerWorksheet',
        parsed.end.column,
        limits.maxColumnsPerWorksheet,
        part,
      );
    }
    consumeXlsxWorksheetBudget(
      budget,
      'rangeAreas',
      1,
      'maxRangeAreas',
      limits,
      part,
    );
    return parsed;
  });
  if (new Set(ranges.map((range) => range.reference)).size !== ranges.length) {
    valueFailure('Protected-range reference list contains duplicates', part);
  }
  return ranges;
}

function intersects(left: XlsxRange, right: XlsxRange): boolean {
  return (
    left.start.row <= right.end.row &&
    left.end.row >= right.start.row &&
    left.start.column <= right.end.column &&
    left.end.column >= right.start.column
  );
}

function selectionRelation(
  selection: XlsxResolvedSheetSelection,
  ranges: readonly XlsxRange[],
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxProtectedRange['selectionRelation'] | null {
  if (selection.kind !== 'selected-ranges') {
    return selection.kind === 'full-sheet' ? 'full-sheet' : null;
  }
  for (const range of ranges) {
    for (const selected of selection.ranges) {
      consumeXlsxWorksheetBudget(
        budget,
        'scannedCells',
        1,
        'maxScannedCells',
        limits,
        part,
      );
      if (intersects(range, selected)) return 'intersects-selection';
    }
  }
  return null;
}

export function parseXlsxProtectedRanges(
  value: unknown,
  selection: XlsxResolvedSheetSelection,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxProtectedRange[] {
  const root = record(value);
  if (!root || Array.isArray(value)) {
    structureFailure('Protected-ranges collection is invalid', part);
  }
  const nodes = records(root.protectedRange);
  if (!nodes || nodes.length === 0) {
    structureFailure('Protected-range collection is empty or invalid', part);
  }
  return nodes.flatMap((node) => {
    const attrs = attributes(node);
    const name = typeof attrs.name === 'string' ? attrs.name : undefined;
    if (name === undefined || name.length === 0) {
      valueFailure('Protected-range name is invalid', part);
    }
    const securityDescriptor =
      typeof attrs.securityDescriptor === 'string'
        ? attrs.securityDescriptor
        : undefined;
    const credential = parseXlsxProtectionCredential(
      attrs,
      {
        algorithmName: 'algorithmName',
        hashValue: 'hashValue',
        legacyHash: 'password',
        saltValue: 'saltValue',
        spinCount: 'spinCount',
      },
      part,
      'Protected range',
    );
    consumeText(
      name.length +
        (securityDescriptor?.length ?? 0) +
        credential.textCharacters,
      budget,
      limits,
      part,
    );
    const ranges = protectedRanges(attrs.sqref, budget, limits, part);
    const relation = selectionRelation(selection, ranges, budget, limits, part);
    if (relation === null) return [];
    return [
      {
        ...(credential.credential === undefined
          ? {}
          : { credential: credential.credential }),
        name,
        ranges,
        ...(securityDescriptor === undefined ? {} : { securityDescriptor }),
        selectionRelation: relation,
      },
    ];
  });
}

interface CapturedNode {
  name: string;
  node: XmlRecord;
}

function capturedAttributes(element: XlsxXmlElement): XmlRecord {
  return Object.fromEntries(
    [...element.attributes].map(([name, value]) => [
      name.startsWith('{}') ? name.slice(2) : name,
      value,
    ]),
  );
}

export class XlsxProtectedRangesCapture implements XlsxXmlEventSink {
  private readonly stack: CapturedNode[] = [];
  private root: XmlRecord | undefined;

  constructor(
    private readonly selection: XlsxResolvedSheetSelection,
    private readonly budget: XlsxWorksheetBudget,
    private readonly limits: ResolvedXlsxResourceLimits,
    private readonly part: string,
  ) {}

  openElement(element: XlsxXmlElement): void {
    const node: XmlRecord = { attrs: capturedAttributes(element) };
    const parent = this.stack.at(-1)?.node;
    if (parent) {
      const current = parent[element.localName];
      if (current === undefined) parent[element.localName] = node;
      else if (Array.isArray(current)) current.push(node);
      else parent[element.localName] = [current, node];
    } else {
      if (element.localName !== 'protectedRanges' || this.root !== undefined) {
        structureFailure('Protected-ranges capture root is invalid', this.part);
      }
      this.root = node;
    }
    this.stack.push({ name: element.localName, node });
  }

  closeElement(element: XlsxXmlElement): void {
    if (this.stack.pop()?.name !== element.localName) {
      structureFailure(
        'Protected-ranges capture nesting is invalid',
        this.part,
      );
    }
  }

  text(value: string): void {
    if (value.trim().length !== 0) {
      structureFailure('Protected-ranges text content is invalid', this.part);
    }
  }

  result(): XlsxProtectedRange[] {
    if (!this.root || this.stack.length !== 0) {
      structureFailure('Protected-ranges capture is incomplete', this.part);
    }
    return parseXlsxProtectedRanges(
      this.root,
      this.selection,
      this.budget,
      this.limits,
      this.part,
    );
  }
}
