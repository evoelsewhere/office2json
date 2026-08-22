import { XlsxParseError } from '../errors';
import type {
  XlsxHeaderFooter,
  XlsxPageBreak,
  XlsxPageMargins,
  XlsxPageSetup,
  XlsxPageSetupProperties,
  XlsxPrintOptions,
  XlsxUniversalMeasure,
} from '../types';
import {
  type ResolvedXlsxResourceLimits,
  XLSX_MAX_COLUMNS,
  XLSX_MAX_ROWS,
  XlsxResourceLimitError,
} from './resource-limits';
import type { XlsxXmlElement, XlsxXmlEventSink } from './streaming-xml';
import {
  consumeXlsxWorksheetBudget,
  type XlsxWorksheetBudget,
} from './worksheet';

type XmlAttributes = Record<string, string>;
type BreakAxis = 'column' | 'row';

const DECIMAL_PATTERN = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/u;
const UNSIGNED_INTEGER_PATTERN = /^(?:0|[1-9]\d*)$/u;
const UNIVERSAL_MEASURE_PATTERN =
  /^(\+?(?:\d+(?:\.\d*)?|\.\d+))(cm|in|mm|pc|pi|pt)$/u;
const HEADER_FOOTER_FIELDS = [
  'evenFooter',
  'evenHeader',
  'firstFooter',
  'firstHeader',
  'oddFooter',
  'oddHeader',
] as const;

type HeaderFooterField = (typeof HEADER_FOOTER_FIELDS)[number];

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

function attributes(element: XlsxXmlElement): XmlAttributes {
  return Object.fromEntries(
    [...element.attributes].flatMap(([name, value]) =>
      name.startsWith('{}') ? [[name.slice(2), value]] : [],
    ),
  );
}

function booleanValue(
  value: string | undefined,
  fallback: boolean,
  message: string,
  part: string,
): boolean {
  if (value === undefined) return fallback;
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  valueFailure(message, part);
}

function unsignedInteger(
  value: string | undefined,
  message: string,
  part: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (!UNSIGNED_INTEGER_PATTERN.test(value)) valueFailure(message, part);
  const parsed = Number(value);
  if (parsed > 0xffff_ffff) {
    valueFailure(message, part);
  }
  return parsed;
}

function finiteNumber(value: string, message: string, part: string): number {
  if (!DECIMAL_PATTERN.test(value)) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < 0) valueFailure(message, part);
  return Object.is(parsed, -0) ? 0 : parsed;
}

function universalMeasure(
  value: string | undefined,
  message: string,
  part: string,
): XlsxUniversalMeasure | undefined {
  if (value === undefined) return undefined;
  const match = UNIVERSAL_MEASURE_PATTERN.exec(value);
  if (!match) valueFailure(message, part);
  const amount = Number(match[1]);
  if (!Number.isFinite(amount) || amount <= 0) valueFailure(message, part);
  return {
    unit: match[2] as XlsxUniversalMeasure['unit'],
    value: amount,
  };
}

function enumValue<const T extends string>(
  value: string | undefined,
  values: Readonly<Record<string, T>>,
  fallback: T,
  message: string,
  part: string,
): T {
  if (value === undefined) return fallback;
  const parsed = values[value];
  if (parsed === undefined) valueFailure(message, part);
  return parsed;
}

export function parseXlsxPrintOptions(
  element: XlsxXmlElement,
  part: string,
): XlsxPrintOptions {
  const attrs = attributes(element);
  const flag = (name: string, message: string) =>
    booleanValue(attrs[name], false, message, part);
  return {
    gridLines: flag('gridLines', 'Worksheet print grid-lines flag is invalid'),
    gridLinesSet: flag(
      'gridLinesSet',
      'Worksheet print grid-lines-set flag is invalid',
    ),
    headings: flag('headings', 'Worksheet print headings flag is invalid'),
    horizontalCentered: flag(
      'horizontalCentered',
      'Worksheet print horizontal-centering flag is invalid',
    ),
    verticalCentered: flag(
      'verticalCentered',
      'Worksheet print vertical-centering flag is invalid',
    ),
  };
}

export function parseXlsxPageMargins(
  element: XlsxXmlElement,
  part: string,
): XlsxPageMargins {
  const attrs = attributes(element);
  const margin = (name: string) => {
    const value = attrs[name];
    if (value === undefined) {
      valueFailure(`Worksheet page ${name} margin is missing`, part);
    }
    return finiteNumber(
      value,
      `Worksheet page ${name} margin is invalid`,
      part,
    );
  };
  return {
    bottom: margin('bottom'),
    footer: margin('footer'),
    header: margin('header'),
    left: margin('left'),
    right: margin('right'),
    top: margin('top'),
  };
}

export function parseXlsxPageSetupProperties(
  element: XlsxXmlElement,
  part: string,
): XlsxPageSetupProperties {
  const attrs = attributes(element);
  return {
    autoPageBreaks: booleanValue(
      attrs.autoPageBreaks,
      true,
      'Worksheet automatic page-break flag is invalid',
      part,
    ),
    fitToPage: booleanValue(
      attrs.fitToPage,
      false,
      'Worksheet fit-to-page flag is invalid',
      part,
    ),
  };
}

export function parseXlsxPageSetup(
  element: XlsxXmlElement,
  part: string,
): XlsxPageSetup {
  const attrs = attributes(element);
  const integer = (name: string, message: string) =>
    unsignedInteger(attrs[name], message, part);
  const scale = integer('scale', 'Worksheet page scale is invalid');
  if (scale !== undefined && (scale < 10 || scale > 400)) {
    valueFailure('Worksheet page scale is invalid', part);
  }
  const paperSize = integer('paperSize', 'Worksheet paper size is invalid');
  if (paperSize === 0) valueFailure('Worksheet paper size is invalid', part);
  const copies = integer('copies', 'Worksheet print copies are invalid');
  if (copies === 0) valueFailure('Worksheet print copies are invalid', part);
  return {
    blackAndWhite: booleanValue(
      attrs.blackAndWhite,
      false,
      'Worksheet black-and-white print flag is invalid',
      part,
    ),
    cellComments: enumValue(
      attrs.cellComments,
      { asDisplayed: 'as-displayed', atEnd: 'at-end', none: 'none' },
      'none',
      'Worksheet printed-comment mode is invalid',
      part,
    ),
    ...(copies === undefined ? {} : { copies }),
    draft: booleanValue(
      attrs.draft,
      false,
      'Worksheet draft print flag is invalid',
      part,
    ),
    errors: enumValue(
      attrs.errors,
      {
        blank: 'blank',
        dash: 'dash',
        displayed: 'displayed',
        NA: 'not-available',
      },
      'displayed',
      'Worksheet printed-error mode is invalid',
      part,
    ),
    ...optionalInteger(
      'firstPageNumber',
      integer('firstPageNumber', 'Worksheet first page number is invalid'),
    ),
    ...optionalInteger(
      'fitToHeight',
      integer('fitToHeight', 'Worksheet fit-to-height value is invalid'),
    ),
    ...optionalInteger(
      'fitToWidth',
      integer('fitToWidth', 'Worksheet fit-to-width value is invalid'),
    ),
    ...optionalInteger(
      'horizontalDpi',
      integer('horizontalDpi', 'Worksheet horizontal DPI is invalid'),
    ),
    orientation: enumValue(
      attrs.orientation,
      { default: 'default', landscape: 'landscape', portrait: 'portrait' },
      'default',
      'Worksheet page orientation is invalid',
      part,
    ),
    pageOrder: enumValue(
      attrs.pageOrder,
      { downThenOver: 'down-then-over', overThenDown: 'over-then-down' },
      'down-then-over',
      'Worksheet page order is invalid',
      part,
    ),
    ...optionalMeasure(
      'paperHeight',
      universalMeasure(
        attrs.paperHeight,
        'Worksheet paper height is invalid',
        part,
      ),
    ),
    ...(paperSize === undefined ? {} : { paperSize }),
    ...optionalMeasure(
      'paperWidth',
      universalMeasure(
        attrs.paperWidth,
        'Worksheet paper width is invalid',
        part,
      ),
    ),
    ...(scale === undefined ? {} : { scale }),
    useFirstPageNumber: booleanValue(
      attrs.useFirstPageNumber,
      false,
      'Worksheet use-first-page-number flag is invalid',
      part,
    ),
    usePrinterDefaults: booleanValue(
      attrs.usePrinterDefaults,
      false,
      'Worksheet use-printer-defaults flag is invalid',
      part,
    ),
    ...optionalInteger(
      'verticalDpi',
      integer('verticalDpi', 'Worksheet vertical DPI is invalid'),
    ),
  };
}

function optionalInteger(
  name: string,
  value: number | undefined,
): Record<string, number> {
  return value === undefined ? {} : { [name]: value };
}

function optionalMeasure(
  name: string,
  value: XlsxUniversalMeasure | undefined,
): Record<string, XlsxUniversalMeasure> {
  return value === undefined ? {} : { [name]: value };
}

export class XlsxHeaderFooterCapture implements XlsxXmlEventSink {
  private attributes: XmlAttributes | undefined;
  private current: HeaderFooterField | undefined;
  private closed = false;
  private readonly seen = new Set<HeaderFooterField>();
  private readonly values: Partial<Record<HeaderFooterField, string>> = {};

  constructor(
    private readonly budget: XlsxWorksheetBudget,
    private readonly limits: ResolvedXlsxResourceLimits,
    private readonly part: string,
  ) {}

  openElement(element: XlsxXmlElement): void {
    if (this.attributes === undefined) {
      if (element.localName !== 'headerFooter') {
        structureFailure(
          'Worksheet header/footer capture root is invalid',
          this.part,
        );
      }
      this.attributes = attributes(element);
      return;
    }
    if (
      this.closed ||
      this.current !== undefined ||
      !HEADER_FOOTER_FIELDS.includes(element.localName as HeaderFooterField)
    ) {
      structureFailure('Worksheet header/footer nesting is invalid', this.part);
    }
    const field = element.localName as HeaderFooterField;
    if (this.seen.has(field)) {
      structureFailure(
        'Worksheet header/footer contains duplicate fields',
        this.part,
      );
    }
    this.seen.add(field);
    this.current = field;
    this.values[field] = '';
  }

  closeElement(element: XlsxXmlElement): void {
    if (element.localName === 'headerFooter') {
      if (this.current !== undefined || this.closed) {
        structureFailure(
          'Worksheet header/footer nesting is invalid',
          this.part,
        );
      }
      this.closed = true;
      return;
    }
    if (this.current !== element.localName) {
      structureFailure('Worksheet header/footer nesting is invalid', this.part);
    }
    this.current = undefined;
  }

  text(value: string): void {
    if (this.current === undefined) {
      if (value.trim().length !== 0) {
        structureFailure('Worksheet header/footer text is invalid', this.part);
      }
      return;
    }
    consumeXlsxWorksheetBudget(
      this.budget,
      'textCharacters',
      value.length,
      'maxTextCharacters',
      this.limits,
      this.part,
    );
    this.values[this.current] += value;
  }

  result(): XlsxHeaderFooter {
    if (this.attributes === undefined || !this.closed) {
      structureFailure(
        'Worksheet header/footer capture is incomplete',
        this.part,
      );
    }
    const attrs = this.attributes;
    return {
      alignWithMargins: booleanValue(
        attrs.alignWithMargins,
        true,
        'Worksheet header/footer margin-alignment flag is invalid',
        this.part,
      ),
      differentFirst: booleanValue(
        attrs.differentFirst,
        false,
        'Worksheet first-header/footer flag is invalid',
        this.part,
      ),
      differentOddEven: booleanValue(
        attrs.differentOddEven,
        false,
        'Worksheet odd/even-header/footer flag is invalid',
        this.part,
      ),
      ...this.values,
      scaleWithDocument: booleanValue(
        attrs.scaleWithDoc,
        true,
        'Worksheet header/footer scale-with-document flag is invalid',
        this.part,
      ),
    };
  }
}

export class XlsxPageBreaksCapture implements XlsxXmlEventSink {
  private readonly breaks: XlsxPageBreak[] = [];
  private childOpen = false;
  private closed = false;
  private expectedCount: number | undefined;
  private expectedManualCount: number | undefined;
  private opened = false;
  private readonly positions = new Set<number>();

  constructor(
    private readonly axis: BreakAxis,
    private readonly budget: XlsxWorksheetBudget,
    private readonly limits: ResolvedXlsxResourceLimits,
    private readonly part: string,
  ) {}

  openElement(element: XlsxXmlElement): void {
    if (!this.opened) {
      const expectedRoot = this.axis === 'row' ? 'rowBreaks' : 'colBreaks';
      if (element.localName !== expectedRoot) {
        structureFailure(
          'Worksheet page-break capture root is invalid',
          this.part,
        );
      }
      this.opened = true;
      const attrs = attributes(element);
      this.expectedCount = unsignedInteger(
        attrs.count,
        'Worksheet page-break count is invalid',
        this.part,
      );
      this.expectedManualCount = unsignedInteger(
        attrs.manualBreakCount,
        'Worksheet manual page-break count is invalid',
        this.part,
      );
      if (
        this.expectedCount !== undefined &&
        this.expectedCount > this.limits.maxRangeAreas
      ) {
        throw new XlsxResourceLimitError(
          'maxRangeAreas',
          this.expectedCount,
          this.limits.maxRangeAreas,
          this.part,
        );
      }
      return;
    }
    if (this.closed || this.childOpen || element.localName !== 'brk') {
      structureFailure('Worksheet page-break nesting is invalid', this.part);
    }
    const attrs = attributes(element);
    const position = unsignedInteger(
      attrs.id,
      'Worksheet page-break position is invalid',
      this.part,
    );
    if (position === undefined) {
      valueFailure('Worksheet page-break position is invalid', this.part);
    }
    const positionLimit =
      this.axis === 'row'
        ? this.limits.maxRowsPerWorksheet
        : this.limits.maxColumnsPerWorksheet;
    if (position > positionLimit) {
      throw new XlsxResourceLimitError(
        this.axis === 'row' ? 'maxRowsPerWorksheet' : 'maxColumnsPerWorksheet',
        position,
        positionLimit,
        this.part,
      );
    }
    if (this.positions.has(position)) {
      valueFailure(
        'Worksheet contains duplicate page-break positions',
        this.part,
      );
    }
    this.positions.add(position);
    const extentLimit =
      this.axis === 'row'
        ? this.limits.maxColumnsPerWorksheet
        : this.limits.maxRowsPerWorksheet;
    const start =
      unsignedInteger(
        attrs.min,
        'Worksheet page-break start is invalid',
        this.part,
      ) ?? 0;
    const end =
      unsignedInteger(
        attrs.max,
        'Worksheet page-break end is invalid',
        this.part,
      ) ?? (this.axis === 'row' ? XLSX_MAX_COLUMNS : XLSX_MAX_ROWS) - 1;
    if (end < start) {
      valueFailure('Worksheet page-break extent is invalid', this.part);
    }
    if (end >= extentLimit) {
      throw new XlsxResourceLimitError(
        this.axis === 'row' ? 'maxColumnsPerWorksheet' : 'maxRowsPerWorksheet',
        Math.max(start, end) + 1,
        extentLimit,
        this.part,
      );
    }
    consumeXlsxWorksheetBudget(
      this.budget,
      'rangeAreas',
      1,
      'maxRangeAreas',
      this.limits,
      this.part,
    );
    this.breaks.push({
      end,
      manual: booleanValue(
        attrs.man,
        false,
        'Worksheet manual page-break flag is invalid',
        this.part,
      ),
      pivot: booleanValue(
        attrs.pt,
        false,
        'Worksheet pivot page-break flag is invalid',
        this.part,
      ),
      position,
      start,
    });
    this.childOpen = true;
  }

  closeElement(element: XlsxXmlElement): void {
    const root = this.axis === 'row' ? 'rowBreaks' : 'colBreaks';
    if (element.localName === 'brk') {
      if (!this.childOpen || this.closed) {
        structureFailure('Worksheet page-break nesting is invalid', this.part);
      }
      this.childOpen = false;
      return;
    }
    if (element.localName !== root || this.childOpen || this.closed) {
      structureFailure('Worksheet page-break nesting is invalid', this.part);
    }
    this.closed = true;
  }

  text(value: string): void {
    if (value.trim().length !== 0) {
      structureFailure('Worksheet page-break text is invalid', this.part);
    }
  }

  result(): XlsxPageBreak[] {
    if (!this.opened || !this.closed) {
      structureFailure('Worksheet page-break capture is incomplete', this.part);
    }
    if (
      this.expectedCount !== undefined &&
      this.expectedCount !== this.breaks.length
    ) {
      structureFailure('Worksheet page-break count does not match', this.part);
    }
    const manualCount = this.breaks.filter((entry) => entry.manual).length;
    if (
      this.expectedManualCount !== undefined &&
      this.expectedManualCount !== manualCount
    ) {
      structureFailure(
        'Worksheet manual page-break count does not match',
        this.part,
      );
    }
    return this.breaks;
  }
}
