import { XlsxParseError } from '../errors';
import type {
  XlsxCell,
  XlsxCellMetadata,
  XlsxCellValue,
  XlsxAutoFilter,
  XlsxColumnRange,
  XlsxConditionalFormatting,
  XlsxDataValidation,
  XlsxDataValidationSettings,
  XlsxFormula,
  XlsxHyperlink,
  XlsxProtectedRange,
  XlsxHeaderFooter,
  XlsxPageBreak,
  XlsxPageMargins,
  XlsxPageSetup,
  XlsxPageSetupProperties,
  XlsxPrintOptions,
  XlsxRange,
  XlsxRichTextRun,
  XlsxRow,
  XlsxSparklineGroup,
  XlsxWorksheetView,
  XlsxWorksheetFormat,
  XlsxWorksheetOutline,
  XlsxWorksheetProtection,
  XlsxWorksheetPrintSettings,
  XlsxColor,
} from '../types';
import { XlsxAutoFilterCapture } from './auto-filter';
import {
  parseXlsxCellReference,
  parseXlsxRangeReference,
  xlsxColumnName,
} from './cell-reference';
import {
  parseXlsxScalarCellValue,
  type XlsxScalarCellType,
} from './cell-value';
import { normalizeXlsxSerialDate } from './date-system';
import { XlsxDataValidationsCapture } from './data-validation';
import { XlsxConditionalFormattingCapture } from './conditional-format';
import { translateXlsxSharedFormula } from './formula';
import { parseXlsxHyperlink, type ParsedXlsxHyperlink } from './hyperlink';
import {
  parseXlsxWorksheetProtection,
  XlsxProtectedRangesCapture,
} from './worksheet-protection';
import {
  parseXlsxPageMargins,
  parseXlsxPageSetup,
  parseXlsxPageSetupProperties,
  parseXlsxPrintOptions,
  XlsxHeaderFooterCapture,
  XlsxPageBreaksCapture,
} from './worksheet-print';
import { xlsxNumberFormatDatePrecision } from './number-format';
import { XlsxPartReader } from './part-reader';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import {
  type XlsxResolvedSheetSelection,
  xlsxSelectionIncludesCell,
  xlsxSelectionIncludesRow,
} from './selection';
import type { XlsxSharedStringTable } from './shared-strings';
import { XlsxWorksheetExtensionsCapture } from './sparkline';
import {
  EMPTY_XLSX_CELL_METADATA,
  resolveXlsxCellMetadata,
  type XlsxCellMetadataBudget,
  type XlsxCellMetadataRegistry,
} from './cell-metadata';
import type { XlsxRelationship } from './relationships';
import { EMPTY_XLSX_STYLE_TABLE, type XlsxStyleTable } from './styles';
import type { XlsxXmlElement, XlsxXmlEventSink } from './streaming-xml';
import {
  type XlsxWorkbookDiscovery,
  XLSX_SPREADSHEET_NAMESPACES,
} from './workbook-discovery';
import {
  normalizeXlsxColumnRanges,
  type XlsxAuthoredColumnRange,
  xlsxMergedRangesOverlap,
} from './worksheet-layout';
import {
  parseXlsxDeclaredDimension,
  parseXlsxWorksheetFormat,
  parseXlsxWorksheetOutline,
  parseXlsxWorksheetTabColor,
} from './worksheet-metadata';
import {
  parseXlsxWorksheetPane,
  parseXlsxWorksheetView,
  parseXlsxWorksheetViewSelection,
  validateXlsxWorksheetView,
} from './worksheet-view';

const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const UNSIGNED_INTEGER_PATTERN = /^(?:0|[1-9]\d*)$/u;
const NONNEGATIVE_DECIMAL_PATTERN = /^(?:\d+(?:\.\d*)?|\.\d+)$/u;
const TABLE_RELATIONSHIP_NAMESPACE = {
  strict: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
  transitional:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
} as const;

type CellType = XlsxScalarCellType | 'inlineStr';
type TextCapture = 'formula' | 'inline' | 'value' | null;

interface PendingFormula {
  reference?: string;
  sharedIndex?: string;
  text: string;
  type: 'array' | 'dataTable' | 'normal' | 'shared';
}

interface PendingCell {
  address: string;
  column: number;
  formula?: PendingFormula;
  hasInlineString: boolean;
  hasValue: boolean;
  inlineMode: 'plain' | 'rich' | 'unset';
  inlineRuns: XlsxRichTextRun[];
  inlineText: string;
  metadata?: XlsxCellMetadata;
  numberFormat?: string;
  selected: boolean;
  style?: number;
  type: CellType;
  valueText: string;
}

interface PendingInlineRun {
  hasText: boolean;
  text: string;
}

export interface XlsxWorksheetBudget {
  conditionalFormattingRules: number;
  formulaCharacters: number;
  formulaGroups: number;
  rangeAreas: number;
  returnedCells: number;
  richTextRuns: number;
  scannedCells: number;
  textCharacters: number;
  validationRules: number;
}

interface SharedFormulaMaster {
  column: number;
  expression: string;
  range: XlsxRange;
  row: number;
}

export interface XlsxWorksheetPayload {
  autoFilter?: XlsxAutoFilter;
  columns: XlsxColumnRange[];
  conditionalFormattings: XlsxConditionalFormatting[];
  declaredDimension?: XlsxRange;
  dataValidationSettings?: XlsxDataValidationSettings;
  dataValidations: XlsxDataValidation[];
  hyperlinks: XlsxHyperlink[];
  mergedRanges: XlsxRange[];
  outline?: XlsxWorksheetOutline;
  print?: XlsxWorksheetPrintSettings;
  protectedRanges: XlsxProtectedRange[];
  protection?: XlsxWorksheetProtection;
  rows: XlsxRow[];
  sheetFormat?: XlsxWorksheetFormat;
  sparklineGroups?: XlsxSparklineGroup[];
  tabColor?: XlsxColor;
  unsupportedExtensions?: true;
  unsupportedMetadata?: true;
  views: XlsxWorksheetView[];
}

export interface XlsxWorksheetSemantics {
  dateSystem: '1900' | '1904';
  dialect: 'strict' | 'transitional';
  drawingRelationshipIds?: string[];
  relationships: ReadonlyMap<string, XlsxRelationship>;
  legacyDrawingRelationshipIds?: string[];
  metadataBudget?: XlsxCellMetadataBudget;
  metadataRegistry?: XlsxCellMetadataRegistry;
  pivotTableRelationshipIds?: string[];
  styles: XlsxStyleTable;
  tableRelationshipIds?: string[];
  workbookViewCount: number;
}

const DEFAULT_WORKSHEET_SEMANTICS: XlsxWorksheetSemantics = Object.freeze({
  dateSystem: '1900',
  dialect: 'transitional',
  relationships: new Map(),
  styles: EMPTY_XLSX_STYLE_TABLE,
  workbookViewCount: 1,
});

function structureFailure(
  part: string,
  cell: string | undefined,
  message: string,
): never {
  throw new XlsxParseError({
    ...(cell === undefined ? {} : { cell }),
    code: 'invalid-document-structure',
    message,
    part,
    severity: 'error',
  });
}

function valueFailure(
  part: string,
  cell: string | undefined,
  message: string,
): never {
  throw new XlsxParseError({
    ...(cell === undefined ? {} : { cell }),
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function formulaFailure(part: string, cell: string, message: string): never {
  throw new XlsxParseError({
    cell,
    code: 'invalid-formula',
    message,
    part,
    severity: 'error',
  });
}

function attribute(
  element: XlsxXmlElement,
  localName: string,
): string | undefined {
  return element.attributes.get(`{}${localName}`);
}

function unsignedInteger(
  value: string | undefined,
  part: string,
  cell: string | undefined,
  message: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (!UNSIGNED_INTEGER_PATTERN.test(value)) valueFailure(part, cell, message);
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > 0xffff_ffff) {
    valueFailure(part, cell, message);
  }
  return parsed;
}

function optionalBoolean(
  value: string | undefined,
  part: string,
  message: string,
): boolean | undefined {
  if (value === undefined) return undefined;
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  valueFailure(part, undefined, message);
}

function optionalHeight(
  value: string | undefined,
  part: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (!NONNEGATIVE_DECIMAL_PATTERN.test(value)) {
    valueFailure(part, undefined, 'Worksheet row height is invalid');
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    valueFailure(
      part,
      undefined,
      'Worksheet row height is outside the finite range',
    );
  }
  return parsed;
}

function optionalWidth(
  value: string | undefined,
  part: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (!NONNEGATIVE_DECIMAL_PATTERN.test(value)) {
    valueFailure(part, undefined, 'Worksheet column width is invalid');
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed > 255) {
    valueFailure(part, undefined, 'Worksheet column width is invalid');
  }
  return parsed;
}

function resolvedStyle(
  rawStyle: number | undefined,
  styles: XlsxStyleTable,
  part: string,
  location: string | undefined,
  message: string,
) {
  if (rawStyle === undefined) return undefined;
  const style = styles.cellXfs[rawStyle];
  if (!style) valueFailure(part, location, message);
  return style;
}

function cellType(
  value: string | undefined,
  part: string,
  cell: string,
): CellType {
  if (value === undefined) return 'n';
  if (
    value === 'b' ||
    value === 'd' ||
    value === 'e' ||
    value === 'inlineStr' ||
    value === 'n' ||
    value === 's' ||
    value === 'str'
  ) {
    return value;
  }
  valueFailure(part, cell, 'Worksheet cell type is invalid');
}

function validateXmlSpace(
  element: XlsxXmlElement,
  part: string,
  cell: string,
): void {
  const value = element.attributes.get(`{${XML_NAMESPACE}}space`);
  if (value !== undefined && value !== 'default' && value !== 'preserve') {
    valueFailure(part, cell, 'Inline-string xml:space value is invalid');
  }
}

export function consumeXlsxWorksheetBudget(
  budget: XlsxWorksheetBudget,
  key: keyof XlsxWorksheetBudget,
  amount: number,
  limitName:
    | 'maxConditionalFormattingRules'
    | 'maxFormulaGroups'
    | 'maxRangeAreas'
    | 'maxReturnedCells'
    | 'maxRichTextRuns'
    | 'maxScannedCells'
    | 'maxTextCharacters'
    | 'maxValidationRules',
  limits: ResolvedXlsxResourceLimits,
  part: string,
): void {
  const actual = budget[key] + amount;
  if (!Number.isSafeInteger(actual) || actual > limits[limitName]) {
    throw new XlsxResourceLimitError(
      limitName,
      actual,
      limits[limitName],
      part,
    );
  }
  budget[key] = actual;
}

export function createXlsxWorksheetBudget(
  sharedStrings: XlsxSharedStringTable,
  initial: Partial<
    Pick<XlsxWorksheetBudget, 'formulaCharacters' | 'textCharacters'>
  > = {},
): XlsxWorksheetBudget {
  let richTextRuns = 0;
  let textCharacters = 0;
  for (const value of sharedStrings.values) {
    textCharacters += value.text.length;
    richTextRuns += value.runs?.length ?? 0;
    for (const run of value.phoneticRuns ?? []) {
      textCharacters += run.text.length;
      richTextRuns += 1;
    }
  }
  return {
    conditionalFormattingRules: 0,
    formulaCharacters: initial.formulaCharacters ?? 0,
    formulaGroups: 0,
    rangeAreas: 0,
    returnedCells: 0,
    richTextRuns,
    scannedCells: 0,
    textCharacters: textCharacters + (initial.textCharacters ?? 0),
    validationRules: 0,
  };
}

export function consumeXlsxWorksheetFormulaCharacters(
  budget: XlsxWorksheetBudget,
  expression: string,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): void {
  if (expression.length > limits.maxFormulaCharacters) {
    throw new XlsxResourceLimitError(
      'maxFormulaCharacters',
      expression.length,
      limits.maxFormulaCharacters,
      part,
    );
  }
  const actual = budget.formulaCharacters + expression.length;
  if (
    !Number.isSafeInteger(actual) ||
    actual > limits.maxTotalFormulaCharacters
  ) {
    throw new XlsxResourceLimitError(
      'maxTotalFormulaCharacters',
      actual,
      limits.maxTotalFormulaCharacters,
      part,
    );
  }
  budget.formulaCharacters = actual;
}

class WorksheetSink implements XlsxXmlEventSink {
  private autoFilter: XlsxAutoFilter | undefined;
  private autoFilterCapture: XlsxAutoFilterCapture | undefined;
  private autoFilterSeen = false;
  private readonly authoredColumns: XlsxAuthoredColumnRange[] = [];
  private capture: TextCapture = null;
  private columnsSeen = false;
  private readonly conditionalFormattingPriorities = new Set<number>();
  private conditionalFormattingCapture:
    XlsxConditionalFormattingCapture | undefined;
  private readonly conditionalFormattings: XlsxConditionalFormatting[] = [];
  private dataValidationSettings: XlsxDataValidationSettings | undefined;
  private dataValidations: XlsxDataValidation[] = [];
  private dataValidationsCapture: XlsxDataValidationsCapture | undefined;
  private dataValidationsSeen = false;
  private currentCell: PendingCell | undefined;
  private currentInlineRun: PendingInlineRun | undefined;
  private currentRow: XlsxRow | undefined;
  private currentRowSelected!: boolean;
  private currentView: XlsxWorksheetView | undefined;
  private declaredDimension: XlsxRange | undefined;
  private dimensionSeen = false;
  private drawingSeen = false;
  private ignoredDepth = 0;
  private readonly hyperlinks: ParsedXlsxHyperlink[] = [];
  private hyperlinksSeen = false;
  private lastCellColumn = 0;
  private lastRow = 0;
  private legacyDrawingSeen = false;
  private mergeCellsExpected: number | undefined;
  private mergeCellsSeen = false;
  private readonly mergedRanges: XlsxRange[] = [];
  private readonly selectedColumnPrefix: Uint32Array;
  private sheetDataSeen = false;
  private sheetFormat: XlsxWorksheetFormat | undefined;
  private sheetFormatSeen = false;
  private sheetPropertiesSeen = false;
  private sheetViewsSeen = false;
  private sparklineGroups: XlsxSparklineGroup[] = [];
  private extensionsCapture: XlsxWorksheetExtensionsCapture | undefined;
  private extensionsSeen = false;
  private unsupportedExtensionsSeen = false;
  private unsupportedMetadataSeen = false;
  private readonly metadataBudget: XlsxCellMetadataBudget;
  private readonly metadataRegistry: XlsxCellMetadataRegistry;
  private readonly stack: XlsxXmlElement[] = [];
  private readonly rows: XlsxRow[] = [];
  private readonly viewIds = new Set<number>();
  private readonly views: XlsxWorksheetView[] = [];
  private outline: XlsxWorksheetOutline | undefined;
  private outlineSeen = false;
  private columnBreaks: XlsxPageBreak[] | undefined;
  private columnBreaksSeen = false;
  private headerFooter: XlsxHeaderFooter | undefined;
  private headerFooterCapture: XlsxHeaderFooterCapture | undefined;
  private headerFooterSeen = false;
  private pageBreaksCapture: XlsxPageBreaksCapture | undefined;
  private pageBreaksKind: 'column' | 'row' | undefined;
  private pageMargins: XlsxPageMargins | undefined;
  private pageMarginsSeen = false;
  private pageSetup: XlsxPageSetup | undefined;
  private pageSetupProperties: XlsxPageSetupProperties | undefined;
  private pageSetupPropertiesSeen = false;
  private pageSetupSeen = false;
  private pivotTablePartsExpected: number | undefined;
  private pivotTablePartsSeen = false;
  private readonly pivotTableRelationshipIds: string[] = [];
  private readonly pivotTableRelationshipIdSet = new Set<string>();
  private printOptions: XlsxPrintOptions | undefined;
  private printOptionsSeen = false;
  private rowBreaks: XlsxPageBreak[] | undefined;
  private rowBreaksSeen = false;
  private readonly protectedRanges: XlsxProtectedRange[] = [];
  private protectedRangesCapture: XlsxProtectedRangesCapture | undefined;
  private protectedRangesSeen = false;
  private protection: XlsxWorksheetProtection | undefined;
  private protectionSeen = false;
  private tabColor: XlsxColor | undefined;
  private tabColorSeen = false;
  private tablePartsExpected: number | undefined;
  private readonly tableRelationshipIds: string[] = [];
  private readonly tableRelationshipIdSet = new Set<string>();
  private tablePartsSeen = false;
  private readonly sharedFormulaMasters = new Map<
    number,
    SharedFormulaMaster
  >();

  constructor(
    private readonly part: string,
    private readonly namespace: string,
    private readonly sharedStrings: XlsxSharedStringTable,
    private readonly budget: XlsxWorksheetBudget,
    private readonly limits: ResolvedXlsxResourceLimits,
    private readonly selection: XlsxResolvedSheetSelection,
    private readonly semantics: XlsxWorksheetSemantics,
  ) {
    this.metadataBudget = semantics.metadataBudget ?? { records: 0 };
    this.metadataRegistry =
      semantics.metadataRegistry ?? EMPTY_XLSX_CELL_METADATA;
    this.selectedColumnPrefix = new Uint32Array(
      this.limits.maxColumnsPerWorksheet + 1,
    );
    if (selection.kind === 'selected-ranges') {
      const differences = new Int32Array(
        this.limits.maxColumnsPerWorksheet + 2,
      );
      for (const range of selection.ranges) {
        const start = Math.min(
          range.start.column,
          this.limits.maxColumnsPerWorksheet,
        );
        const end = Math.min(
          range.end.column,
          this.limits.maxColumnsPerWorksheet,
        );
        differences[start]! += 1;
        differences[end + 1]! -= 1;
      }
      let active = 0;
      let selected = 0;
      Array.from(
        { length: this.limits.maxColumnsPerWorksheet },
        (_, index) => index + 1,
      ).forEach((column) => {
        active += differences[column]!;
        if (active > 0) selected += 1;
        this.selectedColumnPrefix[column] = selected;
      });
    }
  }

  openElement(element: XlsxXmlElement): void {
    if (this.ignoredDepth > 0) {
      this.ignoredDepth += 1;
      this.stack.push(element);
      return;
    }
    if (this.extensionsCapture) {
      this.extensionsCapture.openElement(element);
      this.stack.push(element);
      return;
    }
    if (element.namespace !== this.namespace) {
      structureFailure(
        this.part,
        this.currentCell?.address,
        'Worksheet element has an unsupported namespace',
      );
    }
    if (this.autoFilterCapture) {
      this.autoFilterCapture.openElement(element);
      this.stack.push(element);
      return;
    }
    if (this.dataValidationsCapture) {
      this.dataValidationsCapture.openElement(element);
      this.stack.push(element);
      return;
    }
    if (this.conditionalFormattingCapture) {
      this.conditionalFormattingCapture.openElement(element);
      this.stack.push(element);
      return;
    }
    if (this.protectedRangesCapture) {
      this.protectedRangesCapture.openElement(element);
      this.stack.push(element);
      return;
    }
    if (this.headerFooterCapture) {
      this.headerFooterCapture.openElement(element);
      this.stack.push(element);
      return;
    }
    if (this.pageBreaksCapture) {
      this.pageBreaksCapture.openElement(element);
      this.stack.push(element);
      return;
    }
    const parent = this.stack.at(-1);
    if (!parent) {
      if (element.localName !== 'worksheet') {
        structureFailure(this.part, undefined, 'Worksheet root is missing');
      }
      this.stack.push(element);
      return;
    }

    this.openChild(parent.localName, element);
    this.stack.push(element);
  }

  closeElement(element: XlsxXmlElement): void {
    if (this.ignoredDepth > 0) {
      this.ignoredDepth -= 1;
      this.stack.pop();
      return;
    }
    if (this.extensionsCapture) {
      const capture = this.extensionsCapture;
      capture.closeElement(element);
      if (element.localName === 'extLst') {
        this.sparklineGroups = capture.result();
        if (capture.hasUnsupportedExtension()) {
          this.unsupportedExtensionsSeen = true;
        }
        this.extensionsCapture = undefined;
      }
      this.stack.pop();
      return;
    }
    if (this.autoFilterCapture) {
      this.autoFilterCapture.closeElement(element);
      if (element.localName === 'autoFilter') {
        this.autoFilter = this.autoFilterCapture.result();
        this.autoFilterCapture = undefined;
      }
      this.stack.pop();
      return;
    }
    if (this.dataValidationsCapture) {
      this.dataValidationsCapture.closeElement(element);
      if (element.localName === 'dataValidations') {
        const result = this.dataValidationsCapture.result();
        this.dataValidations = result.rules;
        this.dataValidationSettings = result.settings;
        this.dataValidationsCapture = undefined;
      }
      this.stack.pop();
      return;
    }
    if (this.conditionalFormattingCapture) {
      this.conditionalFormattingCapture.closeElement(element);
      if (element.localName === 'conditionalFormatting') {
        const result = this.conditionalFormattingCapture.result();
        if (result) this.conditionalFormattings.push(result);
        this.conditionalFormattingCapture = undefined;
      }
      this.stack.pop();
      return;
    }
    if (this.protectedRangesCapture) {
      this.protectedRangesCapture.closeElement(element);
      if (element.localName === 'protectedRanges') {
        this.protectedRanges.push(...this.protectedRangesCapture.result());
        this.protectedRangesCapture = undefined;
      }
      this.stack.pop();
      return;
    }
    if (this.headerFooterCapture) {
      this.headerFooterCapture.closeElement(element);
      if (element.localName === 'headerFooter') {
        this.headerFooter = this.headerFooterCapture.result();
        this.headerFooterCapture = undefined;
      }
      this.stack.pop();
      return;
    }
    if (this.pageBreaksCapture) {
      this.pageBreaksCapture.closeElement(element);
      if (
        element.localName === 'rowBreaks' ||
        element.localName === 'colBreaks'
      ) {
        const result = this.pageBreaksCapture.result();
        if (this.pageBreaksKind === 'row') this.rowBreaks = result;
        else this.columnBreaks = result;
        this.pageBreaksCapture = undefined;
        this.pageBreaksKind = undefined;
      }
      this.stack.pop();
      return;
    }
    this.capture = null;
    if (element.localName === 'r') this.closeInlineRun();
    if (element.localName === 'c') this.closeCell();
    if (element.localName === 'row') this.closeRow();
    if (element.localName === 'sheetView') this.closeView();
    this.stack.pop();
  }

  text(value: string): void {
    if (this.ignoredDepth > 0) return;
    if (this.extensionsCapture) {
      this.extensionsCapture.text(value);
      return;
    }
    if (this.autoFilterCapture) {
      this.autoFilterCapture.text(value);
      return;
    }
    if (this.dataValidationsCapture) {
      this.dataValidationsCapture.text(value);
      return;
    }
    if (this.conditionalFormattingCapture) {
      this.conditionalFormattingCapture.text(value);
      return;
    }
    if (this.protectedRangesCapture) {
      this.protectedRangesCapture.text(value);
      return;
    }
    if (this.headerFooterCapture) {
      this.headerFooterCapture.text(value);
      return;
    }
    if (this.pageBreaksCapture) {
      this.pageBreaksCapture.text(value);
      return;
    }
    if (this.capture === 'value') {
      this.currentCell!.valueText += value;
      return;
    }
    if (this.capture === 'formula') {
      this.currentCell!.formula!.text += value;
      return;
    }
    if (this.capture === 'inline') {
      if (this.currentInlineRun) this.currentInlineRun.text += value;
      else this.currentCell!.inlineText += value;
      return;
    }
    if (value.trim().length > 0) {
      structureFailure(
        this.part,
        this.currentCell?.address,
        'Worksheet text is outside a value or inline-string text element',
      );
    }
  }

  result(): XlsxWorksheetPayload {
    if (!this.sheetDataSeen) {
      structureFailure(this.part, undefined, 'Worksheet sheetData is missing');
    }
    if (this.sheetViewsSeen && this.views.length === 0) {
      structureFailure(
        this.part,
        undefined,
        'Worksheet sheetViews collection is empty',
      );
    }
    if (
      this.mergeCellsExpected !== undefined &&
      this.mergeCellsExpected !== this.mergedRanges.length
    ) {
      structureFailure(
        this.part,
        undefined,
        'Worksheet merged-range count does not match',
      );
    }
    if (xlsxMergedRangesOverlap(this.mergedRanges)) {
      valueFailure(this.part, undefined, 'Worksheet merged ranges overlap');
    }
    if (
      this.tablePartsExpected !== undefined &&
      this.tablePartsExpected !== this.tableRelationshipIds.length
    ) {
      structureFailure(
        this.part,
        undefined,
        'Worksheet table-part count does not match',
      );
    }
    if (
      this.pivotTablePartsExpected !== undefined &&
      this.pivotTablePartsExpected !== this.pivotTableRelationshipIds.length
    ) {
      structureFailure(
        this.part,
        undefined,
        'Worksheet pivot-table-part count does not match',
      );
    }
    const print = this.printSettings();
    return {
      ...(this.autoFilter === undefined ? {} : { autoFilter: this.autoFilter }),
      columns: normalizeXlsxColumnRanges(this.authoredColumns).filter((range) =>
        this.columnRangeSelected(range),
      ),
      conditionalFormattings: this.conditionalFormattings,
      ...(this.declaredDimension === undefined
        ? {}
        : { declaredDimension: this.declaredDimension }),
      ...(this.dataValidationSettings === undefined
        ? {}
        : { dataValidationSettings: this.dataValidationSettings }),
      dataValidations: this.dataValidations,
      hyperlinks: this.hyperlinks.flatMap((hyperlink) => {
        const selectionRelation = this.featureSelectionRelation(
          hyperlink.range,
        );
        return selectionRelation === null
          ? []
          : [
              {
                ...(hyperlink.display === undefined
                  ? {}
                  : { display: hyperlink.display }),
                range: hyperlink.range,
                selectionRelation,
                target: hyperlink.target,
                ...(hyperlink.tooltip === undefined
                  ? {}
                  : { tooltip: hyperlink.tooltip }),
              },
            ];
      }),
      mergedRanges: this.mergedRanges.filter((range) =>
        this.mergedRangeSelected(range),
      ),
      ...(this.outline === undefined ? {} : { outline: this.outline }),
      ...(print === undefined ? {} : { print }),
      protectedRanges: this.protectedRanges,
      ...(this.protection === undefined ? {} : { protection: this.protection }),
      rows: this.rows,
      ...(this.sheetFormat === undefined
        ? {}
        : { sheetFormat: this.sheetFormat }),
      ...(this.sparklineGroups.length === 0
        ? {}
        : { sparklineGroups: this.sparklineGroups }),
      ...(this.tabColor === undefined ? {} : { tabColor: this.tabColor }),
      ...(this.unsupportedExtensionsSeen
        ? { unsupportedExtensions: true as const }
        : {}),
      ...(this.unsupportedMetadataSeen
        ? { unsupportedMetadata: true as const }
        : {}),
      views: this.views,
    };
  }

  private columnRangeSelected(range: XlsxColumnRange): boolean {
    if (this.selection.kind !== 'selected-ranges') {
      return this.selection.kind === 'full-sheet';
    }
    return (
      this.selectedColumnPrefix[range.end]! -
        this.selectedColumnPrefix[range.start - 1]! >
      0
    );
  }

  private mergedRangeSelected(range: XlsxRange): boolean {
    return this.featureSelectionRelation(range) !== null;
  }

  private featureSelectionRelation(
    range: XlsxRange,
  ): XlsxHyperlink['selectionRelation'] | null {
    if (this.selection.kind !== 'selected-ranges') {
      return this.selection.kind === 'full-sheet' ? 'full-sheet' : null;
    }
    for (const selected of this.selection.ranges) {
      consumeXlsxWorksheetBudget(
        this.budget,
        'scannedCells',
        1,
        'maxScannedCells',
        this.limits,
        this.part,
      );
      if (
        selected.start.row <= range.end.row &&
        selected.end.row >= range.start.row &&
        selected.start.column <= range.end.column &&
        selected.end.column >= range.start.column
      ) {
        return 'intersects-selection';
      }
    }
    return null;
  }

  private beginIgnore(): void {
    this.ignoredDepth = 1;
  }

  private printSettings(): XlsxWorksheetPrintSettings | undefined {
    if (
      this.columnBreaks === undefined &&
      this.headerFooter === undefined &&
      this.pageMargins === undefined &&
      this.pageSetup === undefined &&
      this.pageSetupProperties === undefined &&
      this.printOptions === undefined &&
      this.rowBreaks === undefined
    ) {
      return undefined;
    }
    return {
      ...(this.columnBreaks === undefined
        ? {}
        : { columnBreaks: this.columnBreaks }),
      ...(this.headerFooter === undefined
        ? {}
        : { headerFooter: this.headerFooter }),
      ...(this.pageMargins === undefined ? {} : { margins: this.pageMargins }),
      ...(this.printOptions === undefined
        ? {}
        : { options: this.printOptions }),
      ...(this.pageSetup === undefined ? {} : { pageSetup: this.pageSetup }),
      ...(this.pageSetupProperties === undefined
        ? {}
        : { properties: this.pageSetupProperties }),
      ...(this.rowBreaks === undefined ? {} : { rowBreaks: this.rowBreaks }),
    };
  }

  private openChild(parent: string, element: XlsxXmlElement): void {
    if (parent === 'worksheet') {
      if (element.localName === 'extLst') {
        if (this.extensionsSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate extension lists',
          );
        }
        this.extensionsSeen = true;
        this.extensionsCapture = new XlsxWorksheetExtensionsCapture(
          this.namespace,
          this.selection,
          this.budget,
          this.limits,
          this.part,
        );
        this.extensionsCapture.openElement(element);
        return;
      }
      if (element.localName === 'dimension') {
        if (this.dimensionSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate dimension elements',
          );
        }
        this.dimensionSeen = true;
        this.declaredDimension = parseXlsxDeclaredDimension(element, this.part);
        consumeXlsxWorksheetBudget(
          this.budget,
          'rangeAreas',
          1,
          'maxRangeAreas',
          this.limits,
          this.part,
        );
        return;
      }
      if (element.localName === 'legacyDrawing') {
        if (this.legacyDrawingSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate legacyDrawing elements',
          );
        }
        this.legacyDrawingSeen = true;
        const relationshipId = element.attributes.get(
          `{${TABLE_RELATIONSHIP_NAMESPACE[this.semantics.dialect]}}id`,
        );
        if (relationshipId === undefined || relationshipId.length === 0) {
          valueFailure(
            this.part,
            undefined,
            'Worksheet legacy drawing relationship reference is invalid',
          );
        }
        this.semantics.legacyDrawingRelationshipIds?.push(relationshipId);
        return;
      }
      if (element.localName === 'drawing') {
        if (this.drawingSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate drawing elements',
          );
        }
        this.drawingSeen = true;
        const relationshipId = element.attributes.get(
          `{${TABLE_RELATIONSHIP_NAMESPACE[this.semantics.dialect]}}id`,
        );
        if (relationshipId === undefined || relationshipId.length === 0) {
          valueFailure(
            this.part,
            undefined,
            'Worksheet drawing relationship reference is invalid',
          );
        }
        this.semantics.drawingRelationshipIds?.push(relationshipId);
        return;
      }
      if (element.localName === 'sheetProtection') {
        if (this.protectionSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate sheetProtection elements',
          );
        }
        this.protectionSeen = true;
        this.protection = parseXlsxWorksheetProtection(
          element,
          this.budget,
          this.limits,
          this.part,
        );
        return;
      }
      if (element.localName === 'protectedRanges') {
        if (this.protectedRangesSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate protectedRanges elements',
          );
        }
        this.protectedRangesSeen = true;
        this.protectedRangesCapture = new XlsxProtectedRangesCapture(
          this.selection,
          this.budget,
          this.limits,
          this.part,
        );
        this.protectedRangesCapture.openElement(element);
        return;
      }
      if (element.localName === 'printOptions') {
        if (this.printOptionsSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate printOptions elements',
          );
        }
        this.printOptionsSeen = true;
        this.printOptions = parseXlsxPrintOptions(element, this.part);
        return;
      }
      if (element.localName === 'pageMargins') {
        if (this.pageMarginsSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate pageMargins elements',
          );
        }
        this.pageMarginsSeen = true;
        this.pageMargins = parseXlsxPageMargins(element, this.part);
        return;
      }
      if (element.localName === 'pageSetup') {
        if (this.pageSetupSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate pageSetup elements',
          );
        }
        this.pageSetupSeen = true;
        this.pageSetup = parseXlsxPageSetup(element, this.part);
        return;
      }
      if (element.localName === 'headerFooter') {
        if (this.headerFooterSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate headerFooter elements',
          );
        }
        this.headerFooterSeen = true;
        this.headerFooterCapture = new XlsxHeaderFooterCapture(
          this.budget,
          this.limits,
          this.part,
        );
        this.headerFooterCapture.openElement(element);
        return;
      }
      if (
        element.localName === 'rowBreaks' ||
        element.localName === 'colBreaks'
      ) {
        const row = element.localName === 'rowBreaks';
        if (row ? this.rowBreaksSeen : this.columnBreaksSeen) {
          structureFailure(
            this.part,
            undefined,
            row
              ? 'Worksheet contains duplicate rowBreaks elements'
              : 'Worksheet contains duplicate colBreaks elements',
          );
        }
        if (row) this.rowBreaksSeen = true;
        else this.columnBreaksSeen = true;
        this.pageBreaksKind = row ? 'row' : 'column';
        this.pageBreaksCapture = new XlsxPageBreaksCapture(
          this.pageBreaksKind,
          this.budget,
          this.limits,
          this.part,
        );
        this.pageBreaksCapture.openElement(element);
        return;
      }
      if (element.localName === 'sheetPr') {
        if (this.sheetPropertiesSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate sheetPr elements',
          );
        }
        this.sheetPropertiesSeen = true;
        return;
      }
      if (element.localName === 'sheetFormatPr') {
        if (this.sheetFormatSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate sheetFormatPr elements',
          );
        }
        this.sheetFormatSeen = true;
        this.sheetFormat = parseXlsxWorksheetFormat(element, this.part);
        return;
      }
      if (element.localName === 'cols') {
        if (this.columnsSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate cols elements',
          );
        }
        this.columnsSeen = true;
        return;
      }
      if (element.localName === 'sheetData') {
        if (this.sheetDataSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate sheetData elements',
          );
        }
        this.sheetDataSeen = true;
        return;
      }
      if (element.localName === 'sheetViews') {
        if (this.sheetViewsSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate sheetViews elements',
          );
        }
        this.sheetViewsSeen = true;
        return;
      }
      if (element.localName === 'mergeCells') {
        if (this.mergeCellsSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate mergeCells elements',
          );
        }
        this.mergeCellsSeen = true;
        this.mergeCellsExpected = unsignedInteger(
          attribute(element, 'count'),
          this.part,
          undefined,
          'Worksheet merged-range count is invalid',
        );
        if (
          this.mergeCellsExpected !== undefined &&
          this.mergeCellsExpected > this.limits.maxMergedRanges
        ) {
          throw new XlsxResourceLimitError(
            'maxMergedRanges',
            this.mergeCellsExpected,
            this.limits.maxMergedRanges,
            this.part,
          );
        }
        return;
      }
      if (element.localName === 'hyperlinks') {
        if (this.hyperlinksSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate hyperlinks elements',
          );
        }
        this.hyperlinksSeen = true;
        return;
      }
      if (element.localName === 'autoFilter') {
        if (this.autoFilterSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate autoFilter elements',
          );
        }
        this.autoFilterSeen = true;
        this.autoFilterCapture = new XlsxAutoFilterCapture(
          this.semantics.styles.differentialStyles.length,
          this.selection,
          this.budget,
          this.limits,
          this.part,
        );
        this.autoFilterCapture.openElement(element);
        return;
      }
      if (element.localName === 'dataValidations') {
        if (this.dataValidationsSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate dataValidations elements',
          );
        }
        this.dataValidationsSeen = true;
        this.dataValidationsCapture = new XlsxDataValidationsCapture(
          this.selection,
          this.budget,
          this.limits,
          this.part,
        );
        this.dataValidationsCapture.openElement(element);
        return;
      }
      if (element.localName === 'conditionalFormatting') {
        this.conditionalFormattingCapture =
          new XlsxConditionalFormattingCapture(
            this.semantics.styles.differentialStyles.length,
            this.conditionalFormattingPriorities,
            this.selection,
            this.budget,
            this.limits,
            this.part,
          );
        this.conditionalFormattingCapture.openElement(element);
        return;
      }
      if (element.localName === 'tableParts') {
        if (this.tablePartsSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate tableParts elements',
          );
        }
        this.tablePartsSeen = true;
        this.tablePartsExpected = unsignedInteger(
          attribute(element, 'count'),
          this.part,
          undefined,
          'Worksheet table-part count is invalid',
        );
        if (
          this.tablePartsExpected === undefined ||
          this.tablePartsExpected === 0
        ) {
          valueFailure(
            this.part,
            undefined,
            'Worksheet table-part count is invalid',
          );
        }
        if (this.tablePartsExpected > this.limits.maxTables) {
          throw new XlsxResourceLimitError(
            'maxTables',
            this.tablePartsExpected,
            this.limits.maxTables,
            this.part,
          );
        }
        return;
      }
      if (element.localName === 'pivotTableParts') {
        if (this.pivotTablePartsSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate pivotTableParts elements',
          );
        }
        this.pivotTablePartsSeen = true;
        this.pivotTablePartsExpected = unsignedInteger(
          attribute(element, 'count'),
          this.part,
          undefined,
          'Worksheet pivot-table-part count is invalid',
        );
        if (
          this.pivotTablePartsExpected === undefined ||
          this.pivotTablePartsExpected === 0
        ) {
          valueFailure(
            this.part,
            undefined,
            'Worksheet pivot-table parts must not be empty',
          );
        }
        return;
      }
      this.beginIgnore();
      return;
    }
    if (parent === 'cols' && element.localName === 'col') {
      this.openColumn(element);
      return;
    }
    if (parent === 'sheetPr') {
      if (element.localName === 'tabColor') {
        if (this.tabColorSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate tab colors',
          );
        }
        this.tabColorSeen = true;
        this.tabColor = parseXlsxWorksheetTabColor(element, this.part);
        return;
      }
      if (element.localName === 'outlinePr') {
        if (this.outlineSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate outline properties',
          );
        }
        this.outlineSeen = true;
        this.outline = parseXlsxWorksheetOutline(element, this.part);
        return;
      }
      if (element.localName === 'pageSetUpPr') {
        if (this.pageSetupPropertiesSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate page setup properties',
          );
        }
        this.pageSetupPropertiesSeen = true;
        this.pageSetupProperties = parseXlsxPageSetupProperties(
          element,
          this.part,
        );
        return;
      }
    }
    if (parent === 'sheetViews' && element.localName === 'sheetView') {
      this.openView(element);
      return;
    }
    if (parent === 'hyperlinks' && element.localName === 'hyperlink') {
      this.openHyperlink(element);
      return;
    }
    if (parent === 'tableParts' && element.localName === 'tablePart') {
      this.openTablePart(element);
      return;
    }
    if (
      parent === 'pivotTableParts' &&
      element.localName === 'pivotTablePart'
    ) {
      const relationshipId = element.attributes.get(
        `{${TABLE_RELATIONSHIP_NAMESPACE[this.semantics.dialect]}}id`,
      );
      if (relationshipId === undefined || relationshipId.length === 0) {
        valueFailure(
          this.part,
          undefined,
          'Worksheet pivot-table relationship reference is invalid',
        );
      }
      if (this.pivotTableRelationshipIdSet.has(relationshipId)) {
        valueFailure(
          this.part,
          undefined,
          'Worksheet contains duplicate pivot-table relationships',
        );
      }
      this.pivotTableRelationshipIdSet.add(relationshipId);
      this.pivotTableRelationshipIds.push(relationshipId);
      this.semantics.pivotTableRelationshipIds?.push(relationshipId);
      return;
    }
    if (parent === 'sheetView') {
      if (element.localName === 'pane') {
        this.openPane(element);
        return;
      }
      if (element.localName === 'selection') {
        this.openViewSelection(element);
        return;
      }
      if (
        element.localName === 'extLst' ||
        element.localName === 'pivotSelection'
      ) {
        if (element.localName === 'extLst') {
          this.unsupportedExtensionsSeen = true;
        }
        this.beginIgnore();
        return;
      }
    }
    if (parent === 'mergeCells' && element.localName === 'mergeCell') {
      this.openMergedRange(element);
      return;
    }
    if (parent === 'sheetData' && element.localName === 'row') {
      this.openRow(element);
      return;
    }
    if (parent === 'row') {
      if (element.localName === 'c') {
        this.openCell(element);
        return;
      }
      if (element.localName === 'extLst') {
        this.unsupportedExtensionsSeen = true;
        this.beginIgnore();
        return;
      }
    }
    if (parent === 'c') {
      if (element.localName === 'f') {
        this.openFormula(element);
        return;
      }
      if (element.localName === 'v') {
        this.openValue();
        return;
      }
      if (element.localName === 'is') {
        this.openInlineString();
        return;
      }
      if (element.localName === 'extLst') {
        this.unsupportedExtensionsSeen = true;
        this.beginIgnore();
        return;
      }
    }
    if (parent === 'is') {
      if (element.localName === 't') {
        this.openInlinePlainText(element);
        return;
      }
      if (element.localName === 'r') {
        this.openInlineRun();
        return;
      }
      if (element.localName === 'rPh' || element.localName === 'phoneticPr') {
        this.beginIgnore();
        return;
      }
    }
    if (parent === 'r') {
      if (element.localName === 'rPr') {
        this.beginIgnore();
        return;
      }
      if (element.localName === 't') {
        this.openInlineRunText(element);
        return;
      }
    }
    structureFailure(
      this.part,
      this.currentCell?.address,
      'Worksheet element nesting is invalid',
    );
  }

  private openColumn(element: XlsxXmlElement): void {
    const start = unsignedInteger(
      attribute(element, 'min'),
      this.part,
      undefined,
      'Worksheet column start is invalid',
    );
    const end = unsignedInteger(
      attribute(element, 'max'),
      this.part,
      undefined,
      'Worksheet column end is invalid',
    );
    if (start === undefined || start === 0) {
      valueFailure(this.part, undefined, 'Worksheet column start is invalid');
    }
    if (end === undefined || end < start) {
      valueFailure(this.part, undefined, 'Worksheet column end is invalid');
    }
    if (end > this.limits.maxColumnsPerWorksheet) {
      throw new XlsxResourceLimitError(
        'maxColumnsPerWorksheet',
        end,
        this.limits.maxColumnsPerWorksheet,
        this.part,
      );
    }
    const actualRanges = this.authoredColumns.length + 1;
    if (actualRanges > this.limits.maxColumnsPerWorksheet) {
      throw new XlsxResourceLimitError(
        'maxColumnsPerWorksheet',
        actualRanges,
        this.limits.maxColumnsPerWorksheet,
        this.part,
      );
    }
    const outlineLevel = unsignedInteger(
      attribute(element, 'outlineLevel'),
      this.part,
      undefined,
      'Worksheet column outline level is invalid',
    );
    if (outlineLevel !== undefined && outlineLevel > 7) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet column outline level is invalid',
      );
    }
    const collapsed = optionalBoolean(
      attribute(element, 'collapsed'),
      this.part,
      'Worksheet column collapsed flag is invalid',
    );
    const hidden = optionalBoolean(
      attribute(element, 'hidden'),
      this.part,
      'Worksheet column hidden flag is invalid',
    );
    optionalBoolean(
      attribute(element, 'bestFit'),
      this.part,
      'Worksheet column bestFit flag is invalid',
    );
    optionalBoolean(
      attribute(element, 'customWidth'),
      this.part,
      'Worksheet column customWidth flag is invalid',
    );
    optionalBoolean(
      attribute(element, 'phonetic'),
      this.part,
      'Worksheet column phonetic flag is invalid',
    );
    const width = optionalWidth(attribute(element, 'width'), this.part);
    const rawStyle = unsignedInteger(
      attribute(element, 'style'),
      this.part,
      undefined,
      'Worksheet column style index is invalid',
    );
    const style = resolvedStyle(
      rawStyle,
      this.semantics.styles,
      this.part,
      undefined,
      'Worksheet column style reference is invalid',
    );
    this.authoredColumns.push({
      ...(collapsed === undefined ? {} : { collapsed }),
      end,
      ...(hidden === undefined ? {} : { hidden }),
      order: this.authoredColumns.length,
      ...(outlineLevel === undefined ? {} : { outlineLevel }),
      start,
      ...(style === undefined ? {} : { style: style.normalizedStyle }),
      ...(width === undefined ? {} : { width }),
    });
  }

  private openHyperlink(element: XlsxXmlElement): void {
    const actual = this.hyperlinks.length + 1;
    if (actual > this.limits.maxHyperlinks) {
      throw new XlsxResourceLimitError(
        'maxHyperlinks',
        actual,
        this.limits.maxHyperlinks,
        this.part,
      );
    }
    const hyperlink = parseXlsxHyperlink(
      element,
      this.semantics.dialect,
      this.semantics.relationships,
      this.part,
    );
    consumeXlsxWorksheetBudget(
      this.budget,
      'rangeAreas',
      1,
      'maxRangeAreas',
      this.limits,
      this.part,
    );
    consumeXlsxWorksheetBudget(
      this.budget,
      'textCharacters',
      hyperlink.textCharacters,
      'maxTextCharacters',
      this.limits,
      this.part,
    );
    this.hyperlinks.push(hyperlink);
  }

  private openTablePart(element: XlsxXmlElement): void {
    const relationshipId = element.attributes.get(
      `{${TABLE_RELATIONSHIP_NAMESPACE[this.semantics.dialect]}}id`,
    );
    if (relationshipId === undefined || relationshipId.length === 0) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet table relationship reference is invalid',
      );
    }
    if (this.tableRelationshipIdSet.has(relationshipId)) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet contains duplicate table relationship references',
      );
    }
    this.tableRelationshipIdSet.add(relationshipId);
    this.tableRelationshipIds.push(relationshipId);
    this.semantics.tableRelationshipIds?.push(relationshipId);
  }

  private openView(element: XlsxXmlElement): void {
    const view = parseXlsxWorksheetView(element, this.part);
    if (view.workbookViewId >= this.semantics.workbookViewCount) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet workbook view reference is out of range',
      );
    }
    if (this.viewIds.has(view.workbookViewId)) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet contains duplicate workbook view references',
      );
    }
    this.viewIds.add(view.workbookViewId);
    this.views.push(view);
    this.currentView = view;
  }

  private openPane(element: XlsxXmlElement): void {
    if (this.currentView!.pane) {
      structureFailure(
        this.part,
        undefined,
        'Worksheet view contains duplicate pane elements',
      );
    }
    this.currentView!.pane = parseXlsxWorksheetPane(
      element,
      this.part,
      this.limits,
    );
  }

  private openViewSelection(element: XlsxXmlElement): void {
    const parsed = parseXlsxWorksheetViewSelection(element, this.part);
    consumeXlsxWorksheetBudget(
      this.budget,
      'rangeAreas',
      parsed.rangeAreaCount,
      'maxRangeAreas',
      this.limits,
      this.part,
    );
    this.currentView!.selections.push(parsed.selection);
  }

  private closeView(): void {
    validateXlsxWorksheetView(this.currentView!, this.part);
    this.currentView = undefined;
  }

  private openMergedRange(element: XlsxXmlElement): void {
    const source = attribute(element, 'ref');
    const range = parseXlsxRangeReference(source);
    if (!range || source?.includes('$')) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet merged-range reference is invalid',
      );
    }
    if (
      range.start.row === range.end.row &&
      range.start.column === range.end.column
    ) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet merged range must contain multiple cells',
      );
    }
    const actual = this.mergedRanges.length + 1;
    if (actual > this.limits.maxMergedRanges) {
      throw new XlsxResourceLimitError(
        'maxMergedRanges',
        actual,
        this.limits.maxMergedRanges,
        this.part,
      );
    }
    this.mergedRanges.push(range);
  }

  private openRow(element: XlsxXmlElement): void {
    const reference = unsignedInteger(
      attribute(element, 'r'),
      this.part,
      undefined,
      'Worksheet row reference is invalid',
    );
    const index = reference ?? this.lastRow + 1;
    if (index <= this.lastRow) {
      valueFailure(this.part, undefined, 'Worksheet rows are out of order');
    }
    if (index > this.limits.maxRowsPerWorksheet) {
      throw new XlsxResourceLimitError(
        'maxRowsPerWorksheet',
        index,
        this.limits.maxRowsPerWorksheet,
        this.part,
      );
    }
    const outlineLevel = unsignedInteger(
      attribute(element, 'outlineLevel'),
      this.part,
      undefined,
      'Worksheet row outline level is invalid',
    );
    if (outlineLevel !== undefined && outlineLevel > 7) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet row outline level is invalid',
      );
    }
    const height = optionalHeight(attribute(element, 'ht'), this.part);
    const collapsed = optionalBoolean(
      attribute(element, 'collapsed'),
      this.part,
      'Worksheet row collapsed flag is invalid',
    );
    const customFormat = optionalBoolean(
      attribute(element, 'customFormat'),
      this.part,
      'Worksheet row customFormat flag is invalid',
    );
    optionalBoolean(
      attribute(element, 'customHeight'),
      this.part,
      'Worksheet row customHeight flag is invalid',
    );
    const hidden = optionalBoolean(
      attribute(element, 'hidden'),
      this.part,
      'Worksheet row hidden flag is invalid',
    );
    const rawStyle = unsignedInteger(
      attribute(element, 's'),
      this.part,
      undefined,
      'Worksheet row style index is invalid',
    );
    if (customFormat === true && rawStyle === undefined) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet custom-formatted row style is missing',
      );
    }
    const style = resolvedStyle(
      rawStyle,
      this.semantics.styles,
      this.part,
      undefined,
      'Worksheet row style reference is invalid',
    );
    this.currentRow = {
      cells: [],
      ...(collapsed === undefined ? {} : { collapsed }),
      ...(height === undefined ? {} : { height }),
      ...(hidden === undefined ? {} : { hidden }),
      index,
      ...(outlineLevel === undefined ? {} : { outlineLevel }),
      ...(style === undefined ? {} : { style: style.normalizedStyle }),
    };
    this.currentRowSelected = xlsxSelectionIncludesRow(this.selection, index);
    this.lastCellColumn = 0;
    this.lastRow = index;
  }

  private closeRow(): void {
    if (this.currentRowSelected) this.rows.push(this.currentRow!);
    this.currentRow = undefined;
  }

  private openCell(element: XlsxXmlElement): void {
    const sourceReference = attribute(element, 'r');
    let address: string;
    let column: number;
    if (sourceReference === undefined) {
      column = this.lastCellColumn + 1;
      if (column > this.limits.maxColumnsPerWorksheet) {
        throw new XlsxResourceLimitError(
          'maxColumnsPerWorksheet',
          column,
          this.limits.maxColumnsPerWorksheet,
          this.part,
        );
      }
      address = `${xlsxColumnName(column)!}${this.currentRow!.index}`;
    } else {
      const parsed = parseXlsxCellReference(sourceReference);
      if (!parsed || parsed.absoluteColumn || parsed.absoluteRow) {
        valueFailure(
          this.part,
          undefined,
          'Worksheet cell reference is invalid',
        );
      }
      address = parsed.address;
      column = parsed.column;
      if (parsed.row !== this.currentRow!.index) {
        valueFailure(
          this.part,
          address,
          'Worksheet cell reference does not belong to its row',
        );
      }
      if (column > this.limits.maxColumnsPerWorksheet) {
        throw new XlsxResourceLimitError(
          'maxColumnsPerWorksheet',
          column,
          this.limits.maxColumnsPerWorksheet,
          this.part,
        );
      }
    }
    if (column <= this.lastCellColumn) {
      valueFailure(this.part, address, 'Worksheet cells are out of order');
    }
    const rawStyle = unsignedInteger(
      attribute(element, 's'),
      this.part,
      address,
      'Worksheet cell style index is invalid',
    );
    const styleXf = resolvedStyle(
      rawStyle,
      this.semantics.styles,
      this.part,
      address,
      'Worksheet cell style reference is invalid',
    );
    consumeXlsxWorksheetBudget(
      this.budget,
      'scannedCells',
      1,
      'maxScannedCells',
      this.limits,
      this.part,
    );
    const selected = xlsxSelectionIncludesCell(
      this.selection,
      this.currentRow!.index,
      column,
    );
    if (selected) {
      consumeXlsxWorksheetBudget(
        this.budget,
        'returnedCells',
        1,
        'maxReturnedCells',
        this.limits,
        this.part,
      );
    }
    const cellMetadata = resolveXlsxCellMetadata(
      this.metadataRegistry,
      'cell',
      unsignedInteger(
        attribute(element, 'cm'),
        this.part,
        address,
        'Worksheet cell metadata index is invalid',
      ),
      this.metadataBudget,
      this.budget,
      this.limits,
      this.part,
      address,
    );
    const valueMetadata = resolveXlsxCellMetadata(
      this.metadataRegistry,
      'value',
      unsignedInteger(
        attribute(element, 'vm'),
        this.part,
        address,
        'Worksheet value metadata index is invalid',
      ),
      this.metadataBudget,
      this.budget,
      this.limits,
      this.part,
      address,
    );
    if (cellMetadata?.unsupported || valueMetadata?.unsupported) {
      this.unsupportedMetadataSeen = true;
    }
    const metadata: XlsxCellMetadata | undefined =
      (cellMetadata?.entries.length ?? 0) === 0 &&
      (valueMetadata?.entries.length ?? 0) === 0
        ? undefined
        : {
            ...(cellMetadata?.entries.length
              ? { cell: cellMetadata.entries }
              : {}),
            ...(valueMetadata?.entries.length
              ? { value: valueMetadata.entries }
              : {}),
          };
    this.currentCell = {
      address,
      column,
      hasInlineString: false,
      hasValue: false,
      inlineMode: 'unset',
      inlineRuns: [],
      inlineText: '',
      ...(metadata === undefined ? {} : { metadata }),
      ...(styleXf?.numberFormat === undefined
        ? {}
        : { numberFormat: styleXf.numberFormat }),
      selected,
      ...(styleXf === undefined ? {} : { style: styleXf.normalizedStyle }),
      type: cellType(attribute(element, 't'), this.part, address),
      valueText: '',
    };
    this.lastCellColumn = column;
  }

  private openValue(): void {
    const cell = this.currentCell!;
    if (cell.hasValue || cell.hasInlineString || cell.type === 'inlineStr') {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet cell value structure is invalid',
      );
    }
    cell.hasValue = true;
    this.capture = 'value';
  }

  private openFormula(element: XlsxXmlElement): void {
    const cell = this.currentCell!;
    if (
      cell.formula !== undefined ||
      cell.hasValue ||
      cell.hasInlineString ||
      cell.type === 'inlineStr' ||
      cell.type === 's'
    ) {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet formula structure is invalid',
      );
    }
    const sourceType = attribute(element, 't');
    const type = sourceType ?? 'normal';
    if (
      type !== 'normal' &&
      type !== 'shared' &&
      type !== 'array' &&
      type !== 'dataTable'
    ) {
      formulaFailure(this.part, cell.address, 'Formula type is invalid');
    }
    const reference = attribute(element, 'ref');
    const sharedIndex = attribute(element, 'si');
    cell.formula = {
      ...(reference === undefined ? {} : { reference }),
      ...(sharedIndex === undefined ? {} : { sharedIndex }),
      text: '',
      type,
    };
    this.capture = 'formula';
  }

  private openInlineString(): void {
    const cell = this.currentCell!;
    if (cell.hasValue || cell.hasInlineString || cell.type !== 'inlineStr') {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet inline-string structure is invalid',
      );
    }
    cell.hasInlineString = true;
  }

  private openInlinePlainText(element: XlsxXmlElement): void {
    const cell = this.currentCell!;
    if (cell.inlineMode !== 'unset') {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet inline-string plain text is out of order',
      );
    }
    cell.inlineMode = 'plain';
    validateXmlSpace(element, this.part, cell.address);
    this.capture = 'inline';
  }

  private openInlineRun(): void {
    const cell = this.currentCell!;
    if (cell.inlineMode === 'plain') {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet inline-string rich text is out of order',
      );
    }
    cell.inlineMode = 'rich';
    consumeXlsxWorksheetBudget(
      this.budget,
      'richTextRuns',
      1,
      'maxRichTextRuns',
      this.limits,
      this.part,
    );
    this.currentInlineRun = { hasText: false, text: '' };
  }

  private openInlineRunText(element: XlsxXmlElement): void {
    const cell = this.currentCell!;
    if (this.currentInlineRun!.hasText) {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet inline-string run has duplicate text',
      );
    }
    this.currentInlineRun!.hasText = true;
    validateXmlSpace(element, this.part, cell.address);
    this.capture = 'inline';
  }

  private closeInlineRun(): void {
    const cell = this.currentCell!;
    if (!this.currentInlineRun!.hasText) {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet inline-string run text is missing',
      );
    }
    cell.inlineRuns.push({ text: this.currentInlineRun!.text });
    cell.inlineText += this.currentInlineRun!.text;
    this.currentInlineRun = undefined;
  }

  private closeCell(): void {
    const cell = this.currentCell!;
    let content: XlsxCell['content'];
    if (cell.formula !== undefined) {
      const formula = this.resolveFormula(cell);
      const cached = cell.hasValue
        ? this.applyNumberFormat(
            parseXlsxScalarCellValue(
              cell.type as XlsxScalarCellType,
              cell.valueText,
              this.sharedStrings,
              this.part,
              cell.address,
            ),
            cell,
          )
        : ({ kind: 'missing' } as const);
      if (cell.selected && cached.kind !== 'missing') {
        this.consumeReturnedText(cached);
      }
      content = { cached, formula, kind: 'formula' };
    } else if (!cell.hasValue && !cell.hasInlineString) {
      content = { kind: 'blank' };
    } else if (cell.hasInlineString) {
      const value: XlsxCellValue = {
        kind: 'text',
        ...(cell.inlineMode === 'rich' ? { runs: cell.inlineRuns } : {}),
        text: cell.inlineText,
      };
      if (cell.selected) this.consumeTextCharacters(value.text);
      content = { kind: 'value', value };
    } else {
      const value = this.applyNumberFormat(
        parseXlsxScalarCellValue(
          cell.type as XlsxScalarCellType,
          cell.valueText,
          this.sharedStrings,
          this.part,
          cell.address,
        ),
        cell,
      );
      if (cell.selected) this.consumeReturnedText(value);
      content = { kind: 'value', value };
    }
    if (!cell.selected) {
      this.currentCell = undefined;
      return;
    }
    const base = {
      address: cell.address,
      column: cell.column,
      ...(cell.metadata === undefined ? {} : { metadata: cell.metadata }),
      ...(cell.style === undefined ? {} : { style: cell.style }),
    };
    if (content.kind === 'blank') {
      this.currentRow!.cells.push({ ...base, content: { kind: 'blank' } });
    } else if (content.kind === 'formula') {
      this.currentRow!.cells.push({ ...base, content });
    } else {
      this.currentRow!.cells.push({
        ...base,
        content: { kind: 'value', value: content.value },
      });
    }
    this.currentCell = undefined;
  }

  private applyNumberFormat(
    value: XlsxCellValue,
    cell: PendingCell,
  ): XlsxCellValue {
    if (value.kind !== 'number' || cell.numberFormat === undefined)
      return value;
    const precision = xlsxNumberFormatDatePrecision(
      cell.numberFormat,
      value.value,
    );
    if (precision === undefined) return value;
    return {
      kind: 'date',
      normalized: normalizeXlsxSerialDate(
        value.value,
        this.semantics.dateSystem,
        precision,
      ),
      precision,
      source: {
        dateSystem: this.semantics.dateSystem,
        kind: 'serial',
        value: value.value,
      },
    };
  }

  private resolveFormula(cell: PendingCell): XlsxFormula {
    const pending = cell.formula!;
    if (pending.text.startsWith('=')) {
      formulaFailure(
        this.part,
        cell.address,
        'Formula expression must not include a leading equals sign',
      );
    }
    if (pending.type === 'normal') {
      if (
        pending.text.length === 0 ||
        pending.reference !== undefined ||
        pending.sharedIndex !== undefined
      ) {
        formulaFailure(this.part, cell.address, 'Normal formula is invalid');
      }
      this.consumeFormulaCharacters(pending.text);
      return { expression: pending.text, kind: 'normal' };
    }
    if (pending.type === 'shared') return this.resolveSharedFormula(cell);

    if (pending.sharedIndex !== undefined) {
      formulaFailure(
        this.part,
        cell.address,
        'Grouped formula shared index is invalid',
      );
    }
    const range = this.formulaRange(pending.reference, cell);
    if (
      range.start.row !== this.currentRow!.index ||
      range.start.column !== cell.column
    ) {
      formulaFailure(
        this.part,
        cell.address,
        'Grouped formula must start at its owning cell',
      );
    }
    if (pending.type === 'array' && pending.text.length === 0) {
      formulaFailure(this.part, cell.address, 'Array formula is empty');
    }
    this.consumeFormulaGroup();
    this.consumeFormulaCharacters(pending.text);
    return {
      expression: pending.text,
      kind: pending.type === 'array' ? 'array' : 'data-table',
      range,
    };
  }

  private resolveSharedFormula(cell: PendingCell): XlsxFormula {
    const pending = cell.formula!;
    const sharedIndex = this.formulaSharedIndex(pending.sharedIndex, cell);
    if (pending.reference !== undefined) {
      if (
        pending.text.length === 0 ||
        this.sharedFormulaMasters.has(sharedIndex)
      ) {
        formulaFailure(
          this.part,
          cell.address,
          'Shared formula master is invalid',
        );
      }
      const range = this.formulaRange(pending.reference, cell);
      if (
        range.start.row !== this.currentRow!.index ||
        range.start.column !== cell.column
      ) {
        formulaFailure(
          this.part,
          cell.address,
          'Shared formula master must start at its owning cell',
        );
      }
      this.consumeFormulaGroup();
      this.sharedFormulaMasters.set(sharedIndex, {
        column: cell.column,
        expression: pending.text,
        range,
        row: this.currentRow!.index,
      });
      this.consumeFormulaCharacters(pending.text);
      return { expression: pending.text, kind: 'normal' };
    }
    if (pending.text.length !== 0) {
      formulaFailure(
        this.part,
        cell.address,
        'Shared formula dependent contains an expression',
      );
    }
    const master = this.sharedFormulaMasters.get(sharedIndex);
    if (
      master === undefined ||
      !this.rangeContains(master.range, this.currentRow!.index, cell.column)
    ) {
      formulaFailure(
        this.part,
        cell.address,
        'Shared formula master is missing or does not own the cell',
      );
    }
    const expression = translateXlsxSharedFormula(
      master.expression,
      { column: master.column, row: master.row },
      { column: cell.column, row: this.currentRow!.index },
    );
    if (expression === undefined) {
      formulaFailure(
        this.part,
        cell.address,
        'Shared formula translation is outside the worksheet grid',
      );
    }
    this.consumeFormulaCharacters(expression);
    return { expression, kind: 'normal' };
  }

  private formulaRange(
    value: string | undefined,
    cell: PendingCell,
  ): XlsxRange {
    const range = parseXlsxRangeReference(value);
    if (!range) {
      formulaFailure(this.part, cell.address, 'Formula range is invalid');
    }
    return range;
  }

  private formulaSharedIndex(
    value: string | undefined,
    cell: PendingCell,
  ): number {
    if (value === undefined || !UNSIGNED_INTEGER_PATTERN.test(value)) {
      formulaFailure(
        this.part,
        cell.address,
        'Shared formula index is invalid',
      );
    }
    const parsed = Number(value);
    if (!Number.isSafeInteger(parsed) || parsed > 0xffff_ffff) {
      formulaFailure(
        this.part,
        cell.address,
        'Shared formula index is invalid',
      );
    }
    return parsed;
  }

  private rangeContains(
    range: XlsxRange,
    row: number,
    column: number,
  ): boolean {
    return row <= range.end.row && column <= range.end.column;
  }

  private consumeFormulaGroup(): void {
    consumeXlsxWorksheetBudget(
      this.budget,
      'formulaGroups',
      1,
      'maxFormulaGroups',
      this.limits,
      this.part,
    );
  }

  private consumeFormulaCharacters(expression: string): void {
    consumeXlsxWorksheetFormulaCharacters(
      this.budget,
      expression,
      this.limits,
      this.part,
    );
  }

  private consumeTextCharacters(value: string): void {
    consumeXlsxWorksheetBudget(
      this.budget,
      'textCharacters',
      value.length,
      'maxTextCharacters',
      this.limits,
      this.part,
    );
  }

  private consumeReturnedText(value: XlsxCellValue): void {
    if (value.kind !== 'text') return;
    this.consumeTextCharacters(value.text);
    if (value.runs) {
      consumeXlsxWorksheetBudget(
        this.budget,
        'richTextRuns',
        value.runs.length,
        'maxRichTextRuns',
        this.limits,
        this.part,
      );
    }
  }
}

export async function parseXlsxWorksheetPart(
  part: string,
  dialect: XlsxWorkbookDiscovery['dialect'],
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  sharedStrings: XlsxSharedStringTable,
  budget: XlsxWorksheetBudget,
  selection: XlsxResolvedSheetSelection = { kind: 'full-sheet' },
  semantics: XlsxWorksheetSemantics = DEFAULT_WORKSHEET_SEMANTICS,
): Promise<XlsxWorksheetPayload> {
  const sink = new WorksheetSink(
    part,
    XLSX_SPREADSHEET_NAMESPACES[dialect],
    sharedStrings,
    budget,
    limits,
    selection,
    semantics,
  );
  await reader.streamXml(part, sink, { required: true });
  return sink.result();
}
