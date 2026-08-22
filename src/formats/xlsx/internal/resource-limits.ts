import type { XlsxDiagnostic, XlsxResourceLimits } from '../types';

export const XLSX_MAX_ROWS = 1_048_576;
export const XLSX_MAX_COLUMNS = 16_384;

export type ResolvedXlsxResourceLimits = Required<XlsxResourceLimits>;

export function defaultXlsxResourceLimits(): ResolvedXlsxResourceLimits {
  const mebibyte = 1024 * 1024;
  return {
    maxCalculationChainEntries: 250_000,
    maxCharts: 10_000,
    maxColumnsPerWorksheet: XLSX_MAX_COLUMNS,
    maxComments: 100_000,
    maxConditionalFormattingRules: 100_000,
    maxDefinedNames: 100_000,
    maxDrawings: 100_000,
    maxEntries: 10_000,
    maxFormulaCharacters: 8_192,
    maxFormulaGroups: 250_000,
    maxHyperlinks: 100_000,
    maxInputBytes: 100 * mebibyte,
    maxMediaBytes: 64 * mebibyte,
    maxMetadataRecords: 250_000,
    maxMergedRanges: 100_000,
    maxPartBytes: 64 * mebibyte,
    maxPivotRecords: 100_000,
    maxRangeAreas: 100_000,
    maxRelationships: 100_000,
    maxReturnedCells: 250_000,
    maxRichTextRuns: 500_000,
    maxRowsPerWorksheet: XLSX_MAX_ROWS,
    maxScannedCells: 1_000_000,
    maxSharedStrings: 1_000_000,
    maxStyles: 65_536,
    maxTables: 10_000,
    maxTextCharacters: 16 * mebibyte,
    maxTotalFormulaCharacters: 8 * mebibyte,
    maxTotalUncompressedBytes: 256 * mebibyte,
    maxTotalXmlNodes: 1_000_000,
    maxValidationRules: 100_000,
    maxWorksheets: 1_000,
    maxXmlBytes: 16 * mebibyte,
    maxXmlDepth: 128,
    maxXmlNodes: 250_000,
  };
}

export class XlsxResourceLimitError extends Error {
  readonly actual: number;
  readonly limit: number;
  readonly limitName: keyof XlsxResourceLimits;
  readonly part?: string;

  constructor(
    limitName: keyof XlsxResourceLimits,
    actual: number,
    limit: number,
    part?: string,
  ) {
    const location = part ? ` for ${part}` : '';
    super(
      `XLSX resource limit ${limitName} exceeded${location}: ${actual} > ${limit}`,
    );
    this.name = 'XlsxResourceLimitError';
    this.actual = actual;
    this.limit = limit;
    this.limitName = limitName;
    if (part) this.part = part;
  }
}

function assertPositiveSafeInteger(
  name: keyof XlsxResourceLimits,
  value: number,
): void {
  if (!Number.isSafeInteger(value) || value <= 0) {
    throw new RangeError(
      `XLSX resource limit ${name} must be a positive safe integer`,
    );
  }
}

function assertNotGreater(
  lowerName: keyof XlsxResourceLimits,
  lowerValue: number,
  upperName: keyof XlsxResourceLimits,
  upperValue: number,
): void {
  if (lowerValue > upperValue) {
    throw new RangeError(`${lowerName} cannot exceed ${upperName}`);
  }
}

export function resolveXlsxResourceLimits(
  limits: XlsxResourceLimits = {},
): ResolvedXlsxResourceLimits {
  const resolved: ResolvedXlsxResourceLimits = {
    ...defaultXlsxResourceLimits(),
    ...limits,
  };
  for (const [name, value] of Object.entries(resolved)) {
    assertPositiveSafeInteger(name as keyof XlsxResourceLimits, value);
  }

  assertNotGreater(
    'maxXmlBytes',
    resolved.maxXmlBytes,
    'maxPartBytes',
    resolved.maxPartBytes,
  );
  assertNotGreater(
    'maxMediaBytes',
    resolved.maxMediaBytes,
    'maxPartBytes',
    resolved.maxPartBytes,
  );
  assertNotGreater(
    'maxReturnedCells',
    resolved.maxReturnedCells,
    'maxScannedCells',
    resolved.maxScannedCells,
  );
  assertNotGreater(
    'maxFormulaCharacters',
    resolved.maxFormulaCharacters,
    'maxTotalFormulaCharacters',
    resolved.maxTotalFormulaCharacters,
  );
  if (resolved.maxRowsPerWorksheet > XLSX_MAX_ROWS) {
    throw new RangeError(
      'maxRowsPerWorksheet cannot exceed the XLSX row limit',
    );
  }
  if (resolved.maxColumnsPerWorksheet > XLSX_MAX_COLUMNS) {
    throw new RangeError(
      'maxColumnsPerWorksheet cannot exceed the XLSX column limit',
    );
  }
  return resolved;
}

export function resourceLimitDiagnostic(
  error: XlsxResourceLimitError,
): XlsxDiagnostic {
  return {
    actual: error.actual,
    code: 'resource-limit-exceeded',
    limit: error.limit,
    limitName: error.limitName,
    message: error.message,
    ...(error.part ? { part: error.part } : {}),
    severity: 'error',
  };
}
