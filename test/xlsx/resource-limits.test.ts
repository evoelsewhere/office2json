import { describe, expect, it } from 'vitest';

import {
  defaultXlsxResourceLimits,
  resourceLimitDiagnostic,
  resolveXlsxResourceLimits,
  XlsxResourceLimitError,
} from '../../src/formats/xlsx/internal/resource-limits';
import type { XlsxResourceLimits } from '../../src/formats/xlsx/types';

const EXPECTED_DEFAULTS: Required<XlsxResourceLimits> = {
  maxCalculationChainEntries: 250_000,
  maxCharts: 10_000,
  maxColumnsPerWorksheet: 16_384,
  maxComments: 100_000,
  maxConditionalFormattingRules: 100_000,
  maxDefinedNames: 100_000,
  maxDrawings: 100_000,
  maxEntries: 10_000,
  maxFormulaCharacters: 8_192,
  maxFormulaGroups: 250_000,
  maxHyperlinks: 100_000,
  maxInputBytes: 100 * 1024 * 1024,
  maxMediaBytes: 64 * 1024 * 1024,
  maxMetadataRecords: 250_000,
  maxMergedRanges: 100_000,
  maxPartBytes: 64 * 1024 * 1024,
  maxPivotRecords: 100_000,
  maxRangeAreas: 100_000,
  maxRelationships: 100_000,
  maxReturnedCells: 250_000,
  maxRichTextRuns: 500_000,
  maxRowsPerWorksheet: 1_048_576,
  maxScannedCells: 1_000_000,
  maxSharedStrings: 1_000_000,
  maxStyles: 65_536,
  maxTables: 10_000,
  maxTextCharacters: 16 * 1024 * 1024,
  maxTotalFormulaCharacters: 8 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxTotalXmlNodes: 1_000_000,
  maxValidationRules: 100_000,
  maxWorksheets: 1_000,
  maxXmlBytes: 16 * 1024 * 1024,
  maxXmlDepth: 128,
  maxXmlNodes: 250_000,
};

describe('XLSX resource limits', () => {
  it('publishes reviewed safe defaults', () => {
    expect(defaultXlsxResourceLimits()).toEqual(EXPECTED_DEFAULTS);
  });

  it('resolves a partial override without mutating it', () => {
    const options: XlsxResourceLimits = { maxReturnedCells: 17 };
    const before = structuredClone(options);

    expect(resolveXlsxResourceLimits(options)).toEqual({
      ...EXPECTED_DEFAULTS,
      maxReturnedCells: 17,
    });
    expect(options).toEqual(before);
  });

  it.each([0, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY])(
    'rejects invalid configured limit %s',
    (value) => {
      expect(() => resolveXlsxResourceLimits({ maxEntries: value })).toThrow(
        'XLSX resource limit maxEntries must be a positive safe integer',
      );
    },
  );

  it('rejects a configured value above the safe-integer range', () => {
    expect(() =>
      resolveXlsxResourceLimits({ maxEntries: Number.MAX_SAFE_INTEGER + 1 }),
    ).toThrow('XLSX resource limit maxEntries must be a positive safe integer');
  });

  it.each([
    [
      { maxPartBytes: 9, maxXmlBytes: 10 },
      'maxXmlBytes cannot exceed maxPartBytes',
    ],
    [
      { maxMediaBytes: 10, maxPartBytes: 9, maxXmlBytes: 9 },
      'maxMediaBytes cannot exceed maxPartBytes',
    ],
    [
      { maxReturnedCells: 10, maxScannedCells: 9 },
      'maxReturnedCells cannot exceed maxScannedCells',
    ],
    [
      { maxFormulaCharacters: 10, maxTotalFormulaCharacters: 9 },
      'maxFormulaCharacters cannot exceed maxTotalFormulaCharacters',
    ],
    [
      { maxRowsPerWorksheet: 1_048_577 },
      'maxRowsPerWorksheet cannot exceed the XLSX row limit',
    ],
    [
      { maxColumnsPerWorksheet: 16_385 },
      'maxColumnsPerWorksheet cannot exceed the XLSX column limit',
    ],
  ] as const)('rejects incompatible limits %#', (limits, message) => {
    expect(() => resolveXlsxResourceLimits(limits)).toThrow(message);
  });

  it('accepts every dependent relationship exactly at its boundary', () => {
    expect(
      resolveXlsxResourceLimits({
        maxColumnsPerWorksheet: 16_384,
        maxFormulaCharacters: 13,
        maxMediaBytes: 17,
        maxPartBytes: 17,
        maxReturnedCells: 19,
        maxRowsPerWorksheet: 1_048_576,
        maxScannedCells: 19,
        maxTotalFormulaCharacters: 13,
        maxXmlBytes: 17,
      }),
    ).toMatchObject({
      maxColumnsPerWorksheet: 16_384,
      maxFormulaCharacters: 13,
      maxMediaBytes: 17,
      maxPartBytes: 17,
      maxReturnedCells: 19,
      maxRowsPerWorksheet: 1_048_576,
      maxScannedCells: 19,
      maxTotalFormulaCharacters: 13,
      maxXmlBytes: 17,
    });
  });

  it('exposes structured limit metadata', () => {
    const error = new XlsxResourceLimitError(
      'maxReturnedCells',
      11,
      10,
      'xl/worksheets/sheet1.xml',
    );

    expect(error).toMatchObject({
      actual: 11,
      limit: 10,
      limitName: 'maxReturnedCells',
      name: 'XlsxResourceLimitError',
      part: 'xl/worksheets/sheet1.xml',
    });
    expect(error.message).toBe(
      'XLSX resource limit maxReturnedCells exceeded for xl/worksheets/sheet1.xml: 11 > 10',
    );
    expect(resourceLimitDiagnostic(error)).toEqual({
      actual: 11,
      code: 'resource-limit-exceeded',
      limit: 10,
      limitName: 'maxReturnedCells',
      message: error.message,
      part: 'xl/worksheets/sheet1.xml',
      severity: 'error',
    });
  });

  it('omits part metadata when a limit is package-wide', () => {
    const error = new XlsxResourceLimitError('maxEntries', 11, 10);

    expect(error.message).toBe(
      'XLSX resource limit maxEntries exceeded: 11 > 10',
    );
    expect(error.part).toBeUndefined();
    expect(resourceLimitDiagnostic(error)).toEqual({
      actual: 11,
      code: 'resource-limit-exceeded',
      limit: 10,
      limitName: 'maxEntries',
      message: error.message,
      severity: 'error',
    });
  });
});
