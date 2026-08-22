import {
  parseXlsxCellReference,
  xlsxColumnName,
} from '../internal/cell-reference';
import type { ResolvedXlsxResourceLimits } from '../internal/resource-limits';
import type {
  XlsxCell,
  XlsxColumnRange,
  XlsxAutoFilter,
  XlsxDrawingObject,
  XlsxHyperlink,
  XlsxRow,
  XlsxWorksheet,
} from '../types';
import { canonicalXlsxJson } from './canonical-json';
import { canonicalXlsxSha256 } from './digest';
import { XlsxWriteError } from './errors';
import { transformXlsxStructuralSourceFormula } from './formula-reference';
import {
  transformXlsxStructuralPageBreak,
  transformXlsxStructuralRange,
  transformXlsxStructuralCell,
  transformXlsxStructuralDrawingAnchor,
  transformXlsxStructuralViewSelection,
  transformXlsxStructuralVisualCell,
} from './structural-reference';
import {
  type XlsxCellEditOperation,
  validateXlsxCellOperations,
} from './operation-validation';
import type {
  ResolvedXlsxWriteLimits,
  XlsxRoundTripDocument,
  XlsxRoundTripSheet,
} from './types';
import { writeLimitFailure } from './write-limits';

interface XlsxOperationImpactBase {
  operationId: string;
  sheetKey: string;
}

export type XlsxCellOperationImpact =
  | (XlsxOperationImpactBase & {
      cell: string;
      kind: Extract<
        XlsxCellEditOperation['kind'],
        'clear-cell' | 'set-cell' | 'set-cell-style' | 'set-hyperlink'
      >;
    })
  | (XlsxOperationImpactBase & {
      kind: 'set-column';
      range: string;
    })
  | (XlsxOperationImpactBase & { kind: 'set-row'; range: string })
  | (XlsxOperationImpactBase & {
      kind: 'delete-columns' | 'delete-rows' | 'insert-columns' | 'insert-rows';
      range: string;
    });

export interface XlsxCellOperationPlan {
  document: XlsxRoundTripDocument;
  impacts: XlsxCellOperationImpact[];
  operations: XlsxCellEditOperation[];
  stateHash: string;
}

function cloneDocument(document: XlsxRoundTripDocument): XlsxRoundTripDocument {
  return JSON.parse(JSON.stringify(document)) as XlsxRoundTripDocument;
}

function operationFailure(
  code:
    | 'operation-precondition-failed'
    | 'preservation-conflict'
    | 'unsupported-edit-operation',
  message: string,
  operation: XlsxCellEditOperation,
  featureClass?: string,
): never {
  throw new XlsxWriteError(code, message, {
    ...('cell' in operation ? { cell: operation.cell } : {}),
    ...(featureClass === undefined ? {} : { featureClass }),
    operationId: operation.operationId,
    ...('row' in operation
      ? { range: String(operation.row) }
      : 'start' in operation
        ? { range: `${operation.start}:${operation.end}` }
        : {}),
    sheetKey: operation.sheetKey,
  });
}

function isCellOperation(
  operation: XlsxCellEditOperation,
): operation is Extract<XlsxCellEditOperation, { cell: string }> {
  return 'cell' in operation;
}

function resolveWorksheet(
  document: XlsxRoundTripDocument,
  operation: XlsxCellEditOperation,
): XlsxWorksheet & { key: string } {
  const sheet = document.sheets.find(
    (candidate) => candidate.key === operation.sheetKey,
  );
  if (!sheet) {
    operationFailure(
      'preservation-conflict',
      'XLSX operation sheet key does not exist in the snapshot',
      operation,
      'worksheet',
    );
  }
  if (sheet.kind !== 'worksheet') {
    operationFailure(
      'preservation-conflict',
      isCellOperation(operation)
        ? 'XLSX cell operation cannot target a chart sheet'
        : 'XLSX row or column operation cannot target a chart sheet',
      operation,
      'chart-sheet',
    );
  }
  return sheet;
}

function resolveRow(
  sheet: XlsxWorksheet,
  operation: Extract<XlsxCellEditOperation, { kind: 'set-row' }>,
): XlsxRow {
  const row = sheet.rows.find((candidate) => candidate.index === operation.row);
  if (!row) {
    operationFailure(
      'preservation-conflict',
      'XLSX set-row operation requires an existing explicit source row',
      operation,
      'missing-row',
    );
  }
  return row;
}

function resolveColumn(
  sheet: XlsxWorksheet,
  operation: Extract<XlsxCellEditOperation, { kind: 'set-column' }>,
): XlsxColumnRange {
  const column = sheet.columns.find(
    (candidate) =>
      candidate.start === operation.start && candidate.end === operation.end,
  );
  if (!column) {
    operationFailure(
      'preservation-conflict',
      'XLSX set-column operation requires an existing exact source range',
      operation,
      'missing-column-range',
    );
  }
  return column;
}

function resolveCell(
  sheet: XlsxWorksheet,
  operation: Extract<XlsxCellEditOperation, { cell: string }>,
): XlsxCell {
  const reference = parseXlsxCellReference(operation.cell)!;
  const row = sheet.rows.find((candidate) => candidate.index === reference.row);
  const cell = row?.cells.find(
    (candidate) => candidate.column === reference.column,
  );
  if (!cell) {
    operationFailure(
      'preservation-conflict',
      'XLSX cell operation requires an existing explicit source cell',
      operation,
      'missing-cell',
    );
  }
  const merged = sheet.mergedRanges.find(
    (range) =>
      reference.row >= range.start.row &&
      reference.row <= range.end.row &&
      reference.column >= range.start.column &&
      reference.column <= range.end.column,
  );
  if (
    merged &&
    (merged.start.row !== reference.row ||
      merged.start.column !== reference.column)
  ) {
    operationFailure(
      'preservation-conflict',
      'XLSX cell operation cannot target a non-anchor merged cell',
      operation,
      'merged-cell',
    );
  }
  return cell;
}

export function xlsxCellTargetState(
  sheet: XlsxRoundTripSheet,
  cell: XlsxCell,
): { cell: XlsxCell; hyperlink?: XlsxHyperlink; sheetKey: string } {
  const hyperlink =
    sheet.kind === 'worksheet'
      ? sheet.hyperlinks.find(
          (candidate) => candidate.range.reference === cell.address,
        )
      : undefined;
  return {
    cell,
    ...(hyperlink === undefined ? {} : { hyperlink }),
    sheetKey: sheet.key,
  };
}

export function xlsxRowTargetState(
  sheet: XlsxRoundTripSheet,
  row: XlsxRow,
): { row: XlsxRow; sheetKey: string } {
  return { row, sheetKey: sheet.key };
}

export function xlsxColumnTargetState(
  sheet: XlsxRoundTripSheet,
  column: XlsxColumnRange,
): { column: XlsxColumnRange; sheetKey: string } {
  return { column, sheetKey: sheet.key };
}

type XlsxStructuralOperation = Extract<
  XlsxCellEditOperation,
  { count: number }
>;

const STRUCTURAL_OPERATION_KINDS = new Set<XlsxCellEditOperation['kind']>([
  'delete-columns',
  'delete-rows',
  'insert-columns',
  'insert-rows',
]);

function isStructuralOperation(
  operation: XlsxCellEditOperation,
): operation is XlsxStructuralOperation {
  return STRUCTURAL_OPERATION_KINDS.has(operation.kind);
}

export function xlsxStructuralTargetState(sheet: XlsxRoundTripSheet): {
  sheet: XlsxRoundTripSheet;
  sheetKey: string;
} {
  return { sheet, sheetKey: sheet.key };
}

function structuralClosureFailure(
  operation: XlsxStructuralOperation,
  featureClass: string,
): never {
  operationFailure(
    'unsupported-edit-operation',
    'XLSX structural edit requires a reference-free worksheet closure',
    operation,
    featureClass,
  );
}

function blockStructuralFeature(
  operation: XlsxStructuralOperation,
  featureClass: string,
  blocked: boolean,
): void {
  if (blocked) structuralClosureFailure(operation, featureClass);
}

function assertStructuralClosure(
  document: XlsxRoundTripDocument,
  sheet: XlsxWorksheet,
  operation: XlsxStructuralOperation,
): void {
  const workbook = document.workbook;
  blockStructuralFeature(
    operation,
    'defined-name-reference',
    workbook.definedNames.length !== 0,
  );
  blockStructuralFeature(
    operation,
    'calculation-chain-reference',
    workbook.calculation.chain !== undefined,
  );
  for (const candidate of document.sheets) {
    if (candidate.kind !== 'worksheet') continue;
    for (const row of candidate.rows) {
      for (const cell of row.cells) {
        blockStructuralFeature(
          operation,
          'formula-reference',
          cell.content.kind === 'formula',
        );
      }
    }
  }
  const featureBlockers: Array<readonly [string, boolean]> = [
    ['protection-reference', sheet.protection !== undefined],
    ['pivot-reference', sheet.pivotTables !== undefined],
    ['query-table-reference', sheet.queryTables !== undefined],
    ['slicer-reference', sheet.slicers !== undefined],
    ['timeline-reference', sheet.timelines !== undefined],
  ];
  for (const [featureClass, blocked] of featureBlockers) {
    blockStructuralFeature(operation, featureClass, blocked);
  }
  blockStructuralFeature(
    operation,
    'data-validation-formula-reference',
    sheet.dataValidations.some(
      (validation) =>
        validation.formula1 !== undefined || validation.formula2 !== undefined,
    ),
  );
  blockStructuralFeature(
    operation,
    'view-pane-reference',
    sheet.views.some((view) => view.pane !== undefined),
  );
  blockStructuralFeature(
    operation,
    'conditional-format-formula-reference',
    sheet.conditionalFormattings.some((format) =>
      format.rules.some(
        (rule) =>
          rule.formulas.length !== 0 ||
          rule.colorScale?.stops.some(
            (stop) => stop.threshold.kind === 'formula',
          ) === true ||
          rule.dataBar?.thresholds.some(
            (threshold) => threshold.kind === 'formula',
          ) === true ||
          rule.iconSet?.thresholds.some(
            (threshold) => threshold.kind === 'formula',
          ) === true,
      ),
    ),
  );
  for (const table of sheet.tables) {
    blockStructuralFeature(
      operation,
      'table-formula-reference',
      table.columns.some(
        (column) =>
          column.calculatedFormula !== undefined ||
          column.totalsFormula !== undefined,
      ),
    );
    const transformed = transformXlsxStructuralRange(table.range, operation);
    blockStructuralFeature(
      operation,
      'table-range-deletion',
      transformed === null,
    );
    if (transformed === null) continue;
    blockStructuralFeature(
      operation,
      'table-column-structure',
      transformed.end.column - transformed.start.column !==
        table.range.end.column - table.range.start.column,
    );
    blockStructuralFeature(
      operation,
      'table-row-structure',
      transformed.end.row - transformed.start.row + 1 <
        Number(table.headerRow) + Number(table.totalsRow),
    );
    if (operation.kind === 'delete-rows') {
      const deletedEnd = operation.index + operation.count - 1;
      blockStructuralFeature(
        operation,
        'table-header-row',
        table.headerRow &&
          table.range.start.row >= operation.index &&
          table.range.start.row <= deletedEnd,
      );
      blockStructuralFeature(
        operation,
        'table-totals-row',
        table.totalsRow &&
          table.range.end.row >= operation.index &&
          table.range.end.row <= deletedEnd,
      );
    }
  }
  for (const comment of sheet.comments) {
    const reference = parseXlsxCellReference(comment.reference)!;
    blockStructuralFeature(
      operation,
      'comment-anchor-deletion',
      transformXlsxStructuralCell(
        reference.row,
        reference.column,
        operation,
      ) === null,
    );
  }
  const drawingHasChart = (object: XlsxDrawingObject): boolean =>
    object.kind === 'chart' ||
    (object.kind === 'group' && object.children.some(drawingHasChart));
  for (const drawing of sheet.drawings) {
    blockStructuralFeature(
      operation,
      'drawing-chart-reference',
      drawingHasChart(drawing.object),
    );
    blockStructuralFeature(
      operation,
      'drawing-anchor-deletion',
      transformXlsxStructuralDrawingAnchor(drawing, operation) === null,
    );
  }
  for (const group of sheet.sparklineGroups ?? []) {
    for (const sparkline of group.sparklines) {
      const formula = transformXlsxStructuralSourceFormula(
        sparkline.dataFormula,
        sheet.name,
        sheet.name,
        operation,
      );
      blockStructuralFeature(
        operation,
        'sparkline-formula-reference',
        formula.kind === 'unsupported',
      );
      blockStructuralFeature(
        operation,
        'sparkline-source-deletion',
        formula.kind === 'deleted',
      );
      const location = parseXlsxCellReference(sparkline.location)!;
      blockStructuralFeature(
        operation,
        'sparkline-location-deletion',
        transformXlsxStructuralCell(
          location.row,
          location.column,
          operation,
        ) === null,
      );
    }
  }
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      blockStructuralFeature(
        operation,
        'cell-metadata-reference',
        cell.metadata !== undefined,
      );
    }
  }
  const columnOperation =
    operation.kind === 'delete-columns' || operation.kind === 'insert-columns';
  blockStructuralFeature(
    operation,
    'column-definition',
    columnOperation && sheet.columns.length !== 0,
  );
}

function transformStructuralAutoFilter(
  autoFilter: XlsxAutoFilter,
  operation: XlsxStructuralOperation,
): XlsxAutoFilter | undefined {
  const filterRange = transformXlsxStructuralRange(autoFilter.range, operation);
  if (filterRange === null) return undefined;
  autoFilter.range = filterRange;
  if (autoFilter.sort !== undefined) {
    const sortRange = transformXlsxStructuralRange(
      autoFilter.sort.range,
      operation,
    );
    if (sortRange === null) {
      delete autoFilter.sort;
    } else {
      autoFilter.sort.range = sortRange;
      autoFilter.sort.conditions = autoFilter.sort.conditions.flatMap(
        (condition) => {
          const range = transformXlsxStructuralRange(
            condition.range,
            operation,
          );
          return range === null ? [] : [{ ...condition, range }];
        },
      );
    }
  }
  return autoFilter;
}

function transformStructuralLayoutReferences(
  sheet: XlsxWorksheet,
  operation: XlsxStructuralOperation,
): void {
  if (sheet.declaredDimension !== undefined) {
    const dimension = transformXlsxStructuralRange(
      sheet.declaredDimension,
      operation,
    );
    if (dimension === null) delete sheet.declaredDimension;
    else sheet.declaredDimension = dimension;
  }
  sheet.mergedRanges = sheet.mergedRanges.flatMap((range) => {
    const transformed = transformXlsxStructuralRange(range, operation);
    return transformed === null ? [] : [transformed];
  });
  sheet.hyperlinks = sheet.hyperlinks.flatMap((hyperlink) => {
    const transformed = transformXlsxStructuralRange(
      hyperlink.range,
      operation,
    );
    return transformed === null ? [] : [{ ...hyperlink, range: transformed }];
  });
  if (sheet.autoFilter !== undefined) {
    const autoFilter = transformStructuralAutoFilter(
      sheet.autoFilter,
      operation,
    );
    if (autoFilter === undefined) delete sheet.autoFilter;
    else sheet.autoFilter = autoFilter;
  }
  for (const table of sheet.tables) {
    table.range = transformXlsxStructuralRange(table.range, operation)!;
    if (table.autoFilter !== undefined) {
      const autoFilter = transformStructuralAutoFilter(
        table.autoFilter,
        operation,
      );
      if (autoFilter === undefined) delete table.autoFilter;
      else table.autoFilter = autoFilter;
    }
  }
  sheet.comments = sheet.comments.map((comment) => {
    const reference = parseXlsxCellReference(comment.reference)!;
    return {
      ...comment,
      reference: transformXlsxStructuralCell(
        reference.row,
        reference.column,
        operation,
      )!.address,
    };
  });
  sheet.drawings = sheet.drawings.map((drawing) => ({
    ...drawing,
    ...transformXlsxStructuralDrawingAnchor(drawing, operation)!,
  }));
  if (sheet.sparklineGroups !== undefined) {
    for (const group of sheet.sparklineGroups) {
      for (const sparkline of group.sparklines) {
        const formula = transformXlsxStructuralSourceFormula(
          sparkline.dataFormula,
          sheet.name,
          sheet.name,
          operation,
        );
        const location = parseXlsxCellReference(sparkline.location)!;
        const supportedFormula = formula as Extract<
          ReturnType<typeof transformXlsxStructuralSourceFormula>,
          { expression: string }
        >;
        sparkline.dataFormula = supportedFormula.expression;
        sparkline.location = transformXlsxStructuralCell(
          location.row,
          location.column,
          operation,
        )!.address;
      }
    }
  }
  const hadDataValidations = sheet.dataValidations.length !== 0;
  sheet.dataValidations = sheet.dataValidations.flatMap((validation) => {
    const ranges = validation.ranges.flatMap((range) => {
      const transformed = transformXlsxStructuralRange(range, operation);
      return transformed === null ? [] : [transformed];
    });
    return ranges.length === 0 ? [] : [{ ...validation, ranges }];
  });
  if (hadDataValidations && sheet.dataValidations.length === 0) {
    delete sheet.dataValidationSettings;
  }
  sheet.conditionalFormattings = sheet.conditionalFormattings.flatMap(
    (format) => {
      const ranges = format.ranges.flatMap((range) => {
        const transformed = transformXlsxStructuralRange(range, operation);
        return transformed === null ? [] : [transformed];
      });
      return ranges.length === 0 ? [] : [{ ...format, ranges }];
    },
  );
  sheet.protectedRanges = sheet.protectedRanges.flatMap((protectedRange) => {
    const ranges = protectedRange.ranges.flatMap((range) => {
      const transformed = transformXlsxStructuralRange(range, operation);
      return transformed === null ? [] : [transformed];
    });
    return ranges.length === 0 ? [] : [{ ...protectedRange, ranges }];
  });
  if (sheet.print !== undefined) {
    for (const [field, axis] of [
      ['columnBreaks', 'column'],
      ['rowBreaks', 'row'],
    ] as const) {
      const breaks = sheet.print[field];
      if (breaks === undefined) continue;
      const transformed = breaks.flatMap((pageBreak) => {
        const result = transformXlsxStructuralPageBreak(
          pageBreak,
          axis,
          operation,
        );
        return result === null ? [] : [result];
      });
      if (breaks.length !== 0 && transformed.length === 0) {
        delete sheet.print[field];
      } else {
        sheet.print[field] = transformed;
      }
    }
  }
  for (const view of sheet.views) {
    if (view.topLeftCell !== undefined) {
      view.topLeftCell = transformXlsxStructuralVisualCell(
        view.topLeftCell,
        operation,
      );
    }
    view.selections = view.selections.flatMap((selection) => {
      const transformed = transformXlsxStructuralViewSelection(
        selection,
        operation,
      );
      return transformed === null ? [] : [transformed];
    });
  }
}

function structuralRange(operation: XlsxStructuralOperation): string {
  const end = operation.index + operation.count - 1;
  return `${operation.index}:${end}`;
}

function transformRows(
  sheet: XlsxWorksheet,
  operation: XlsxStructuralOperation,
  readerLimits: ResolvedXlsxResourceLimits,
): number {
  const end = operation.index + operation.count - 1;
  const updates = sheet.rows
    .filter((row) => row.index >= operation.index)
    .reduce((total, row) => total + 1 + row.cells.length, 0);
  if (operation.kind === 'insert-rows') {
    for (const row of sheet.rows) {
      if (row.index < operation.index) continue;
      if (row.index + operation.count > readerLimits.maxRowsPerWorksheet) {
        operationFailure(
          'preservation-conflict',
          'XLSX row insertion would move an authored row outside the grid',
          operation,
          'grid-overflow',
        );
      }
      row.index += operation.count;
      for (const cell of row.cells) {
        cell.address = `${xlsxColumnName(cell.column)!}${row.index}`;
      }
    }
    return updates;
  }
  const precedingRows = sheet.rows.filter((row) => row.index < operation.index);
  const shiftedRows = sheet.rows.filter((row) => row.index > end);
  for (const row of shiftedRows) {
    row.index -= operation.count;
    for (const cell of row.cells) {
      cell.address = `${xlsxColumnName(cell.column)!}${row.index}`;
    }
  }
  sheet.rows = [...precedingRows, ...shiftedRows];
  return updates;
}

function transformColumns(
  sheet: XlsxWorksheet,
  operation: XlsxStructuralOperation,
  readerLimits: ResolvedXlsxResourceLimits,
): number {
  const end = operation.index + operation.count - 1;
  let updates = 0;
  for (const row of sheet.rows) {
    updates += row.cells.filter(
      (cell) => cell.column >= operation.index,
    ).length;
    if (operation.kind === 'insert-columns') {
      for (const cell of row.cells) {
        if (cell.column < operation.index) continue;
        if (
          cell.column + operation.count >
          readerLimits.maxColumnsPerWorksheet
        ) {
          operationFailure(
            'preservation-conflict',
            'XLSX column insertion would move an authored cell outside the grid',
            operation,
            'grid-overflow',
          );
        }
        cell.column += operation.count;
        cell.address = `${xlsxColumnName(cell.column)!}${row.index}`;
      }
      continue;
    }
    const precedingCells = row.cells.filter(
      (cell) => cell.column < operation.index,
    );
    const shiftedCells = row.cells.filter((cell) => cell.column > end);
    for (const cell of shiftedCells) {
      cell.column -= operation.count;
      cell.address = `${xlsxColumnName(cell.column)!}${row.index}`;
    }
    row.cells = [...precedingCells, ...shiftedCells];
  }
  return updates;
}

function applyRowOperation(
  row: XlsxRow,
  operation: Extract<XlsxCellEditOperation, { kind: 'set-row' }>,
): void {
  if (operation.height !== undefined) row.height = operation.height;
  if (operation.hidden !== undefined) row.hidden = operation.hidden;
}

function applyColumnOperation(
  column: XlsxColumnRange,
  operation: Extract<XlsxCellEditOperation, { kind: 'set-column' }>,
): void {
  if (operation.width !== undefined) column.width = operation.width;
  if (operation.hidden !== undefined) column.hidden = operation.hidden;
}

function applyCellOperation(
  document: XlsxRoundTripDocument,
  sheet: XlsxWorksheet,
  cell: XlsxCell,
  operation: Extract<XlsxCellEditOperation, { cell: string }>,
): void {
  if (operation.kind === 'set-hyperlink') {
    const reference = parseXlsxCellReference(operation.cell)!;
    const conflict = sheet.hyperlinks.find(
      (candidate) =>
        candidate.range.reference !== operation.cell &&
        reference.row >= candidate.range.start.row &&
        reference.row <= candidate.range.end.row &&
        reference.column >= candidate.range.start.column &&
        reference.column <= candidate.range.end.column,
    );
    if (conflict) {
      operationFailure(
        'preservation-conflict',
        'XLSX hyperlink operation overlaps a multi-cell hyperlink range',
        operation,
        'hyperlink-range',
      );
    }
    const index = sheet.hyperlinks.findIndex(
      (candidate) => candidate.range.reference === operation.cell,
    );
    if (operation.target === null) {
      if (index >= 0) sheet.hyperlinks.splice(index, 1);
      return;
    }
    if (index >= 0) {
      sheet.hyperlinks[index] = {
        ...sheet.hyperlinks[index]!,
        target: structuredClone(operation.target),
      };
      return;
    }
    sheet.hyperlinks.push({
      range: {
        end: { column: reference.column, row: reference.row },
        reference: reference.address,
        start: { column: reference.column, row: reference.row },
      },
      selectionRelation: 'full-sheet',
      target: structuredClone(operation.target),
    });
    return;
  }
  delete cell.displayText;
  if (operation.kind === 'set-cell-style') {
    const styleKey = canonicalXlsxJson(operation.style);
    let style = document.styles.findIndex(
      (candidate) => canonicalXlsxJson(candidate) === styleKey,
    );
    if (style < 0) {
      if (operation.style.checkbox === true) {
        operationFailure(
          'unsupported-edit-operation',
          'XLSX cannot append a checkbox style without a feature-property-bag edit',
          operation,
          'append-checkbox-style',
        );
      }
      document.styles.push(structuredClone(operation.style));
      style = document.styles.length - 1;
    }
    cell.style = style;
    return;
  }
  if (operation.kind === 'clear-cell') {
    cell.content = { kind: 'blank' };
    return;
  }
  cell.content =
    operation.content.kind === 'formula'
      ? {
          cached: { kind: 'missing' },
          formula: {
            expression: operation.content.expression,
            kind: 'normal',
          },
          kind: 'formula',
        }
      : { kind: 'value', value: structuredClone(operation.content.value) };
}

export async function replayXlsxCellOperations(
  baseDocument: XlsxRoundTripDocument,
  value: unknown,
  writeLimits: ResolvedXlsxWriteLimits,
  readerLimits: ResolvedXlsxResourceLimits,
): Promise<XlsxCellOperationPlan> {
  const operations = validateXlsxCellOperations(
    value,
    writeLimits,
    readerLimits,
  );
  const structuralOperation = operations.find(
    (operation): operation is XlsxStructuralOperation =>
      isStructuralOperation(operation),
  );
  if (
    structuralOperation &&
    operations.some((operation) => !isStructuralOperation(operation))
  ) {
    structuralClosureFailure(structuralOperation, 'mixed-operation-closure');
  }
  const document = cloneDocument(baseDocument);
  const sheetKeys = new Set<string>();
  for (const sheet of document.sheets) {
    if (sheetKeys.has(sheet.key)) {
      throw new XlsxWriteError(
        'snapshot-integrity-failed',
        'XLSX snapshot sheet keys must be unique',
        { objectKey: sheet.key },
      );
    }
    sheetKeys.add(sheet.key);
  }
  const impacts: XlsxCellOperationImpact[] = [];
  let referenceUpdates = 0;
  for (const operation of operations) {
    const sheet = resolveWorksheet(document, operation);
    if (isStructuralOperation(operation)) {
      assertStructuralClosure(document, sheet, operation);
      if (
        operation.ifMatch !== undefined &&
        operation.ifMatch !==
          (await canonicalXlsxSha256(xlsxStructuralTargetState(sheet)))
      ) {
        operationFailure(
          'operation-precondition-failed',
          'XLSX operation precondition does not match the target worksheet',
          operation,
        );
      }
      referenceUpdates +=
        (sheet.declaredDimension === undefined ? 0 : 1) +
        sheet.mergedRanges.length +
        sheet.hyperlinks.length +
        (sheet.autoFilter === undefined
          ? 0
          : 1 +
            (sheet.autoFilter.sort === undefined
              ? 0
              : 1 + sheet.autoFilter.sort.conditions.length)) +
        sheet.dataValidations.reduce(
          (total, validation) => total + validation.ranges.length,
          0,
        ) +
        sheet.conditionalFormattings.reduce(
          (total, format) => total + format.ranges.length,
          0,
        ) +
        sheet.protectedRanges.reduce(
          (total, protectedRange) => total + protectedRange.ranges.length,
          0,
        ) +
        ((sheet.print?.rowBreaks?.length ?? 0) +
          (sheet.print?.columnBreaks?.length ?? 0)) *
          2 +
        sheet.views.reduce(
          (total, view) =>
            total +
            (view.topLeftCell === undefined ? 0 : 1) +
            view.selections.reduce(
              (selectionTotal, selection) =>
                selectionTotal +
                selection.ranges.length +
                (selection.activeCell === undefined ? 0 : 1),
              0,
            ),
          0,
        ) +
        sheet.tables.reduce(
          (total, table) =>
            total +
            1 +
            (table.autoFilter === undefined
              ? 0
              : 1 +
                (table.autoFilter.sort === undefined
                  ? 0
                  : 1 + table.autoFilter.sort.conditions.length)),
          0,
        ) +
        sheet.comments.length +
        sheet.drawings.reduce(
          (total, drawing) =>
            total +
            (drawing.from === undefined ? 0 : 1) +
            (drawing.to === undefined ? 0 : 1),
          0,
        ) +
        (sheet.sparklineGroups?.reduce(
          (total, group) => total + group.sparklines.length * 2,
          0,
        ) ?? 0);
      transformStructuralLayoutReferences(sheet, operation);
      referenceUpdates += operation.kind.endsWith('-rows')
        ? transformRows(sheet, operation, readerLimits)
        : transformColumns(sheet, operation, readerLimits);
      if (referenceUpdates > writeLimits.maxReferenceUpdates) {
        writeLimitFailure(
          'maxReferenceUpdates',
          referenceUpdates,
          writeLimits.maxReferenceUpdates,
        );
      }
      impacts.push({
        kind: operation.kind,
        operationId: operation.operationId,
        range: structuralRange(operation),
        sheetKey: operation.sheetKey,
      });
      continue;
    }
    if (operation.kind === 'set-row') {
      const row = resolveRow(sheet, operation);
      if (
        operation.ifMatch !== undefined &&
        operation.ifMatch !==
          (await canonicalXlsxSha256(xlsxRowTargetState(sheet, row)))
      ) {
        operationFailure(
          'operation-precondition-failed',
          'XLSX operation precondition does not match the target row',
          operation,
        );
      }
      applyRowOperation(row, operation);
      impacts.push({
        kind: operation.kind,
        operationId: operation.operationId,
        range: String(operation.row),
        sheetKey: operation.sheetKey,
      });
      continue;
    }
    if (operation.kind === 'set-column') {
      const column = resolveColumn(sheet, operation);
      if (
        operation.ifMatch !== undefined &&
        operation.ifMatch !==
          (await canonicalXlsxSha256(xlsxColumnTargetState(sheet, column)))
      ) {
        operationFailure(
          'operation-precondition-failed',
          'XLSX operation precondition does not match the target column range',
          operation,
        );
      }
      applyColumnOperation(column, operation);
      impacts.push({
        kind: operation.kind,
        operationId: operation.operationId,
        range: `${operation.start}:${operation.end}`,
        sheetKey: operation.sheetKey,
      });
      continue;
    }
    const cellOperation = operation;
    const cell = resolveCell(sheet, cellOperation);
    if (
      cellOperation.ifMatch !== undefined &&
      cellOperation.ifMatch !==
        (await canonicalXlsxSha256(xlsxCellTargetState(sheet, cell)))
    ) {
      operationFailure(
        'operation-precondition-failed',
        'XLSX operation precondition does not match the target cell',
        cellOperation,
      );
    }
    applyCellOperation(document, sheet, cell, cellOperation);
    if (document.styles.length > readerLimits.maxStyles) {
      throw new XlsxWriteError(
        'resource-limit-exceeded',
        'XLSX edited normalized styles exceed their reader limit',
        {
          actual: document.styles.length,
          limit: readerLimits.maxStyles,
          limitName: 'maxStyles',
          operationId: cellOperation.operationId,
        },
      );
    }
    impacts.push({
      cell: cellOperation.cell,
      kind: cellOperation.kind,
      operationId: cellOperation.operationId,
      sheetKey: cellOperation.sheetKey,
    });
  }
  return {
    document,
    impacts,
    operations,
    stateHash: await canonicalXlsxSha256(document),
  };
}
