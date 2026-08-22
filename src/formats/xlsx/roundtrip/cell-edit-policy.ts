import { xlsxNumberFormatDatePrecision } from '../internal/number-format';
import type { XlsxCell } from '../types';
import { XlsxWriteError } from './errors';
import type { XlsxPackageGraph } from './internal/package-graph';
import type { XlsxCellOperationPlan } from './operation-planner';
import type { XlsxRoundTripDocument, XlsxWriteOptions } from './types';

const SAFE_PART_CONTENT_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.custom-properties+xml',
  'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  'application/vnd.openxmlformats-officedocument.theme+xml',
  'application/vnd.openxmlformats-package.core-properties+xml',
  'application/vnd.openxmlformats-package.content-types+xml',
  'application/vnd.openxmlformats-package.relationships+xml',
]);
const SAFE_RELATIONSHIP_KINDS = new Set([
  'core-properties',
  'custom-properties',
  'extended-properties',
  'hyperlink',
  'officeDocument',
  'sharedStrings',
  'styles',
  'theme',
  'worksheet',
]);
const UNSAFE_FORMULA_PATTERN =
  /(?:\[|(?:CALL|DDE|EXEC|FILTERXML|HYPERLINK|REGISTER\.ID|RTD|WEBSERVICE)\s*\()/iu;
const TABLE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
const COMMENT_CONTENT_TYPES = new Set([
  'application/vnd.ms-excel.person+xml',
  'application/vnd.ms-excel.threadedcomments+xml',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  'application/vnd.openxmlformats-officedocument.vmlDrawing',
]);
const COMMENT_RELATIONSHIP_KINDS = new Set([
  'comments',
  'person',
  'threadedComment',
  'vmlDrawing',
]);
const DRAWING_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawing+xml';
const CHART_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const DRAWING_RELATIONSHIP_KINDS = new Set(['chart', 'drawing', 'image']);

function editFailure(
  code:
    | 'formula-rewrite-unsupported'
    | 'opaque-content-conflict'
    | 'preservation-conflict'
    | 'signed-package-conflict',
  message: string,
  fields: {
    cell?: string;
    featureClass?: string;
    operationId?: string;
    part?: string;
    sheetKey?: string;
  } = {},
): never {
  throw new XlsxWriteError(code, message, fields);
}

export function xlsxCellEditRelationshipKind(type: string): string {
  return type.slice(type.lastIndexOf('/') + 1);
}

export function xlsxCellEditPartIsSafe(contentType: string): boolean {
  return SAFE_PART_CONTENT_TYPES.has(contentType);
}

export function xlsxCellEditRelationshipIsSafe(type: string): boolean {
  return SAFE_RELATIONSHIP_KINDS.has(xlsxCellEditRelationshipKind(type));
}

export function xlsxCellEditFormulaIsUnsafe(expression: string): boolean {
  return UNSAFE_FORMULA_PATTERN.test(expression);
}

export function assertXlsxSafeCellEditSource(
  graph: XlsxPackageGraph,
  options: XlsxWriteOptions,
  allowTables = false,
  allowCommentAnchors = false,
  allowDrawingAnchors = false,
): void {
  const partAllowed = (contentType: string): boolean =>
    xlsxCellEditPartIsSafe(contentType) ||
    (allowTables && contentType === TABLE_CONTENT_TYPE) ||
    (allowCommentAnchors && COMMENT_CONTENT_TYPES.has(contentType)) ||
    (allowDrawingAnchors &&
      (contentType === DRAWING_CONTENT_TYPE ||
        contentType === CHART_CONTENT_TYPE ||
        contentType.startsWith('image/')));
  const containsVml = graph.parts.some(
    (part) =>
      part.contentType ===
      'application/vnd.openxmlformats-officedocument.vmlDrawing',
  );
  const commentVmlClosure =
    containsVml && graph.parts.every((part) => partAllowed(part.contentType));
  if (graph.containsDigitalSignatures) {
    editFailure(
      'signed-package-conflict',
      'Signed XLSX packages cannot be edited without invalidating signatures',
      { featureClass: 'digital-signature' },
    );
  }
  if (graph.containsActiveContent) {
    editFailure(
      'preservation-conflict',
      'Active XLSX package content cannot enter the cell-edit closure',
      { featureClass: 'active-content' },
    );
  }
  if (graph.containsOpaqueContent && !commentVmlClosure) {
    editFailure(
      'opaque-content-conflict',
      options.acknowledgeOpaqueContent
        ? 'Opaque XLSX content has no proven independent cell-edit closure'
        : 'Opaque XLSX content requires acknowledgement and a proven independent closure',
      { featureClass: 'opaque-content' },
    );
  }
  const unknownPart = graph.parts.find(
    (part) => !partAllowed(part.contentType),
  );
  if (unknownPart) {
    editFailure(
      'opaque-content-conflict',
      'XLSX cell editing encountered an unsupported package part',
      { featureClass: 'unsupported-part', part: unknownPart.name },
    );
  }
  const unknownRelationship = graph.relationships.find(
    (relationship) =>
      !xlsxCellEditRelationshipIsSafe(relationship.type) &&
      !(
        allowTables &&
        xlsxCellEditRelationshipKind(relationship.type) === 'table'
      ) &&
      !(
        allowCommentAnchors &&
        COMMENT_RELATIONSHIP_KINDS.has(
          xlsxCellEditRelationshipKind(relationship.type),
        )
      ) &&
      !(
        allowDrawingAnchors &&
        DRAWING_RELATIONSHIP_KINDS.has(
          xlsxCellEditRelationshipKind(relationship.type),
        )
      ),
  );
  if (unknownRelationship) {
    editFailure(
      'opaque-content-conflict',
      'XLSX cell editing encountered an unsupported relationship dependency',
      {
        featureClass: 'unsupported-relationship',
        ...(unknownRelationship.owner === null
          ? {}
          : { part: unknownRelationship.owner }),
      },
    );
  }
}

function cellAt(
  document: XlsxRoundTripDocument,
  sheetKey: string,
  address: string,
): XlsxCell {
  const sheet = document.sheets.find(
    (candidate) => candidate.key === sheetKey,
  )!;
  if (sheet.kind !== 'worksheet') {
    throw new TypeError('Expected an XLSX worksheet edit target');
  }
  for (const row of sheet.rows) {
    const cell = row.cells.find((candidate) => candidate.address === address);
    if (cell) return cell;
  }
  throw new TypeError('Expected an existing XLSX cell edit target');
}

export function assertXlsxCellEditFormulaClosure(
  baseDocument: XlsxRoundTripDocument,
  plan: XlsxCellOperationPlan,
): void {
  const contentOperations = plan.operations.filter(
    (operation) =>
      operation.kind === 'clear-cell' || operation.kind === 'set-cell',
  );
  if (contentOperations.length === 0) return;
  const targets = new Set<string>();
  for (const impact of plan.impacts) {
    if (impact.kind === 'clear-cell' || impact.kind === 'set-cell') {
      targets.add(`${impact.sheetKey}\u0000${impact.cell}`);
    }
  }
  for (const sheet of baseDocument.sheets) {
    if (sheet.kind !== 'worksheet') continue;
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (
          cell.content.kind === 'formula' &&
          !targets.has(`${sheet.key}\u0000${cell.address}`)
        ) {
          editFailure(
            'formula-rewrite-unsupported',
            'XLSX cell edit dependency closure contains an unaffected formula',
            {
              cell: cell.address,
              featureClass: 'formula-dependency',
              sheetKey: sheet.key,
            },
          );
        }
      }
    }
  }
  if (baseDocument.workbook.definedNames.length !== 0) {
    editFailure(
      'formula-rewrite-unsupported',
      'XLSX cell edit dependency closure contains defined-name formulas',
      { featureClass: 'defined-name' },
    );
  }
  for (const operation of contentOperations) {
    if (
      operation.kind === 'set-cell' &&
      operation.content.kind === 'formula' &&
      xlsxCellEditFormulaIsUnsafe(operation.content.expression)
    ) {
      editFailure(
        'formula-rewrite-unsupported',
        'XLSX cell edit formula uses an external-capable function or reference',
        {
          cell: operation.cell,
          featureClass: 'external-formula',
          operationId: operation.operationId,
          sheetKey: operation.sheetKey,
        },
      );
    }
  }
}

export function assertXlsxCellEditStyleClosure(
  baseDocument: XlsxRoundTripDocument,
  plan: XlsxCellOperationPlan,
): void {
  for (const impact of plan.impacts) {
    if (!('cell' in impact)) continue;
    const cell = cellAt(plan.document, impact.sheetKey, impact.cell);
    const sourceCell = cellAt(baseDocument, impact.sheetKey, impact.cell);
    if (impact.kind === 'set-cell-style') {
      const sourceValue =
        sourceCell.content.kind === 'value'
          ? sourceCell.content.value
          : undefined;
      const numericValue =
        sourceValue?.kind === 'number'
          ? sourceValue.value
          : sourceValue?.kind === 'date' && sourceValue.source.kind === 'serial'
            ? sourceValue.source.value
            : undefined;
      if (numericValue !== undefined) {
        const targetFormat =
          cell.style === undefined
            ? undefined
            : plan.document.styles[cell.style]?.numberFormat;
        const targetPrecision =
          targetFormat === undefined
            ? undefined
            : xlsxNumberFormatDatePrecision(targetFormat, numericValue);
        const sourcePrecision =
          sourceValue?.kind === 'date' ? sourceValue.precision : undefined;
        if (sourcePrecision !== targetPrecision) {
          editFailure(
            'preservation-conflict',
            'XLSX style edit changes the cell date-value interpretation',
            {
              cell: impact.cell,
              featureClass: 'date-style-conversion',
              operationId: impact.operationId,
              sheetKey: impact.sheetKey,
            },
          );
        }
      }
      continue;
    }
    if (cell.content.kind !== 'value' || cell.content.value.kind !== 'number') {
      continue;
    }
    const format =
      sourceCell.style === undefined
        ? undefined
        : baseDocument.styles[sourceCell.style]?.numberFormat;
    if (
      format !== undefined &&
      xlsxNumberFormatDatePrecision(format, cell.content.value.value) !==
        undefined
    ) {
      editFailure(
        'preservation-conflict',
        'XLSX number edit targets a date-formatted cell',
        {
          cell: impact.cell,
          featureClass: 'date-formatted-cell',
          operationId: impact.operationId,
          sheetKey: impact.sheetKey,
        },
      );
    }
  }
}

export function xlsxPlannedCell(
  document: XlsxRoundTripDocument,
  sheetKey: string,
  address: string,
): XlsxCell {
  return cellAt(document, sheetKey, address);
}
