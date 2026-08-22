import JSZip from 'jszip';

import { XlsxParseError } from './errors';
import {
  assertXlsxArchiveWithinLimits,
  assertXlsxInputWithinLimits,
  copyXlsxInputBytes,
} from './internal/archive';
import { loadXlsxCalculationChain } from './internal/calculation-chain';
import { validateXlsxChartSheetPart } from './internal/chart-sheet';
import {
  EMPTY_XLSX_CELL_METADATA,
  loadXlsxCellMetadata,
  type XlsxCellMetadataBudget,
  type XlsxCellMetadataRegistry,
} from './internal/cell-metadata';
import {
  type XlsxCommentBudget,
  type XlsxCommentPersonTable,
  loadXlsxCommentPersons,
  loadXlsxWorksheetComments,
} from './internal/comments';
import {
  type XlsxDrawingBudget,
  XlsxMediaSession,
  loadXlsxDrawings,
} from './internal/drawing';
import { loadXlsxDocumentProperties } from './internal/document-properties';
import {
  loadXlsxExternalMetadata,
  loadXlsxQueryTables,
  type XlsxExternalMetadataLoadResult,
} from './internal/external-metadata';
import { XlsxPartReader } from './internal/part-reader';
import {
  EMPTY_XLSX_RICH_VALUES,
  hydrateXlsxRichValueImages,
  loadXlsxRichValues,
  type XlsxRichValueRegistry,
} from './internal/rich-value';
import {
  loadXlsxPivotCaches,
  loadXlsxPivotTables,
  type XlsxPivotBudget,
  type XlsxPivotCacheLoadResult,
} from './internal/pivot';
import {
  resolveXlsxResourceLimits,
  resourceLimitDiagnostic,
  XlsxResourceLimitError,
} from './internal/resource-limits';
import { resolveXlsxSelection } from './internal/selection';
import {
  activeXlsxContentDiagnostic,
  scanXlsxActiveContent,
} from './internal/security';
import {
  loadXlsxAnalyticCaches,
  loadXlsxAnalyticDisplays,
  type XlsxAnalyticCacheLoadResult,
  type XlsxSlicerBudget,
} from './internal/slicer';
import { loadXlsxStyles } from './internal/styles';
import { discoverXlsxWorkbook } from './internal/workbook-discovery';
import {
  xlsxDefinedNameFormulaCharacters,
  xlsxDefinedNameTextCharacters,
} from './internal/workbook-defined-names';
import { parseXlsxWorkbookManifest } from './internal/workbook-manifest';
import { loadXlsxSharedStrings } from './internal/workbook-tables';
import { createXlsxTableRegistry, loadXlsxTables } from './internal/table';
import { loadXlsxWorksheetRelationships } from './internal/worksheet-relationships';
import {
  createXlsxWorksheetBudget,
  parseXlsxWorksheetPart,
} from './internal/worksheet';
import type {
  XlsxComment,
  XlsxDiagnostic,
  XlsxDocument,
  XlsxDocumentProperties,
  XlsxDrawing,
  XlsxInput,
  XlsxParseOptions,
  XlsxParseResult,
} from './types';

function failResource(
  error: XlsxResourceLimitError,
  diagnostics: XlsxDiagnostic[],
): never {
  const diagnostic = resourceLimitDiagnostic(error);
  diagnostics.push(diagnostic);
  throw new XlsxParseError(diagnostic, { cause: error });
}

function recoverOptionalFeature(
  error: unknown,
  options: XlsxParseOptions,
  diagnostics: XlsxDiagnostic[],
): boolean {
  if (!(error instanceof XlsxParseError) || options.errorMode === 'strict') {
    return false;
  }
  const warning: XlsxDiagnostic = {
    ...error.diagnostic,
    severity: 'warning',
  };
  const last = diagnostics.at(-1);
  if (last === error.diagnostic) {
    diagnostics[diagnostics.length - 1] = warning;
  } else {
    diagnostics.push(warning);
  }
  return true;
}

function assertOptions(options: XlsxParseOptions): void {
  if (
    options.errorMode !== undefined &&
    options.errorMode !== 'strict' &&
    options.errorMode !== 'tolerant'
  ) {
    throw new TypeError('XLSX errorMode is invalid');
  }
  if (
    options.displayTextMode !== undefined &&
    options.displayTextMode !== 'none' &&
    options.displayTextMode !== 'supported'
  ) {
    throw new TypeError('XLSX displayTextMode is invalid');
  }
  if (
    options.imageMode !== undefined &&
    options.imageMode !== 'base64' &&
    options.imageMode !== 'blob' &&
    options.imageMode !== 'both' &&
    options.imageMode !== 'none'
  ) {
    throw new TypeError('XLSX imageMode is invalid');
  }
  if (
    options.pivotCacheMode !== undefined &&
    options.pivotCacheMode !== 'metadata' &&
    options.pivotCacheMode !== 'none' &&
    options.pivotCacheMode !== 'records'
  ) {
    throw new TypeError('XLSX pivotCacheMode is invalid');
  }
}

async function openXlsxPackage(
  input: XlsxInput,
  diagnostics: XlsxDiagnostic[],
  limits: ReturnType<typeof resolveXlsxResourceLimits>,
): Promise<JSZip> {
  try {
    assertXlsxInputWithinLimits(input, limits);
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    throw error;
  }

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(await copyXlsxInputBytes(input));
  } catch (cause) {
    const diagnostic: XlsxDiagnostic = {
      code: 'invalid-package',
      message: 'Failed to open XLSX OPC package',
      severity: 'error',
    };
    diagnostics.push(diagnostic);
    throw new XlsxParseError(diagnostic, { cause });
  }

  try {
    assertXlsxArchiveWithinLimits(zip, limits);
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    throw error;
  }
  return zip;
}

async function parseXlsxCore(
  input: XlsxInput,
  options: XlsxParseOptions,
  activeContentMode: 'diagnose' | 'reject',
): Promise<XlsxParseResult> {
  assertOptions(options);
  const diagnostics: XlsxDiagnostic[] = [];
  const limits = resolveXlsxResourceLimits(options.limits);
  const zip = await openXlsxPackage(input, diagnostics, limits);
  const reader = new XlsxPartReader(zip, diagnostics, limits);
  let discovery: Awaited<ReturnType<typeof discoverXlsxWorkbook>>;
  let manifest: Awaited<ReturnType<typeof parseXlsxWorkbookManifest>>;
  let sharedStrings: Awaited<ReturnType<typeof loadXlsxSharedStrings>>;
  let styles: Awaited<ReturnType<typeof loadXlsxStyles>>;
  let selections: ReturnType<typeof resolveXlsxSelection>;
  try {
    discovery = await discoverXlsxWorkbook(reader, limits);
    const activeContent = scanXlsxActiveContent(zip, discovery);
    if (
      activeContent.length > 0 &&
      options.errorMode === 'strict' &&
      activeContentMode === 'reject'
    ) {
      const diagnostic = activeXlsxContentDiagnostic(
        activeContent[0]!,
        'error',
      );
      diagnostics.push(diagnostic);
      throw new XlsxParseError(diagnostic);
    }
    diagnostics.push(
      ...activeContent.map((finding) =>
        activeXlsxContentDiagnostic(finding, 'warning'),
      ),
    );
    manifest = await parseXlsxWorkbookManifest(discovery, reader, limits);
    selections = resolveXlsxSelection(
      options.selection,
      manifest.sheets,
      limits,
    );
    styles = await loadXlsxStyles(discovery, reader, limits);
    sharedStrings = await loadXlsxSharedStrings(discovery, reader, limits);
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    throw error;
  }
  const budget = createXlsxWorksheetBudget(sharedStrings, {
    formulaCharacters: xlsxDefinedNameFormulaCharacters(
      manifest.properties.definedNames,
    ),
    textCharacters:
      xlsxDefinedNameTextCharacters(manifest.properties.definedNames) +
      manifest.protectionTextCharacters,
  });
  const metadataBudget: XlsxCellMetadataBudget = { records: 0 };
  let metadataRegistry: XlsxCellMetadataRegistry = EMPTY_XLSX_CELL_METADATA;
  let richValues: XlsxRichValueRegistry = EMPTY_XLSX_RICH_VALUES;
  if (
    !Number.isSafeInteger(budget.textCharacters) ||
    budget.textCharacters > limits.maxTextCharacters
  ) {
    failResource(
      new XlsxResourceLimitError(
        'maxTextCharacters',
        budget.textCharacters,
        limits.maxTextCharacters,
        discovery.part,
      ),
      diagnostics,
    );
  }
  let persons: XlsxCommentPersonTable = { byId: new Map(), values: [] };
  const pivotBudget: XlsxPivotBudget = { records: 0 };
  let pivotCacheResult: XlsxPivotCacheLoadResult = {
    caches: [],
    registry: new Map(),
  };
  const slicerBudget: XlsxSlicerBudget = { objects: 0 };
  let analyticCacheResult: XlsxAnalyticCacheLoadResult = {
    registry: new Map(),
    slicerCaches: [],
    timelineCaches: [],
  };
  let externalMetadataResult: XlsxExternalMetadataLoadResult = {
    connections: [],
    externalLinks: [],
  };
  let documentProperties: XlsxDocumentProperties | undefined;
  try {
    const chain = await loadXlsxCalculationChain(
      manifest.workbookRelationships,
      discovery,
      reader,
      manifest.sheetIdIndexes,
      limits,
    );
    if (chain.length !== 0) manifest.properties.calculation.chain = chain;
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
  }
  try {
    richValues = await loadXlsxRichValues(
      manifest.workbookRelationships,
      discovery,
      reader,
      limits,
      metadataBudget,
    );
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
  }
  try {
    metadataRegistry = await loadXlsxCellMetadata(
      manifest.workbookRelationships,
      discovery,
      reader,
      limits,
      metadataBudget,
      richValues,
    );
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
  }
  try {
    documentProperties = await loadXlsxDocumentProperties(
      discovery,
      reader,
      limits,
      budget,
    );
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
  }
  try {
    pivotCacheResult = await loadXlsxPivotCaches(
      manifest.pivotCaches,
      options.pivotCacheMode ?? 'metadata',
      discovery,
      reader,
      limits,
      pivotBudget,
      budget,
    );
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
  }
  try {
    externalMetadataResult = await loadXlsxExternalMetadata(
      manifest.workbookRelationships,
      discovery,
      reader,
      limits,
      budget,
    );
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
  }
  try {
    analyticCacheResult = await loadXlsxAnalyticCaches(
      manifest.workbookRelationships,
      discovery,
      reader,
      limits,
      budget,
    );
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
  }
  try {
    persons = await loadXlsxCommentPersons(discovery, reader, limits, budget);
  } catch (error) {
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
  }
  const sheets: XlsxDocument['sheets'] = [];
  const tableRegistry = createXlsxTableRegistry();
  const commentBudget: XlsxCommentBudget = { comments: 0 };
  const drawingBudget: XlsxDrawingBudget = { charts: 0, drawings: 0 };
  const media = new XlsxMediaSession(options.imageMode ?? 'none', limits);
  try {
    for (const [index, sheet] of manifest.sheets.entries()) {
      const selection = selections[index]!;
      if (sheet.kind === 'chart-sheet') {
        if (selection.kind === 'full-sheet') {
          await validateXlsxChartSheetPart(
            manifest.sheetParts[index]!,
            discovery.dialect,
            reader,
          );
        }
        sheets.push({
          ...sheet,
          payload:
            selection.kind === 'full-sheet' ? 'full-sheet' : 'not-selected',
        });
        continue;
      }
      if (selection.kind === 'not-selected') {
        sheets.push({ ...sheet, payload: 'not-selected', rows: [] });
        continue;
      }
      const worksheetRelationships = await loadXlsxWorksheetRelationships(
        manifest.sheetParts[index]!,
        reader,
        limits,
      );
      const legacyDrawingRelationshipIds: string[] = [];
      const drawingRelationshipIds: string[] = [];
      const tableRelationshipIds: string[] = [];
      const pivotTableRelationshipIds: string[] = [];
      const parsedPayload = await parseXlsxWorksheetPart(
        manifest.sheetParts[index]!,
        discovery.dialect,
        reader,
        limits,
        sharedStrings,
        budget,
        selection,
        {
          dateSystem: manifest.properties.dateSystem,
          dialect: discovery.dialect,
          drawingRelationshipIds,
          legacyDrawingRelationshipIds,
          metadataBudget,
          metadataRegistry,
          pivotTableRelationshipIds,
          relationships: worksheetRelationships,
          styles,
          tableRelationshipIds,
          workbookViewCount: manifest.properties.views.length,
        },
      );
      const { unsupportedExtensions, unsupportedMetadata, ...payload } =
        parsedPayload;
      if (unsupportedExtensions) {
        const diagnostic: XlsxDiagnostic = {
          code: 'unsupported-feature',
          message: 'Worksheet extension content was omitted',
          part: manifest.sheetParts[index]!,
          severity:
            options.errorMode === 'strict' && activeContentMode === 'reject'
              ? 'error'
              : 'warning',
          sheet: sheet.name,
        };
        diagnostics.push(diagnostic);
        if (diagnostic.severity === 'error') {
          throw new XlsxParseError(diagnostic);
        }
      }
      if (unsupportedMetadata) {
        const diagnostic: XlsxDiagnostic = {
          code: 'unsupported-feature',
          message: 'Worksheet modern metadata content was omitted',
          part: manifest.sheetParts[index]!,
          severity:
            options.errorMode === 'strict' && activeContentMode === 'reject'
              ? 'error'
              : 'warning',
          sheet: sheet.name,
        };
        diagnostics.push(diagnostic);
        if (diagnostic.severity === 'error') {
          throw new XlsxParseError(diagnostic);
        }
      }
      let drawings: XlsxDrawing[] = [];
      const mediaCheckpoint = media.checkpoint();
      try {
        drawings = await loadXlsxDrawings(
          drawingRelationshipIds,
          worksheetRelationships,
          discovery,
          reader,
          limits,
          drawingBudget,
          budget,
          selection,
          media,
          manifest.sheetParts[index]!,
        );
      } catch (error) {
        media.rollback(mediaCheckpoint);
        if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
      }
      let comments: XlsxComment[] = [];
      try {
        comments = await loadXlsxWorksheetComments(
          manifest.sheetParts[index]!,
          legacyDrawingRelationshipIds,
          worksheetRelationships,
          discovery,
          reader,
          limits,
          commentBudget,
          budget,
          selection,
          persons,
        );
      } catch (error) {
        if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
      }
      const tables = await loadXlsxTables(
        tableRelationshipIds,
        worksheetRelationships,
        discovery,
        reader,
        limits,
        tableRegistry,
        styles.differentialStyles.length,
        budget,
        selection,
        manifest.sheetParts[index]!,
      );
      let pivotTables: Awaited<ReturnType<typeof loadXlsxPivotTables>> = [];
      try {
        pivotTables = await loadXlsxPivotTables(
          pivotTableRelationshipIds,
          worksheetRelationships,
          pivotCacheResult.registry,
          discovery,
          reader,
          limits,
          budget,
          selection,
          manifest.sheetParts[index]!,
        );
      } catch (error) {
        if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
      }
      let analyticDisplays: Awaited<
        ReturnType<typeof loadXlsxAnalyticDisplays>
      > = { slicers: [], timelines: [] };
      try {
        analyticDisplays = await loadXlsxAnalyticDisplays(
          worksheetRelationships,
          analyticCacheResult.registry,
          discovery,
          reader,
          limits,
          budget,
          slicerBudget,
          selection,
          manifest.sheetParts[index]!,
        );
      } catch (error) {
        if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
      }
      let queryTables: Awaited<ReturnType<typeof loadXlsxQueryTables>> = [];
      try {
        queryTables = await loadXlsxQueryTables(
          worksheetRelationships,
          new Set(
            externalMetadataResult.connections.map(
              (connection) => connection.id,
            ),
          ),
          discovery,
          reader,
          limits,
          budget,
          manifest.sheetParts[index]!,
        );
      } catch (error) {
        if (!recoverOptionalFeature(error, options, diagnostics)) throw error;
      }
      sheets.push({
        ...sheet,
        ...payload,
        comments,
        drawings,
        payload:
          selection.kind === 'full-sheet' ? 'full-sheet' : 'selected-ranges',
        ...(pivotTables.length === 0 ? {} : { pivotTables }),
        ...(queryTables.length === 0 ? {} : { queryTables }),
        ...(analyticDisplays.slicers.length === 0
          ? {}
          : { slicers: analyticDisplays.slicers }),
        tables,
        ...(analyticDisplays.timelines.length === 0
          ? {}
          : { timelines: analyticDisplays.timelines }),
      });
    }
    await hydrateXlsxRichValueImages(sheets, media, reader);
  } catch (error) {
    media.revokeAll();
    if (error instanceof XlsxResourceLimitError) {
      failResource(error, diagnostics);
    }
    throw error;
  }
  const document: XlsxDocument = {
    ...(externalMetadataResult.connections.length === 0
      ? {}
      : { connections: externalMetadataResult.connections }),
    differentialStyles: [...styles.differentialStyles],
    ...(documentProperties === undefined ? {} : { documentProperties }),
    ...(externalMetadataResult.externalLinks.length === 0
      ? {}
      : { externalLinks: externalMetadataResult.externalLinks }),
    namedStyles: [...styles.namedStyles],
    ...(pivotCacheResult.caches.length === 0
      ? {}
      : { pivotCaches: pivotCacheResult.caches }),
    sheets,
    ...(analyticCacheResult.slicerCaches.length === 0
      ? {}
      : { slicerCaches: analyticCacheResult.slicerCaches }),
    styles: [...styles.styles],
    ...(analyticCacheResult.timelineCaches.length === 0
      ? {}
      : { timelineCaches: analyticCacheResult.timelineCaches }),
    workbook: {
      ...manifest.properties,
      ...(persons.values.length === 0
        ? {}
        : { commentPersons: persons.values }),
    },
  };
  return { diagnostics, document };
}

export async function parseXlsxWithDiagnostics(
  input: XlsxInput,
  options: XlsxParseOptions = {},
): Promise<XlsxParseResult> {
  return parseXlsxCore(input, options, 'reject');
}

export async function parseXlsxPreservingActiveContent(
  input: XlsxInput,
  options: XlsxParseOptions,
): Promise<XlsxDocument> {
  return (await parseXlsxCore(input, options, 'diagnose')).document;
}

export async function parseXlsx(
  input: XlsxInput,
  options: XlsxParseOptions = {},
): Promise<XlsxDocument> {
  return (await parseXlsxWithDiagnostics(input, options)).document;
}
