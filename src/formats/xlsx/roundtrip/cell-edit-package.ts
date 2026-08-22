import { readZipEntryBytes } from '../../../common/archive/read-entry';
import JSZip from 'jszip';

import { XlsxPartReader } from '../internal/part-reader';
import { getXlsxRelationshipPartName } from '../internal/package-identity';
import type { ResolvedXlsxResourceLimits } from '../internal/resource-limits';
import { discoverXlsxWorkbook } from '../internal/workbook-discovery';
import { parseXlsxWorkbookManifest } from '../internal/workbook-manifest';
import { loadXlsxStyles, type XlsxStyleTable } from '../internal/styles';
import type { XlsxWorksheet } from '../types';
import { XlsxWriteError } from './errors';
import {
  inspectXlsxPackageGraph,
  type XlsxPackageGraph,
} from './internal/package-graph';
import {
  assertXlsxCellEditFormulaClosure,
  assertXlsxCellEditStyleClosure,
  assertXlsxSafeCellEditSource,
  xlsxPlannedCell,
} from './cell-edit-policy';
import {
  generateBoundedXlsxZip,
  verifyXlsxCellEditR1Parts,
} from './cell-edit-verification';
import type { XlsxCellOperationPlan } from './operation-planner';
import type {
  ResolvedXlsxWriteLimits,
  XlsxPartFidelity,
  XlsxRoundTripDocument,
  XlsxWriteOptions,
} from './types';
import {
  appendXlsxStylesPart,
  xlsxAppendedStyleRecordCount,
} from './style-append';
import {
  patchXlsxHyperlinks,
  type XlsxHyperlinkPatch,
} from './hyperlink-patch';
import {
  patchXlsxHyperlinkRelationships,
  planXlsxExternalHyperlinkRelationships,
} from './hyperlink-relationships';
import {
  patchXlsxWorksheetPartWithReport,
  type XlsxWorksheetCellPatch,
} from './worksheet-patch';
import {
  patchXlsxWorksheetProperties,
  type XlsxWorksheetPropertyPatch,
} from './worksheet-properties-patch';
import { writeLimitFailure } from './write-limits';
import {
  patchXlsxWorksheetStructure,
  type XlsxWorksheetStructurePatch,
} from './worksheet-structure-patch';
import { patchXlsxTableStructure } from './table-structure-patch';
import {
  patchXlsxCommentAnchors,
  patchXlsxCommentVmlAnchors,
} from './comment-structure-patch';
import { patchXlsxDrawingStructure } from './drawing-structure-patch';
import { patchXlsxChartStructure } from './chart-structure-patch';
import { patchXlsxSparklineStructure } from './sparkline-structure-patch';

export interface XlsxCellEditPackage {
  data: Uint8Array;
  graph: XlsxPackageGraph;
  parts: XlsxPartFidelity[];
}

export function xlsxStructuralRelationshipTargets(
  graph: Pick<XlsxPackageGraph, 'relationships'>,
  owner: string,
  type: string,
): string[] {
  return [
    ...new Set(
      graph.relationships
        .filter(
          (relationship) =>
            relationship.owner === owner &&
            relationship.mode === 'internal' &&
            relationship.type === type,
        )
        .map((relationship) => relationship.target),
    ),
  ];
}

const TABLE_STRUCTURAL_OPERATION_KINDS = new Set([
  'delete-columns',
  'delete-rows',
  'insert-columns',
  'insert-rows',
]);

async function packageContext(
  bytes: Uint8Array,
  readerLimits: ResolvedXlsxResourceLimits,
): Promise<{ sheetParts: string[]; styles: XlsxStyleTable }> {
  const archive = await JSZip.loadAsync(bytes);
  const reader = new XlsxPartReader(archive, [], readerLimits);
  const discovery = await discoverXlsxWorkbook(reader, readerLimits);
  const manifest = await parseXlsxWorkbookManifest(
    discovery,
    reader,
    readerLimits,
  );
  return {
    sheetParts: manifest.sheetParts,
    styles: await loadXlsxStyles(discovery, reader, readerLimits),
  };
}

function finalPatches(
  plan: XlsxCellOperationPlan,
  styles: XlsxStyleTable,
  appendedStyleXfs: ReadonlyMap<number, number>,
): Map<string, XlsxWorksheetCellPatch[]> {
  const bySheet = new Map<string, Map<string, XlsxWorksheetCellPatch>>();
  const styleCells = new Set<string>();
  const contentCells = new Set<string>();
  for (const impact of plan.impacts) {
    if (impact.kind === 'set-cell-style') {
      styleCells.add(`${impact.sheetKey}\u0000${impact.cell}`);
    } else if (impact.kind === 'clear-cell' || impact.kind === 'set-cell') {
      contentCells.add(`${impact.sheetKey}\u0000${impact.cell}`);
    }
  }
  for (const impact of plan.impacts) {
    if (!('cell' in impact) || impact.kind === 'set-hyperlink') continue;
    let sheet = bySheet.get(impact.sheetKey);
    if (!sheet) {
      sheet = new Map();
      bySheet.set(impact.sheetKey, sheet);
    }
    const cell = xlsxPlannedCell(plan.document, impact.sheetKey, impact.cell);
    const styleCell = styleCells.has(`${impact.sheetKey}\u0000${impact.cell}`);
    const xmlStyleIndex = styleCell
      ? cell.style !== undefined && appendedStyleXfs.has(cell.style)
        ? appendedStyleXfs.get(cell.style)
        : styles.cellXfs.findIndex(
            (candidate) => candidate.normalizedStyle === cell.style,
          )
      : undefined;
    sheet.set(impact.cell, {
      cell,
      contentChanged: contentCells.has(
        `${impact.sheetKey}\u0000${impact.cell}`,
      ),
      operationId: impact.operationId,
      ...(xmlStyleIndex === undefined ? {} : { xmlStyleIndex }),
    });
  }
  return new Map(
    [...bySheet].map(([sheetKey, patches]) => [
      sheetKey,
      [...patches.values()],
    ]),
  );
}

function finalHyperlinkPatches(
  plan: XlsxCellOperationPlan,
): Map<string, XlsxHyperlinkPatch[]> {
  const bySheet = new Map<string, Map<string, XlsxHyperlinkPatch>>();
  for (const operation of plan.operations) {
    if (operation.kind !== 'set-hyperlink') continue;
    let sheet = bySheet.get(operation.sheetKey);
    if (!sheet) {
      sheet = new Map();
      bySheet.set(operation.sheetKey, sheet);
    }
    sheet.set(operation.cell, {
      cell: operation.cell,
      operationId: operation.operationId,
      target: operation.target,
    });
  }
  return new Map(
    [...bySheet].map(([sheetKey, patches]) => [
      sheetKey,
      [...patches.values()],
    ]),
  );
}

function finalPropertyPatches(
  plan: XlsxCellOperationPlan,
): Map<string, XlsxWorksheetPropertyPatch[]> {
  const bySheet = new Map<string, Map<string, XlsxWorksheetPropertyPatch>>();
  for (const operation of plan.operations) {
    if (operation.kind !== 'set-column' && operation.kind !== 'set-row') {
      continue;
    }
    let sheet = bySheet.get(operation.sheetKey);
    if (!sheet) {
      sheet = new Map();
      bySheet.set(operation.sheetKey, sheet);
    }
    const key =
      operation.kind === 'set-row'
        ? `row:${operation.row}`
        : `column:${operation.start}:${operation.end}`;
    const previous = sheet.get(key);
    if (operation.kind === 'set-row') {
      const next: Extract<XlsxWorksheetPropertyPatch, { kind: 'set-row' }> = {
        ...(previous as
          Extract<XlsxWorksheetPropertyPatch, { kind: 'set-row' }> | undefined),
        kind: operation.kind,
        operationId: operation.operationId,
        row: operation.row,
      };
      if (operation.height !== undefined) next.height = operation.height;
      if (operation.hidden !== undefined) next.hidden = operation.hidden;
      sheet.set(key, next);
    } else {
      const next: Extract<XlsxWorksheetPropertyPatch, { kind: 'set-column' }> =
        {
          ...(previous as
            | Extract<XlsxWorksheetPropertyPatch, { kind: 'set-column' }>
            | undefined),
          end: operation.end,
          kind: operation.kind,
          operationId: operation.operationId,
          start: operation.start,
        };
      if (operation.hidden !== undefined) next.hidden = operation.hidden;
      if (operation.width !== undefined) next.width = operation.width;
      sheet.set(key, next);
    }
  }
  return new Map(
    [...bySheet].map(([sheetKey, patches]) => [
      sheetKey,
      [...patches.values()],
    ]),
  );
}

function finalStructuralPatches(
  plan: XlsxCellOperationPlan,
): Map<string, XlsxWorksheetStructurePatch[]> {
  const bySheet = new Map<string, XlsxWorksheetStructurePatch[]>();
  for (const operation of plan.operations) {
    if (
      operation.kind !== 'delete-columns' &&
      operation.kind !== 'delete-rows' &&
      operation.kind !== 'insert-columns' &&
      operation.kind !== 'insert-rows'
    ) {
      continue;
    }
    const patches = bySheet.get(operation.sheetKey) ?? [];
    patches.push({
      count: operation.count,
      index: operation.index,
      kind: operation.kind,
      operationId: operation.operationId,
    });
    bySheet.set(operation.sheetKey, patches);
  }
  return bySheet;
}

export async function writeXlsxCellEditPackage(
  sourceBytes: Uint8Array,
  sourceGraph: XlsxPackageGraph,
  baseDocument: XlsxRoundTripDocument,
  plan: XlsxCellOperationPlan,
  options: XlsxWriteOptions,
  writeLimits: ResolvedXlsxWriteLimits,
  readerLimits: ResolvedXlsxResourceLimits,
): Promise<XlsxCellEditPackage> {
  const structuralClosure = TABLE_STRUCTURAL_OPERATION_KINDS.has(
    plan.operations[0]!.kind,
  );
  assertXlsxSafeCellEditSource(
    sourceGraph,
    options,
    structuralClosure,
    structuralClosure,
    structuralClosure,
  );
  assertXlsxCellEditFormulaClosure(baseDocument, plan);
  assertXlsxCellEditStyleClosure(baseDocument, plan);
  const context = await packageContext(sourceBytes, readerLimits);
  const appendedStyles = plan.document.styles.slice(baseDocument.styles.length);
  const outputStyleRecords =
    context.styles.recordCount + xlsxAppendedStyleRecordCount(appendedStyles);
  if (outputStyleRecords > readerLimits.maxStyles) {
    throw new XlsxWriteError(
      'resource-limit-exceeded',
      'XLSX appended style records exceed the reader limit',
      {
        actual: outputStyleRecords,
        limit: readerLimits.maxStyles,
        limitName: 'maxStyles',
      },
    );
  }
  if (appendedStyles.length !== 0 && context.styles.part === null) {
    throw new XlsxWriteError(
      'preservation-conflict',
      'XLSX cannot append styles without an existing styles part',
      { featureClass: 'missing-styles-part' },
    );
  }
  const archive = await JSZip.loadAsync(sourceBytes);
  const appendedStyleXfs = new Map<number, number>();
  const dirtyParts = new Set<string>();
  const addedParts = new Set<string>();
  const changedRelationshipOwners = new Set<string>();
  let generatedXmlBytes = 0;
  let patchBytes = 0;
  let patchCount = 0;
  let tableDependencyEdges = 0;
  let commentDependencyEdges = 0;
  let drawingDependencyEdges = 0;
  let chartDependencyEdges = 0;
  const visitedChartParts = new Set<string>();
  if (appendedStyles.length !== 0) {
    const part = context.styles.part!;
    const entry = archive.file(part)!;
    const source = await readZipEntryBytes(entry, readerLimits.maxPartBytes);
    const appended = appendXlsxStylesPart(
      source,
      appendedStyles,
      writeLimits,
      part,
    );
    generatedXmlBytes += appended.data.byteLength;
    patchBytes += appended.patchBytes;
    patchCount += appended.patchCount;
    archive.file(part, appended.data, { date: entry.date });
    dirtyParts.add(part);
    for (const [offset, xmlStyleIndex] of appended.cellXfIndexes.entries()) {
      appendedStyleXfs.set(baseDocument.styles.length + offset, xmlStyleIndex);
    }
  }
  const patches = finalPatches(plan, context.styles, appendedStyleXfs);
  const hyperlinkPatches = finalHyperlinkPatches(plan);
  const propertyPatches = finalPropertyPatches(plan);
  const structuralPatches = finalStructuralPatches(plan);
  const sheetKeys = new Set([
    ...patches.keys(),
    ...hyperlinkPatches.keys(),
    ...propertyPatches.keys(),
    ...structuralPatches.keys(),
  ]);
  for (const sheetKey of sheetKeys) {
    const sheetPatches = patches.get(sheetKey) ?? [];
    const sheet = baseDocument.sheets.find(
      (candidate) => candidate.key === sheetKey,
    )!;
    const part = context.sheetParts[sheet.index]!;
    const entry = archive.file(part)!;
    const source = await readZipEntryBytes(entry, readerLimits.maxPartBytes);
    const structuralPatched = patchXlsxWorksheetStructure(
      source,
      structuralPatches.get(sheetKey) ?? [],
      writeLimits,
      part,
    );
    const sparklinePatched = patchXlsxSparklineStructure(
      structuralPatched.data,
      structuralPatches.get(sheetKey) ?? [],
      writeLimits,
      part,
      sheet.name,
    );
    const cellPatched = patchXlsxWorksheetPartWithReport(
      sparklinePatched.data,
      sheetPatches,
      writeLimits,
      part,
    );
    const propertyPatched = patchXlsxWorksheetProperties(
      cellPatched.data,
      propertyPatches.get(sheetKey) ?? [],
      writeLimits,
      part,
    );
    const finalSheet = plan.document.sheets.find(
      (candidate) => candidate.key === sheetKey,
    ) as XlsxWorksheet & { key: string };
    const relationshipPlan = planXlsxExternalHyperlinkRelationships(
      propertyPatched.data,
      sourceGraph.relationships.filter(
        (relationship) => relationship.owner === part,
      ),
      finalSheet.hyperlinks,
      part,
    );
    const requestedHyperlinks = (
      hyperlinkPatches.get(sheetKey) ?? []
    ).map<XlsxHyperlinkPatch>((request) => ({
      ...request,
      relationshipId: relationshipPlan.idsByCell.get(request.cell),
    }));
    const officeRelationshipNamespace =
      sourceGraph.conformance === 'strict'
        ? 'http://purl.oclc.org/ooxml/officeDocument/relationships'
        : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const hyperlinkPatched = patchXlsxHyperlinks(
      propertyPatched.data,
      requestedHyperlinks,
      writeLimits,
      part,
      officeRelationshipNamespace,
    );
    generatedXmlBytes += hyperlinkPatched.data.byteLength;
    if (generatedXmlBytes > writeLimits.maxGeneratedXmlBytes) {
      writeLimitFailure(
        'maxGeneratedXmlBytes',
        generatedXmlBytes,
        writeLimits.maxGeneratedXmlBytes,
        part,
      );
    }
    patchBytes +=
      structuralPatched.patchBytes +
      sparklinePatched.patchBytes +
      cellPatched.patchBytes +
      propertyPatched.patchBytes +
      hyperlinkPatched.patchBytes;
    if (patchBytes > writeLimits.maxPatchBytes) {
      writeLimitFailure(
        'maxPatchBytes',
        patchBytes,
        writeLimits.maxPatchBytes,
        part,
      );
    }
    patchCount +=
      structuralPatched.patchCount +
      sparklinePatched.patchCount +
      cellPatched.patchCount +
      propertyPatched.patchCount +
      hyperlinkPatched.patchCount;
    if (patchCount > writeLimits.maxPatchCount) {
      writeLimitFailure(
        'maxPatchCount',
        patchCount,
        writeLimits.maxPatchCount,
        part,
      );
    }
    archive.file(part, hyperlinkPatched.data, { date: entry.date });
    dirtyParts.add(part);
    if (relationshipPlan.changed) {
      const relationshipPart = getXlsxRelationshipPartName(part);
      const relationshipEntry = archive.file(relationshipPart);
      const relationshipBytes = relationshipEntry
        ? await readZipEntryBytes(relationshipEntry, readerLimits.maxPartBytes)
        : null;
      const relationshipPatched = patchXlsxHyperlinkRelationships(
        relationshipBytes,
        relationshipPlan,
        `${officeRelationshipNamespace}/hyperlink`,
        writeLimits,
        readerLimits,
        relationshipPart,
      );
      generatedXmlBytes += relationshipPatched.data.byteLength;
      patchBytes += relationshipPatched.patchBytes;
      patchCount += relationshipPatched.patchCount;
      archive.file(relationshipPart, relationshipPatched.data, {
        date: relationshipEntry?.date ?? entry.date,
      });
      if (relationshipEntry) dirtyParts.add(relationshipPart);
      else addedParts.add(relationshipPart);
      changedRelationshipOwners.add(part);
    }
    const requestedStructure = structuralPatches.get(sheetKey) ?? [];
    const tableRelationshipType = `${officeRelationshipNamespace}/table`;
    const tableParts = sourceGraph.relationships
      .filter(
        (relationship) =>
          relationship.owner === part &&
          relationship.mode === 'internal' &&
          relationship.type === tableRelationshipType,
      )
      .map((relationship) => relationship.target);
    for (const tablePart of tableParts) {
      const tableEntry = archive.file(tablePart)!;
      const tableSource = await readZipEntryBytes(
        tableEntry,
        readerLimits.maxPartBytes,
      );
      const tablePatched = patchXlsxTableStructure(
        tableSource,
        requestedStructure,
        writeLimits,
        tablePart,
      );
      if (tablePatched.patchCount === 0) continue;
      generatedXmlBytes += tablePatched.data.byteLength;
      patchBytes += tablePatched.patchBytes;
      patchCount += tablePatched.patchCount;
      tableDependencyEdges += 1;
      archive.file(tablePart, tablePatched.data, { date: tableEntry.date });
      dirtyParts.add(tablePart);
    }
    const commentRelationshipKinds = new Map<string, 'comments' | 'vml'>([
      [`${officeRelationshipNamespace}/comments`, 'comments'],
      [`${officeRelationshipNamespace}/vmlDrawing`, 'vml'],
      [
        'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
        'comments',
      ],
    ] as const);
    const commentParts = sourceGraph.relationships.filter(
      (relationship) =>
        relationship.owner === part &&
        commentRelationshipKinds.has(relationship.type),
    );
    for (const relationship of commentParts) {
      const commentPart = relationship.target;
      const commentEntry = archive.file(commentPart)!;
      const commentSource = await readZipEntryBytes(
        commentEntry,
        readerLimits.maxPartBytes,
      );
      const commentPatched =
        commentRelationshipKinds.get(relationship.type) === 'vml'
          ? patchXlsxCommentVmlAnchors(
              commentSource,
              requestedStructure,
              writeLimits,
              commentPart,
            )
          : patchXlsxCommentAnchors(
              commentSource,
              requestedStructure,
              writeLimits,
              commentPart,
            );
      if (commentPatched.patchCount === 0) continue;
      generatedXmlBytes += commentPatched.data.byteLength;
      patchBytes += commentPatched.patchBytes;
      patchCount += commentPatched.patchCount;
      commentDependencyEdges += 1;
      archive.file(commentPart, commentPatched.data, {
        date: commentEntry.date,
      });
      dirtyParts.add(commentPart);
    }
    const drawingRelationshipType = `${officeRelationshipNamespace}/drawing`;
    const drawingParts = xlsxStructuralRelationshipTargets(
      sourceGraph,
      part,
      drawingRelationshipType,
    );
    for (const drawingPart of drawingParts) {
      const drawingEntry = archive.file(drawingPart)!;
      const drawingSource = await readZipEntryBytes(
        drawingEntry,
        readerLimits.maxPartBytes,
      );
      const drawingPatched = patchXlsxDrawingStructure(
        drawingSource,
        requestedStructure,
        writeLimits,
        drawingPart,
      );
      let drawingEdgeCharged = false;
      if (drawingPatched.patchCount !== 0) {
        generatedXmlBytes += drawingPatched.data.byteLength;
        patchBytes += drawingPatched.patchBytes;
        patchCount += drawingPatched.patchCount;
        drawingDependencyEdges += 1;
        drawingEdgeCharged = true;
        archive.file(drawingPart, drawingPatched.data, {
          date: drawingEntry.date,
        });
        dirtyParts.add(drawingPart);
      }
      const chartRelationshipType = `${officeRelationshipNamespace}/chart`;
      const chartParts = xlsxStructuralRelationshipTargets(
        sourceGraph,
        drawingPart,
        chartRelationshipType,
      );
      for (const chartPart of chartParts) {
        if (visitedChartParts.has(chartPart)) continue;
        visitedChartParts.add(chartPart);
        const chartEntry = archive.file(chartPart)!;
        const chartSource = await readZipEntryBytes(
          chartEntry,
          readerLimits.maxPartBytes,
        );
        const chartPatched = patchXlsxChartStructure(
          chartSource,
          requestedStructure,
          writeLimits,
          chartPart,
          sheet.name,
        );
        if (chartPatched.patchCount === 0) continue;
        generatedXmlBytes += chartPatched.data.byteLength;
        patchBytes += chartPatched.patchBytes;
        patchCount += chartPatched.patchCount;
        if (!drawingEdgeCharged) {
          drawingDependencyEdges += 1;
          drawingEdgeCharged = true;
        }
        chartDependencyEdges += 1;
        archive.file(chartPart, chartPatched.data, { date: chartEntry.date });
        dirtyParts.add(chartPart);
      }
    }
  }
  const dirtyPartCount = dirtyParts.size + addedParts.size;
  if (dirtyPartCount > writeLimits.maxDirtyParts) {
    writeLimitFailure(
      'maxDirtyParts',
      dirtyPartCount,
      writeLimits.maxDirtyParts,
    );
  }
  if (dirtyParts.size > writeLimits.maxPatchedParts) {
    writeLimitFailure(
      'maxPatchedParts',
      dirtyParts.size,
      writeLimits.maxPatchedParts,
    );
  }
  let dependencyEdges =
    plan.impacts.length +
    appendedStyles.length +
    changedRelationshipOwners.size;
  dependencyEdges += tableDependencyEdges;
  dependencyEdges += commentDependencyEdges;
  dependencyEdges += drawingDependencyEdges;
  dependencyEdges += chartDependencyEdges;
  if (dependencyEdges > writeLimits.maxDependencyEdges) {
    writeLimitFailure(
      'maxDependencyEdges',
      dependencyEdges,
      writeLimits.maxDependencyEdges,
    );
  }
  if (generatedXmlBytes > writeLimits.maxGeneratedXmlBytes) {
    writeLimitFailure(
      'maxGeneratedXmlBytes',
      generatedXmlBytes,
      writeLimits.maxGeneratedXmlBytes,
    );
  }
  if (patchBytes > writeLimits.maxPatchBytes) {
    writeLimitFailure('maxPatchBytes', patchBytes, writeLimits.maxPatchBytes);
  }
  if (patchCount > writeLimits.maxPatchCount) {
    writeLimitFailure('maxPatchCount', patchCount, writeLimits.maxPatchCount);
  }
  const data = await generateBoundedXlsxZip(
    archive,
    writeLimits.maxOutputBytes,
  );
  const graph = await inspectXlsxPackageGraph(data, readerLimits);
  const fidelityParts = verifyXlsxCellEditR1Parts(
    sourceGraph,
    graph,
    dirtyParts,
    addedParts,
    changedRelationshipOwners,
  );
  return { data, graph, parts: fidelityParts };
}
