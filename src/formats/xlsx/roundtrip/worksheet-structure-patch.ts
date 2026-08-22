import {
  parseXlsxCellReference,
  parseXlsxRangeReference,
  xlsxColumnName,
} from '../internal/cell-reference';
import { XLSX_MAX_COLUMNS, XLSX_MAX_ROWS } from '../internal/resource-limits';
import type { XlsxPanePosition } from '../types';
import { XlsxWriteError } from './errors';
import { xlsxMatchingCloseToken } from './hyperlink-patch';
import {
  transformXlsxStructuralPageBreak,
  transformXlsxStructuralRange,
  transformXlsxStructuralViewSelection,
  transformXlsxStructuralVisualCell,
} from './structural-reference';
import type { ResolvedXlsxWriteLimits } from './types';
import {
  decodeXlsxXml,
  encodeXlsxXml,
  tokenizeXlsxXml,
  xlsxXmlLocalName,
  type XlsxXmlAttributeSpan,
  type XlsxXmlTagToken,
} from './worksheet-patch';
import { writeLimitFailure } from './write-limits';

export interface XlsxWorksheetStructurePatch {
  count: number;
  index: number;
  kind: 'delete-columns' | 'delete-rows' | 'insert-columns' | 'insert-rows';
  operationId: string;
}

export interface XlsxWorksheetStructurePatchResult {
  data: Uint8Array;
  patchBytes: number;
  patchCount: number;
}

interface TextPatch {
  end: number;
  replacement: string;
  start: number;
}

function failure(
  message: string,
  part: string,
  request?: XlsxWorksheetStructurePatch,
  featureClass = 'worksheet-structure-xml',
): never {
  throw new XlsxWriteError('preservation-conflict', message, {
    featureClass,
    ...(request === undefined
      ? {}
      : {
          operationId: request.operationId,
          range: `${request.index}:${request.index + request.count - 1}`,
        }),
    part,
  });
}

function attribute(
  token: XlsxXmlTagToken,
  name: string,
): XlsxXmlAttributeSpan | undefined {
  return token.attributes.find((candidate) => candidate.name === name);
}

function attributePatch(span: XlsxXmlAttributeSpan, value: string): TextPatch {
  return {
    end: span.end,
    replacement: ` ${span.name}="${value}"`,
    start: span.start,
  };
}

function directRows(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  part: string,
): { rows: XlsxXmlTagToken[]; sheetData: XlsxXmlTagToken } {
  const prefix = root.name.slice(0, -'worksheet'.length);
  const sheetData = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}sheetData`,
  );
  if (!sheetData) {
    failure('XLSX worksheet sheetData cannot patch structure', part);
  }
  const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(sheetData));
  return {
    rows: tokens.filter(
      (token) =>
        !token.closing &&
        token.depth === sheetData.depth + 1 &&
        token.name === `${prefix}row` &&
        token.start >= sheetData.end &&
        token.end <= close.start,
    ),
    sheetData,
  };
}

function shiftedIndex(
  value: number,
  request: XlsxWorksheetStructurePatch,
): number | null {
  const end = request.index + request.count - 1;
  if (request.kind.startsWith('insert-')) {
    if (value < request.index) return value;
    return value + request.count;
  }
  if (value < request.index) return value;
  if (value <= end) return null;
  return value - request.count;
}

function pageBreakPatches(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  prefix: string,
  request: XlsxWorksheetStructurePatch,
  part: string,
  axis: 'column' | 'row',
): TextPatch[] {
  const name = axis === 'row' ? 'rowBreaks' : 'colBreaks';
  const container = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}${name}`,
  );
  if (!container) return [];
  const containerIndex = tokens.indexOf(container);
  const close = xlsxMatchingCloseToken(tokens, containerIndex);
  const entries = tokens
    .slice(containerIndex, tokens.indexOf(close))
    .filter(
      (token) =>
        !token.closing &&
        token.depth === container.depth + 1 &&
        token.name === `${prefix}brk`,
    );
  const unsigned = (
    entry: XlsxXmlTagToken,
    attributeName: string,
    fallback?: number,
  ): number => {
    const source = attribute(entry, attributeName)?.value;
    if (source === undefined) {
      if (fallback !== undefined) return fallback;
      failure('XLSX structural page-break value is invalid', part, request);
    }
    if (!/^(?:0|[1-9]\d*)$/u.test(source)) {
      failure('XLSX structural page-break value is invalid', part, request);
    }
    return Number(source);
  };
  const transformedEntries = entries.map((entry) => {
    const manualSource = attribute(entry, 'man')?.value;
    const pivotSource = attribute(entry, 'pt')?.value;
    const flag = (source: string | undefined): boolean => {
      if (source === undefined || source === '0' || source === 'false') {
        return false;
      }
      if (source === '1' || source === 'true') return true;
      failure('XLSX structural page-break flag is invalid', part, request);
    };
    const extentLimit = axis === 'row' ? XLSX_MAX_COLUMNS : XLSX_MAX_ROWS;
    const pageBreak = {
      end: unsigned(entry, 'max', extentLimit - 1),
      manual: flag(manualSource),
      pivot: flag(pivotSource),
      position: unsigned(entry, 'id'),
      start: unsigned(entry, 'min', 0),
    };
    return {
      entry,
      pageBreak,
      transformed: transformXlsxStructuralPageBreak(pageBreak, axis, request),
    };
  });
  const remaining = transformedEntries.filter(
    (entry) => entry.transformed !== null,
  );
  if (entries.length !== 0 && remaining.length === 0) {
    return [{ end: close.end, replacement: '', start: container.start }];
  }
  const patches: TextPatch[] = [];
  for (const item of transformedEntries) {
    if (item.transformed === null) {
      const itemClose = xlsxMatchingCloseToken(
        tokens,
        tokens.indexOf(item.entry),
      );
      patches.push({
        end: itemClose.end,
        replacement: '',
        start: item.entry.start,
      });
      continue;
    }
    let missing: string | undefined;
    for (const [attributeName, previous, next] of [
      ['id', item.pageBreak.position, item.transformed.position],
      ['min', item.pageBreak.start, item.transformed.start],
      ['max', item.pageBreak.end, item.transformed.end],
    ] as const) {
      if (previous === next) continue;
      const source = attribute(item.entry, attributeName);
      if (source) patches.push(attributePatch(source, String(next)));
      else missing = `${attributeName}="${next}"`;
    }
    if (missing !== undefined) {
      const insertion = item.entry.start + 1 + item.entry.name.length;
      patches.push({
        end: insertion,
        replacement: ` ${missing}`,
        start: insertion,
      });
    }
  }
  for (const [attributeName, value] of [
    ['count', remaining.length],
    [
      'manualBreakCount',
      remaining.filter((entry) => entry.transformed!.manual).length,
    ],
  ] as const) {
    const source = attribute(container, attributeName);
    if (source && source.value !== String(value)) {
      patches.push(attributePatch(source, String(value)));
    }
  }
  return patches;
}

function worksheetViewPatches(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  prefix: string,
  request: XlsxWorksheetStructurePatch,
  part: string,
): TextPatch[] {
  const container = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}sheetViews`,
  );
  if (!container) return [];
  const containerIndex = tokens.indexOf(container);
  const close = xlsxMatchingCloseToken(tokens, containerIndex);
  const views = tokens
    .slice(containerIndex, tokens.indexOf(close))
    .filter(
      (token) =>
        !token.closing &&
        token.depth === container.depth + 1 &&
        token.name === `${prefix}sheetView`,
    );
  const patches: TextPatch[] = [];
  for (const view of views) {
    const viewIndex = tokens.indexOf(view);
    const viewClose = xlsxMatchingCloseToken(tokens, viewIndex);
    const children = tokens.slice(viewIndex, tokens.indexOf(viewClose));
    if (
      children.some(
        (token) =>
          !token.closing &&
          token.depth === view.depth + 1 &&
          token.name === `${prefix}pane`,
      )
    ) {
      failure(
        'XLSX structural worksheet pane cannot be preserved',
        part,
        request,
        'view-pane-reference',
      );
    }
    const topLeft = attribute(view, 'topLeftCell');
    if (topLeft) {
      if (!parseXlsxCellReference(topLeft.value)) {
        failure('XLSX structural view cell is invalid', part, request);
      }
      const transformed = transformXlsxStructuralVisualCell(
        topLeft.value,
        request,
      );
      if (transformed !== topLeft.value) {
        patches.push(attributePatch(topLeft, transformed));
      }
    }
    const selections = children.filter(
      (token) =>
        !token.closing &&
        token.depth === view.depth + 1 &&
        token.name === `${prefix}selection`,
    );
    for (const selection of selections) {
      const reference = attribute(selection, 'sqref');
      if (!reference) {
        failure('XLSX structural view selection is invalid', part, request);
      }
      const ranges = reference.value
        .trim()
        .split(/\s+/u)
        .map((value) => {
          const range = parseXlsxRangeReference(value);
          if (!range) {
            failure('XLSX structural view selection is invalid', part, request);
          }
          return range;
        });
      const activeCellSource = attribute(selection, 'activeCell');
      if (
        activeCellSource !== undefined &&
        !parseXlsxCellReference(activeCellSource.value)
      ) {
        failure('XLSX structural view active cell is invalid', part, request);
      }
      const activeCellIdSource = attribute(selection, 'activeCellId');
      const activeCellId =
        activeCellIdSource === undefined
          ? undefined
          : /^(?:0|[1-9]\d*)$/u.test(activeCellIdSource.value)
            ? Number(activeCellIdSource.value)
            : undefined;
      if (
        activeCellIdSource !== undefined &&
        (activeCellId === undefined || activeCellId >= ranges.length)
      ) {
        failure(
          'XLSX structural view active cell ID is invalid',
          part,
          request,
        );
      }
      const paneSource = attribute(selection, 'pane')?.value ?? 'topLeft';
      const pane = (
        {
          bottomLeft: 'bottom-left',
          bottomRight: 'bottom-right',
          topLeft: 'top-left',
          topRight: 'top-right',
        } satisfies Record<string, XlsxPanePosition>
      )[paneSource];
      if (!pane) {
        failure('XLSX structural view pane is invalid', part, request);
      }
      const transformed = transformXlsxStructuralViewSelection(
        {
          ...(activeCellSource === undefined
            ? {}
            : { activeCell: activeCellSource.value }),
          ...(activeCellId === undefined ? {} : { activeCellId }),
          pane,
          ranges,
        },
        request,
      );
      if (transformed === null) {
        const selectionClose = xlsxMatchingCloseToken(
          tokens,
          tokens.indexOf(selection),
        );
        patches.push({
          end: selectionClose.end,
          replacement: '',
          start: selection.start,
        });
        continue;
      }
      const transformedReference = transformed.ranges
        .map((range) => range.reference)
        .join(' ');
      if (
        transformed.ranges.length !== ranges.length ||
        transformed.ranges.some(
          (range, index) => range.reference !== ranges[index]!.reference,
        )
      ) {
        patches.push(attributePatch(reference, transformedReference));
      }
      if (
        activeCellSource &&
        transformed.activeCell !== activeCellSource.value
      ) {
        patches.push(attributePatch(activeCellSource, transformed.activeCell!));
      }
      if (activeCellIdSource && transformed.activeCellId !== activeCellId) {
        patches.push(
          attributePatch(activeCellIdSource, String(transformed.activeCellId)),
        );
      }
    }
  }
  return patches;
}

function layoutPatches(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  prefix: string,
  request: XlsxWorksheetStructurePatch,
  part: string,
): TextPatch[] {
  const patches: TextPatch[] = [
    ...worksheetViewPatches(tokens, root, prefix, request, part),
    ...pageBreakPatches(tokens, root, prefix, request, part, 'row'),
    ...pageBreakPatches(tokens, root, prefix, request, part, 'column'),
  ];
  const dimension = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}dimension`,
  );
  if (dimension) {
    const reference = attribute(dimension, 'ref');
    const range = parseXlsxRangeReference(reference?.value);
    if (!reference || !range) {
      failure('XLSX structural dimension reference is invalid', part, request);
    }
    const transformed = transformXlsxStructuralRange(range, request);
    if (transformed === null) {
      const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(dimension));
      patches.push({ end: close.end, replacement: '', start: dimension.start });
    } else if (transformed.reference !== range.reference) {
      patches.push(attributePatch(reference, transformed.reference));
    }
  }
  const autoFilter = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}autoFilter`,
  );
  if (autoFilter) {
    const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(autoFilter));
    const reference = attribute(autoFilter, 'ref');
    const range = parseXlsxRangeReference(reference?.value);
    if (!reference || !range) {
      failure('XLSX structural auto-filter range is invalid', part, request);
    }
    const transformed = transformXlsxStructuralRange(range, request);
    if (transformed === null) {
      patches.push({
        end: close.end,
        replacement: '',
        start: autoFilter.start,
      });
    } else {
      if (transformed.reference !== range.reference) {
        patches.push(attributePatch(reference, transformed.reference));
      }
      const sortState = tokens.find(
        (token) =>
          !token.closing &&
          token.depth === autoFilter.depth + 1 &&
          token.name === `${prefix}sortState` &&
          token.start >= autoFilter.end &&
          token.end <= close.start,
      );
      if (sortState) {
        const sortClose = xlsxMatchingCloseToken(
          tokens,
          tokens.indexOf(sortState),
        );
        const sortReference = attribute(sortState, 'ref');
        const sortRange = parseXlsxRangeReference(sortReference?.value);
        if (!sortReference || !sortRange) {
          failure('XLSX structural sort range is invalid', part, request);
        }
        const transformedSort = transformXlsxStructuralRange(
          sortRange,
          request,
        );
        if (transformedSort === null) {
          patches.push({
            end: sortClose.end,
            replacement: '',
            start: sortState.start,
          });
        } else {
          if (transformedSort.reference !== sortRange.reference) {
            patches.push(
              attributePatch(sortReference, transformedSort.reference),
            );
          }
          const conditions = tokens.filter(
            (token) =>
              !token.closing &&
              token.depth === sortState.depth + 1 &&
              token.name === `${prefix}sortCondition` &&
              token.start >= sortState.end &&
              token.end <= sortClose.start,
          );
          for (const condition of conditions) {
            const conditionReference = attribute(condition, 'ref');
            const conditionRange = parseXlsxRangeReference(
              conditionReference?.value,
            );
            if (!conditionReference || !conditionRange) {
              failure(
                'XLSX structural sort-condition range is invalid',
                part,
                request,
              );
            }
            const transformedCondition = transformXlsxStructuralRange(
              conditionRange,
              request,
            );
            if (transformedCondition === null) {
              const conditionClose = xlsxMatchingCloseToken(
                tokens,
                tokens.indexOf(condition),
              );
              patches.push({
                end: conditionClose.end,
                replacement: '',
                start: condition.start,
              });
            } else if (
              transformedCondition.reference !== conditionRange.reference
            ) {
              patches.push(
                attributePatch(
                  conditionReference,
                  transformedCondition.reference,
                ),
              );
            }
          }
        }
      }
    }
  }
  const protectedRanges = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}protectedRanges`,
  );
  if (protectedRanges) {
    const containerIndex = tokens.indexOf(protectedRanges);
    const close = xlsxMatchingCloseToken(tokens, containerIndex);
    const entries = tokens
      .slice(containerIndex, tokens.indexOf(close))
      .filter(
        (token) =>
          !token.closing &&
          token.depth === protectedRanges.depth + 1 &&
          token.name === `${prefix}protectedRange`,
      );
    const transformedEntries = entries.map((entry) => {
      const reference = attribute(entry, 'sqref');
      if (!reference) {
        failure('XLSX structural protected range is invalid', part, request);
      }
      const ranges = reference.value
        .trim()
        .split(/\s+/u)
        .map((value) => {
          const range = parseXlsxRangeReference(value);
          if (!range) {
            failure(
              'XLSX structural protected range is invalid',
              part,
              request,
            );
          }
          return range;
        });
      const transformed = ranges.flatMap((range) => {
        const result = transformXlsxStructuralRange(range, request);
        return result === null ? [] : [result];
      });
      return { entry, ranges, reference, transformed };
    });
    const remaining = transformedEntries.filter(
      (entry) => entry.transformed.length !== 0,
    );
    if (entries.length !== 0 && remaining.length === 0) {
      patches.push({
        end: close.end,
        replacement: '',
        start: protectedRanges.start,
      });
    } else {
      for (const item of transformedEntries) {
        if (item.transformed.length === 0) {
          const itemClose = xlsxMatchingCloseToken(
            tokens,
            tokens.indexOf(item.entry),
          );
          patches.push({
            end: itemClose.end,
            replacement: '',
            start: item.entry.start,
          });
        } else if (
          item.transformed.length !== item.ranges.length ||
          item.transformed.some(
            (range, index) => range.reference !== item.ranges[index]!.reference,
          )
        ) {
          patches.push(
            attributePatch(
              item.reference,
              item.transformed.map((range) => range.reference).join(' '),
            ),
          );
        }
      }
    }
  }
  const conditionalFormats = tokens.filter(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}conditionalFormatting`,
  );
  for (const format of conditionalFormats) {
    const formatIndex = tokens.indexOf(format);
    const close = xlsxMatchingCloseToken(tokens, formatIndex);
    const closeIndex = tokens.indexOf(close);
    const rules = tokens
      .slice(formatIndex, closeIndex)
      .filter(
        (token) =>
          !token.closing &&
          token.depth === format.depth + 1 &&
          token.name === `${prefix}cfRule`,
      );
    const hasFormula = rules.some((rule) => {
      const ruleIndex = tokens.indexOf(rule);
      const ruleClose = xlsxMatchingCloseToken(tokens, ruleIndex);
      return tokens
        .slice(ruleIndex, tokens.indexOf(ruleClose))
        .some(
          (token) =>
            !token.closing &&
            ((token.depth === rule.depth + 1 &&
              token.name === `${prefix}formula`) ||
              (token.name === `${prefix}cfvo` &&
                attribute(token, 'type')?.value === 'formula')),
        );
    });
    if (hasFormula) {
      failure(
        'XLSX structural conditional-format formula cannot be preserved',
        part,
        request,
        'conditional-format-formula-reference',
      );
    }
    const reference = attribute(format, 'sqref');
    if (!reference) {
      failure(
        'XLSX structural conditional-format range is invalid',
        part,
        request,
      );
    }
    const ranges = reference.value
      .trim()
      .split(/\s+/u)
      .map((value) => {
        const range = parseXlsxRangeReference(value);
        if (!range) {
          failure(
            'XLSX structural conditional-format range is invalid',
            part,
            request,
          );
        }
        return range;
      });
    const transformed = ranges.flatMap((range) => {
      const result = transformXlsxStructuralRange(range, request);
      return result === null ? [] : [result];
    });
    if (transformed.length === 0) {
      patches.push({ end: close.end, replacement: '', start: format.start });
    } else if (
      transformed.length !== ranges.length ||
      transformed.some(
        (range, index) => range.reference !== ranges[index]!.reference,
      )
    ) {
      patches.push(
        attributePatch(
          reference,
          transformed.map((range) => range.reference).join(' '),
        ),
      );
    }
  }
  const dataValidations = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}dataValidations`,
  );
  if (dataValidations) {
    const close = xlsxMatchingCloseToken(
      tokens,
      tokens.indexOf(dataValidations),
    );
    const entries = tokens.filter(
      (token) =>
        !token.closing &&
        token.depth === dataValidations.depth + 1 &&
        token.name === `${prefix}dataValidation` &&
        token.start >= dataValidations.end &&
        token.end <= close.start,
    );
    const transformedEntries = entries.map((entry) => {
      const entryClose = xlsxMatchingCloseToken(tokens, tokens.indexOf(entry));
      const hasFormula = tokens.some(
        (token) =>
          !token.closing &&
          token.depth === entry.depth + 1 &&
          (token.name === `${prefix}formula1` ||
            token.name === `${prefix}formula2`) &&
          token.start >= entry.end &&
          token.end <= entryClose.start,
      );
      if (hasFormula) {
        failure(
          'XLSX structural data-validation formula cannot be preserved',
          part,
          request,
          'data-validation-formula-reference',
        );
      }
      const reference = attribute(entry, 'sqref');
      if (!reference) {
        failure(
          'XLSX structural data-validation range is invalid',
          part,
          request,
        );
      }
      const ranges = reference.value
        .trim()
        .split(/\s+/u)
        .map((value) => {
          const range = parseXlsxRangeReference(value);
          if (!range) {
            failure(
              'XLSX structural data-validation range is invalid',
              part,
              request,
            );
          }
          return range;
        });
      const transformed = ranges.flatMap((range) => {
        const result = transformXlsxStructuralRange(range, request);
        return result === null ? [] : [result];
      });
      return { entry, entryClose, ranges, reference, transformed };
    });
    const remaining = transformedEntries.filter(
      (entry) => entry.transformed.length !== 0,
    );
    if (entries.length !== 0 && remaining.length === 0) {
      patches.push({
        end: close.end,
        replacement: '',
        start: dataValidations.start,
      });
    } else {
      for (const item of transformedEntries) {
        if (item.transformed.length === 0) {
          patches.push({
            end: item.entryClose.end,
            replacement: '',
            start: item.entry.start,
          });
        } else if (
          item.transformed.length !== item.ranges.length ||
          item.transformed.some(
            (range, index) => range.reference !== item.ranges[index]!.reference,
          )
        ) {
          patches.push(
            attributePatch(
              item.reference,
              item.transformed.map((range) => range.reference).join(' '),
            ),
          );
        }
      }
      const count = attribute(dataValidations, 'count');
      if (count && count.value !== String(remaining.length)) {
        patches.push(attributePatch(count, String(remaining.length)));
      }
    }
  }
  const hyperlinks = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}hyperlinks`,
  );
  if (hyperlinks) {
    const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(hyperlinks));
    const entries = tokens.filter(
      (token) =>
        !token.closing &&
        token.depth === hyperlinks.depth + 1 &&
        token.name === `${prefix}hyperlink` &&
        token.start >= hyperlinks.end &&
        token.end <= close.start,
    );
    const transformedEntries = entries.map((entry) => {
      const reference = attribute(entry, 'ref');
      const range = parseXlsxRangeReference(reference?.value);
      if (!reference || !range) {
        failure('XLSX structural hyperlink range is invalid', part, request);
      }
      return {
        entry,
        range,
        reference,
        transformed: transformXlsxStructuralRange(range, request),
      };
    });
    const remaining = transformedEntries.filter(
      (entry) => entry.transformed !== null,
    );
    if (remaining.length === 0) {
      patches.push({
        end: close.end,
        replacement: '',
        start: hyperlinks.start,
      });
    } else {
      for (const item of transformedEntries) {
        if (item.transformed === null) {
          const itemClose = xlsxMatchingCloseToken(
            tokens,
            tokens.indexOf(item.entry),
          );
          patches.push({
            end: itemClose.end,
            replacement: '',
            start: item.entry.start,
          });
        } else if (item.transformed.reference !== item.range.reference) {
          patches.push(
            attributePatch(item.reference, item.transformed.reference),
          );
        }
      }
    }
  }
  const mergeCells = tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}mergeCells`,
  );
  if (!mergeCells) return patches;
  const mergeClose = xlsxMatchingCloseToken(tokens, tokens.indexOf(mergeCells));
  const entries = tokens.filter(
    (token) =>
      !token.closing &&
      token.depth === mergeCells.depth + 1 &&
      token.name === `${prefix}mergeCell` &&
      token.start >= mergeCells.end &&
      token.end <= mergeClose.start,
  );
  const transformedEntries = entries.map((entry) => {
    const reference = attribute(entry, 'ref');
    const range = parseXlsxRangeReference(reference?.value);
    if (!reference || !range) {
      failure('XLSX structural merged range is invalid', part, request);
    }
    return {
      entry,
      range,
      reference,
      transformed: transformXlsxStructuralRange(range, request),
    };
  });
  const remaining = transformedEntries.filter(
    (entry) => entry.transformed !== null,
  );
  if (remaining.length === 0) {
    patches.push({
      end: mergeClose.end,
      replacement: '',
      start: mergeCells.start,
    });
    return patches;
  }
  for (const item of transformedEntries) {
    if (item.transformed === null) {
      const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(item.entry));
      patches.push({
        end: close.end,
        replacement: '',
        start: item.entry.start,
      });
    } else if (item.transformed.reference !== item.range.reference) {
      patches.push(attributePatch(item.reference, item.transformed.reference));
    }
  }
  const count = attribute(mergeCells, 'count');
  if (count && count.value !== String(remaining.length)) {
    patches.push(attributePatch(count, String(remaining.length)));
  }
  return patches;
}

function patchOne(
  bytes: Uint8Array,
  request: XlsxWorksheetStructurePatch,
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxWorksheetStructurePatchResult {
  const decoded = decodeXlsxXml(bytes, part);
  const tokens = tokenizeXlsxXml(decoded.text, part);
  const root = tokens.find(
    (token) =>
      token.depth === 0 && xlsxXmlLocalName(token.name) === 'worksheet',
  );
  if (!root)
    failure('XLSX worksheet root cannot patch structure', part, request);
  const { rows } = directRows(tokens, root, part);
  const prefix = root.name.slice(0, -'worksheet'.length);
  const patches = layoutPatches(tokens, root, prefix, request, part);
  for (const row of rows) {
    const rowReference = attribute(row, 'r');
    const rowIndex = Number(rowReference?.value);
    if (!rowReference || !Number.isSafeInteger(rowIndex)) {
      failure('XLSX structural target row reference is invalid', part, request);
    }
    const rowClose = xlsxMatchingCloseToken(tokens, tokens.indexOf(row));
    if (request.kind.endsWith('-rows')) {
      const shifted = shiftedIndex(rowIndex, request);
      if (shifted === null) {
        patches.push({ end: rowClose.end, replacement: '', start: row.start });
        continue;
      }
      if (shifted !== rowIndex) {
        patches.push(attributePatch(rowReference, String(shifted)));
      }
    }
    if (request.kind.endsWith('-columns')) {
      const spans = attribute(row, 'spans');
      if (spans)
        patches.push({ end: spans.end, replacement: '', start: spans.start });
    }
    const cells = tokens.filter(
      (token) =>
        !token.closing &&
        token.depth === row.depth + 1 &&
        token.name === `${prefix}c` &&
        token.start >= row.end &&
        token.end <= rowClose.start,
    );
    for (const cell of cells) {
      const reference = attribute(cell, 'r');
      const parsed = parseXlsxCellReference(reference?.value);
      if (!reference || !parsed) {
        failure(
          'XLSX structural target cell reference is invalid',
          part,
          request,
        );
      }
      if (request.kind.endsWith('-rows')) {
        const shiftedRow = shiftedIndex(parsed.row, request)!;
        if (shiftedRow !== parsed.row) {
          patches.push(
            attributePatch(
              reference,
              `${xlsxColumnName(parsed.column)!}${shiftedRow}`,
            ),
          );
        }
        continue;
      }
      const shiftedColumn = shiftedIndex(parsed.column, request);
      if (shiftedColumn === null) {
        const close = xlsxMatchingCloseToken(tokens, tokens.indexOf(cell));
        patches.push({ end: close.end, replacement: '', start: cell.start });
      } else if (shiftedColumn !== parsed.column) {
        patches.push(
          attributePatch(
            reference,
            `${xlsxColumnName(shiftedColumn)!}${parsed.row}`,
          ),
        );
      }
    }
  }
  let patchBytes = 0;
  for (const patch of patches) {
    patchBytes += encodeXlsxXml({
      bom: false,
      encoding: decoded.encoding,
      text: patch.replacement,
    }).byteLength;
  }
  patches.sort((left, right) => right.start - left.start);
  let output = decoded.text;
  for (const patch of patches) {
    output = `${output.slice(0, patch.start)}${patch.replacement}${output.slice(patch.end)}`;
  }
  const data = encodeXlsxXml({ ...decoded, text: output });
  if (data.byteLength > limits.maxGeneratedXmlBytes) {
    writeLimitFailure(
      'maxGeneratedXmlBytes',
      data.byteLength,
      limits.maxGeneratedXmlBytes,
      part,
    );
  }
  return { data, patchBytes, patchCount: patches.length };
}

export function patchXlsxWorksheetStructure(
  bytes: Uint8Array,
  requested: readonly XlsxWorksheetStructurePatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxWorksheetStructurePatchResult {
  let data: Uint8Array = bytes.slice();
  let patchBytes = 0;
  let patchCount = 0;
  for (const request of requested) {
    const result = patchOne(data, request, limits, part);
    data = result.data;
    patchBytes += result.patchBytes;
    patchCount += result.patchCount;
    if (patchBytes > limits.maxPatchBytes) {
      writeLimitFailure(
        'maxPatchBytes',
        patchBytes,
        limits.maxPatchBytes,
        part,
      );
    }
    if (patchCount > limits.maxPatchCount) {
      writeLimitFailure(
        'maxPatchCount',
        patchCount,
        limits.maxPatchCount,
        part,
      );
    }
  }
  return { data, patchBytes, patchCount };
}
