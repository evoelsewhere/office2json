import { describe, expect, it } from 'vitest';

import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import { patchXlsxWorksheetStructure } from '../../src/formats/xlsx/roundtrip/worksheet-structure-patch';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import { XLSX_SPREADSHEET_NS } from '../black-box/xlsx-package';

const PART = 'xl/worksheets/sheet1.xml';

function bytes(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function capture(action: () => unknown): XlsxWriteError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected structural patch to fail');
}

function source(): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1" spans="1:2"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="3" hidden="1" spans="3:3"><c r="C3" s="2"><v>3</v></c></row></sheetData></worksheet>`;
}

describe('XLSX worksheet structural patching', () => {
  it('applies ordered row and column insertions and deletions', () => {
    const result = patchXlsxWorksheetStructure(
      bytes(source()),
      [
        { count: 2, index: 2, kind: 'insert-rows', operationId: 'insert-rows' },
        { count: 1, index: 1, kind: 'delete-rows', operationId: 'delete-rows' },
        {
          count: 1,
          index: 2,
          kind: 'insert-columns',
          operationId: 'insert-columns',
        },
        {
          count: 1,
          index: 1,
          kind: 'delete-columns',
          operationId: 'delete-columns',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="4" hidden="1"><c r="C4" s="2"><v>3</v></c></row></sheetData></worksheet>`,
    );
    expect(result.patchCount).toBe(8);
  });

  it('preserves an owned byte copy when no operation is requested', () => {
    const input = bytes(source());
    const result = patchXlsxWorksheetStructure(
      input,
      [],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(result).toEqual({ data: input, patchBytes: 0, patchCount: 0 });
    expect(result.data).not.toBe(input);
  });

  it('distinguishes every insertion and deletion boundary', () => {
    const rows = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData>${[
      1, 2, 3, 4, 5,
    ]
      .map((row) => `<row r="${row}"><c r="A${row}"><v>${row}</v></c></row>`)
      .join('')}</sheetData></worksheet>`;
    const deletedRows = patchXlsxWorksheetStructure(
      bytes(rows),
      [{ count: 2, index: 2, kind: 'delete-rows', operationId: 'rows' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(deletedRows.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>4</v></c></row><row r="3"><c r="A3"><v>5</v></c></row></sheetData></worksheet>`,
    );
    const insertedRows = patchXlsxWorksheetStructure(
      bytes(rows),
      [{ count: 1, index: 2, kind: 'insert-rows', operationId: 'rows' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(insertedRows.data)).toContain(
      '<row r="1"><c r="A1"><v>1</v></c></row><row r="3"><c r="A3"><v>2</v></c></row>',
    );

    const columns = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1" spans="1:5"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1"><v>4</v></c><c r="E1"><v>5</v></c></row></sheetData></worksheet>`;
    const deletedColumns = patchXlsxWorksheetStructure(
      bytes(columns),
      [{ count: 2, index: 2, kind: 'delete-columns', operationId: 'columns' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(deletedColumns.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>4</v></c><c r="C1"><v>5</v></c></row></sheetData></worksheet>`,
    );
    expect(deletedColumns.patchCount).toBe(5);
    const insertedColumns = patchXlsxWorksheetStructure(
      bytes(columns),
      [{ count: 1, index: 2, kind: 'insert-columns', operationId: 'columns' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(insertedColumns.data)).toContain(
      '<c r="A1"><v>1</v></c><c r="C1"><v>2</v></c>',
    );
    expect(insertedColumns.patchCount).toBe(5);
  });

  it('patches only direct prefixed rows and cells', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}"><wrapper><s:row r="6"/><s:sheetData><s:row r="9"/></s:sheetData></wrapper><s:sheetData><other r="1"/><wrapper><s:row r="7"/></wrapper><s:row r="1"><wrapper><s:c r="Z9"/></wrapper><other r="A1"/><s:c r="A1"><s:v>1</s:v></s:c></s:row></s:sheetData><wrapper><s:row r="8"/></wrapper></s:worksheet>`;
    const result = patchXlsxWorksheetStructure(
      bytes(xml),
      [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'prefixed' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}"><wrapper><s:row r="6"/><s:sheetData><s:row r="9"/></s:sheetData></wrapper><s:sheetData><other r="1"/><wrapper><s:row r="7"/></wrapper><s:row r="2"><wrapper><s:c r="Z9"/></wrapper><other r="A1"/><s:c r="A2"><s:v>1</s:v></s:c></s:row></s:sheetData><wrapper><s:row r="8"/></wrapper></s:worksheet>`,
    );
  });

  it('transforms declared dimensions and merged ranges with exact counts', () => {
    const xml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><dimension ref="A1:C5"/><sheetData><row r="1"><c r="A1"/></row></sheetData><mergeCells count="2"><mergeCell ref="A2:B3"/><mergeCell ref="C1:C5"/></mergeCells><hyperlinks><hyperlink ref="A2:B3" location="Removed!A1"/><hyperlink ref="C1:C5" location="Kept!A1"/></hyperlinks></worksheet>`;
    const deleted = patchXlsxWorksheetStructure(
      bytes(xml),
      [{ count: 2, index: 2, kind: 'delete-rows', operationId: 'layout' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(deleted.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><dimension ref="A1:C3"/><sheetData><row r="1"><c r="A1"/></row></sheetData><mergeCells count="1"><mergeCell ref="C1:C3"/></mergeCells><hyperlinks><hyperlink ref="C1:C3" location="Kept!A1"/></hyperlinks></worksheet>`,
    );
    const removed = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><dimension ref="A2:B3"/><sheetData/><mergeCells count="1"><mergeCell ref="A2:B3"/></mergeCells><hyperlinks><hyperlink ref="A2:B3" location="Removed!A1"/></hyperlinks></worksheet>`,
      ),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'remove-layout',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removed.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/></worksheet>`,
    );
    const removedSort = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A1:C5"><sortState ref="A2:C3"><sortCondition ref="A2:A3"/></sortState></autoFilter></worksheet>`,
      ),
      [{ count: 2, index: 2, kind: 'delete-rows', operationId: 'sort' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removedSort.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A1:C3"></autoFilter></worksheet>`,
    );
  });

  it('selects only direct filter and sort nodes and avoids no-op patches', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:foreign"><s:sheetData/><wrapper><s:autoFilter ref="Z9"><s:sortState ref="Z9"><s:sortCondition ref="Z9"/></s:sortState></s:autoFilter><s:sortState ref="Z8"/></wrapper><x:autoFilter ref="Z9"/><s:autoFilter ref="A1:C3"><wrapper><s:sortState ref="Z9"/></wrapper><wrapper><s:sortCondition ref="Z8"/></wrapper><x:sortState ref="Z9"/><s:sortState ref="A1:C3"><other ref="Z9"/><wrapper><s:sortCondition ref="Z9"/></wrapper><x:sortCondition ref="Z9"/><s:sortCondition ref="A1:A3"/></s:sortState><wrapper><s:sortCondition ref="Z9"/></wrapper></s:autoFilter><wrapper><s:sortState ref="Z9"/></wrapper></s:worksheet>`;
    const unchanged = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 5,
          kind: 'insert-rows',
          operationId: 'unchanged-filter',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(unchanged.data)).toBe(xml);
    expect(unchanged.patchCount).toBe(0);
    const changed = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-rows',
          operationId: 'changed-filter',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const output = new TextDecoder().decode(changed.data);
    expect(output).toContain('<s:autoFilter ref="A2:C4">');
    expect(output).toContain('<s:sortState ref="A2:C4">');
    expect(output).toContain('<s:sortCondition ref="A2:A4"/>');
    expect(output).toContain(
      '<wrapper><s:autoFilter ref="Z9"><s:sortState ref="Z9"><s:sortCondition ref="Z9"/></s:sortState></s:autoFilter><s:sortState ref="Z8"/></wrapper>',
    );
    expect(output).toContain('<wrapper><s:sortState ref="Z9"/></wrapper>');
    expect(output).toContain('<wrapper><s:sortCondition ref="Z9"/></wrapper>');
    expect(output).toContain('<wrapper><s:sortCondition ref="Z8"/></wrapper>');
  });

  it('does not treat a following worksheet sort state as part of an auto-filter', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}"><s:sheetData/><s:autoFilter ref="A1:C3"/><wrapper><s:sortState ref="Z7"/></wrapper></s:worksheet>`;
    const result = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 0,
          kind: 'insert-rows',
          operationId: 'filter-ownership',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );

    expect(new TextDecoder().decode(result.data)).toBe(
      `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}"><s:sheetData/><s:autoFilter ref="A2:C4"/><wrapper><s:sortState ref="Z7"/></wrapper></s:worksheet>`,
    );
  });

  it('transforms worksheet filter and sort ranges', () => {
    const xml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A1:C5"><filterColumn colId="0"/><sortState ref="A1:C5"><sortCondition ref="A2:A3"/><sortCondition ref="B1:B5" descending="1"/></sortState></autoFilter></worksheet>`;
    const result = patchXlsxWorksheetStructure(
      bytes(xml),
      [{ count: 2, index: 2, kind: 'delete-rows', operationId: 'filter' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A1:C3"><filterColumn colId="0"/><sortState ref="A1:C3"><sortCondition ref="B1:B3" descending="1"/></sortState></autoFilter></worksheet>`,
    );
    const removed = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A2:C3"/></worksheet>`,
      ),
      [{ count: 2, index: 2, kind: 'delete-rows', operationId: 'filter' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removed.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/></worksheet>`,
    );

    const removedRange = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><dataValidations><dataValidation sqref="A1 C2:C3"/></dataValidations></worksheet>`,
      ),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'remove-one-validation-range',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removedRange.data)).toContain(
      '<dataValidation sqref="A1"/>',
    );
    const mixedShift = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><dataValidations><dataValidation sqref="A1 C2"/></dataValidations></worksheet>`,
      ),
      [
        {
          count: 1,
          index: 2,
          kind: 'insert-rows',
          operationId: 'shift-one-validation-range',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(mixedShift.data)).toContain(
      '<dataValidation sqref="A1 C3"/>',
    );
    const emptyXml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><dataValidations count="0" disablePrompts="1"/></worksheet>`;
    const empty = patchXlsxWorksheetStructure(
      bytes(emptyXml),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-rows',
          operationId: 'empty-validations',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(empty.data)).toBe(emptyXml);
    expect(empty.patchCount).toBe(0);
  });

  it('transforms named protected ranges and removes empty collections', () => {
    const xml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><protectedRanges><protectedRange name="One" sqref="A1:A5 C2:C3"/><protectedRange name="Two" sqref="B2:B3"/></protectedRanges></worksheet>`;
    const result = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'protected-ranges',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><protectedRanges><protectedRange name="One" sqref="A1:A3"/></protectedRanges></worksheet>`,
    );
    const removed = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><protectedRanges><protectedRange name="Input" sqref="A2:B3"/></protectedRanges></worksheet>`,
      ),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'remove-protected-ranges',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removed.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/></worksheet>`,
    );
    const mixed = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><protectedRanges><protectedRange name="Input" sqref="A1 C2"/></protectedRanges></worksheet>`,
      ),
      [
        {
          count: 1,
          index: 2,
          kind: 'insert-rows',
          operationId: 'mixed-protected-range',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(mixed.data)).toContain(
      '<protectedRange name="Input" sqref="A1 C3"/>',
    );
    const removedRange = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><protectedRanges><protectedRange name="Input" sqref="A1 C2:C3"/></protectedRanges></worksheet>`,
      ),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'remove-protected-range-segment',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removedRange.data)).toContain(
      '<protectedRange name="Input" sqref="A1"/>',
    );
    const emptyXml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><protectedRanges/></worksheet>`;
    const empty = patchXlsxWorksheetStructure(
      bytes(emptyXml),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-rows',
          operationId: 'empty-protected-ranges',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(empty.data)).toBe(emptyXml);
    expect(empty.patchCount).toBe(0);
  });

  it('transforms simple worksheet views and selection identities', () => {
    const xml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetViews><sheetView workbookViewId="0" topLeftCell="A2"><selection pane="topRight" activeCell="B2" activeCellId="1" sqref="A1:A2 B2"/></sheetView></sheetViews><sheetData/></worksheet>`;
    const inserted = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-rows',
          operationId: 'view-insert',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(inserted.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetViews><sheetView workbookViewId="0" topLeftCell="A3"><selection pane="topRight" activeCell="B3" activeCellId="1" sqref="A2:A3 B3"/></sheetView></sheetViews><sheetData/></worksheet>`,
    );
    const deleted = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetViews><sheetView workbookViewId="0" topLeftCell="A2"><selection activeCell="B3" activeCellId="1" sqref="A2 B3:B4"/></sheetView></sheetViews><sheetData/></worksheet>`,
      ),
      [
        {
          count: 1,
          index: 2,
          kind: 'delete-rows',
          operationId: 'view-delete',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(deleted.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetViews><sheetView workbookViewId="0" topLeftCell="A2"><selection activeCell="B2" activeCellId="0" sqref="B2:B3"/></sheetView></sheetViews><sheetData/></worksheet>`,
    );
    const removed = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetViews><sheetView workbookViewId="0"><selection sqref="A2"/></sheetView></sheetViews><sheetData/></worksheet>`,
      ),
      [
        {
          count: 1,
          index: 2,
          kind: 'delete-rows',
          operationId: 'view-remove-selection',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removed.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetViews><sheetView workbookViewId="0"></sheetView></sheetViews><sheetData/></worksheet>`,
    );
    const mixed = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetViews><sheetView><selection sqref="A1 C2"/></sheetView></sheetViews><sheetData/></worksheet>`,
      ),
      [
        {
          count: 1,
          index: 2,
          kind: 'insert-rows',
          operationId: 'view-mixed-ranges',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(mixed.data)).toContain(
      '<selection sqref="A1 C3"/>',
    );
    const removedRange = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetViews><sheetView><selection sqref="A1 C2:C3"/></sheetView></sheetViews><sheetData/></worksheet>`,
      ),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'view-removed-range',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removedRange.data)).toContain(
      '<selection sqref="A1"/>',
    );
  });

  it('selects only owned view nodes and rejects panes and malformed references', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:foreign"><wrapper><s:sheetViews><s:sheetView topLeftCell="Z9"><s:selection sqref="Z9"/></s:sheetView></s:sheetViews><s:sheetView topLeftCell="Z6"><s:pane/><s:selection sqref="Z6"/></s:sheetView></wrapper><x:sheetViews><x:sheetView topLeftCell="Z9"/></x:sheetViews><s:sheetViews><wrapper><s:sheetView topLeftCell="Z8"/></wrapper><x:sheetView topLeftCell="Z9"/><s:sheetView workbookViewId="0" topLeftCell="A1"><wrapper><s:pane/></wrapper><x:pane/><other/><wrapper><s:selection sqref="Z7"/></wrapper><x:selection sqref="Z9"/><s:selection pane="bottomLeft" activeCell="A1" activeCellId="0" sqref=" A1  B2 "/></s:sheetView></s:sheetViews><wrapper><s:sheetView topLeftCell="Z5"><s:pane/><s:selection sqref="Z5"/></s:sheetView></wrapper><s:sheetData/></s:worksheet>`;
    const unchanged = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 5,
          kind: 'insert-rows',
          operationId: 'view-no-op',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(unchanged.data)).toBe(xml);
    expect(unchanged.patchCount).toBe(0);
    const pane = capture(() =>
      patchXlsxWorksheetStructure(
        bytes(
          `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetViews><sheetView workbookViewId="0"><pane xSplit="1"/></sheetView></sheetViews><sheetData/></worksheet>`,
        ),
        [
          {
            count: 1,
            index: 1,
            kind: 'insert-rows',
            operationId: 'view-pane',
          },
        ],
        defaultXlsxWriteLimits(),
        PART,
      ),
    );
    expect(pane.diagnostic).toMatchObject({
      featureClass: 'view-pane-reference',
      message: 'XLSX structural worksheet pane cannot be preserved',
    });
    for (const [view, message] of [
      [
        '<sheetView topLeftCell="bad"/>',
        'XLSX structural view cell is invalid',
      ],
      [
        '<sheetView><selection/></sheetView>',
        'XLSX structural view selection is invalid',
      ],
      [
        '<sheetView><selection sqref="bad"/></sheetView>',
        'XLSX structural view selection is invalid',
      ],
      [
        '<sheetView><selection activeCell="bad" sqref="A1"/></sheetView>',
        'XLSX structural view active cell is invalid',
      ],
      [
        '<sheetView><selection activeCellId="1" sqref="A1"/></sheetView>',
        'XLSX structural view active cell ID is invalid',
      ],
      [
        '<sheetView><selection activeCellId="x1" sqref="A1"/></sheetView>',
        'XLSX structural view active cell ID is invalid',
      ],
      [
        '<sheetView><selection activeCellId="1x" sqref="A1 A2"/></sheetView>',
        'XLSX structural view active cell ID is invalid',
      ],
      [
        '<sheetView><selection activeCellId="1.0" sqref="A1 A2"/></sheetView>',
        'XLSX structural view active cell ID is invalid',
      ],
      [
        '<sheetView><selection pane="bad" sqref="A1"/></sheetView>',
        'XLSX structural view pane is invalid',
      ],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(
              `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetViews>${view}</sheetViews><sheetData/></worksheet>`,
            ),
            [
              {
                count: 1,
                index: 1,
                kind: 'insert-rows',
                operationId: 'bad-view',
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic.message,
      ).toBe(message);
    }
  });

  it('transforms row and column page breaks with exact counts', () => {
    const xml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><rowBreaks count="2" manualBreakCount="1"><brk id="2" min="0" max="2" man="1"/><brk id="3" min="0" max="2"/></rowBreaks><colBreaks count="2" manualBreakCount="1"><brk id="2" min="0" max="2" man="true"/><brk id="3" min="1" max="1"/></colBreaks></worksheet>`;
    const result = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 2,
          kind: 'delete-rows',
          operationId: 'page-breaks',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><rowBreaks count="1" manualBreakCount="0"><brk id="2" min="0" max="2"/></rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="2" min="0" max="1" man="true"/></colBreaks></worksheet>`,
    );
    const removed = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><rowBreaks><brk id="2"/></rowBreaks><colBreaks><brk id="2" min="1" max="1"/></colBreaks></worksheet>`,
      ),
      [
        {
          count: 1,
          index: 2,
          kind: 'delete-rows',
          operationId: 'remove-page-breaks',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removed.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/></worksheet>`,
    );
  });

  it('patches authored page-break bounds and preserves full-grid defaults', () => {
    const xml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><rowBreaks><brk id="1" min="1"/></rowBreaks><colBreaks><brk id="1" max="2"/></colBreaks></worksheet>`;
    const result = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-columns',
          operationId: 'page-break-columns',
        },
        {
          count: 1,
          index: 2,
          kind: 'insert-rows',
          operationId: 'page-break-rows',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><rowBreaks><brk id="1" min="2"/></rowBreaks><colBreaks><brk id="2" max="3"/></colBreaks></worksheet>`,
    );
    const missingBound = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><rowBreaks><brk id="1" max="2"/></rowBreaks></worksheet>`,
      ),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-columns',
          operationId: 'missing-break-bound',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(missingBound.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><rowBreaks><brk min="1" id="1" max="3"/></rowBreaks></worksheet>`,
    );
    const fullXml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><rowBreaks count="1" manualBreakCount="0"><brk id="1048576"/></rowBreaks><colBreaks count="1" manualBreakCount="0"><brk id="16384"/></colBreaks></worksheet>`;
    const full = patchXlsxWorksheetStructure(
      bytes(fullXml),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-rows',
          operationId: 'full-row-break',
        },
        {
          count: 1,
          index: 1,
          kind: 'insert-columns',
          operationId: 'full-column-break',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(full.data)).toBe(fullXml);
    expect(full.patchCount).toBe(0);
    const emptyXml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><rowBreaks count="0" manualBreakCount="0"/><colBreaks/></worksheet>`;
    const empty = patchXlsxWorksheetStructure(
      bytes(emptyXml),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-rows',
          operationId: 'empty-breaks',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(empty.data)).toBe(emptyXml);
    expect(empty.patchCount).toBe(0);
  });

  it('selects only owned page breaks and rejects malformed values', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:foreign"><s:sheetData/><wrapper><s:rowBreaks><s:brk id="9"/></s:rowBreaks></wrapper><x:rowBreaks><x:brk id="9"/></x:rowBreaks><s:rowBreaks count="2" manualBreakCount="0"><wrapper><s:brk id="8"/></wrapper><x:brk id="9"/><s:brk id="1" min="0" max="1" man="0"/><s:brk id="4" min="0" max="1" pt="false"/></s:rowBreaks><wrapper><s:brk id="7"/></wrapper></s:worksheet>`;
    const unchanged = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 5,
          kind: 'insert-rows',
          operationId: 'break-no-op',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(unchanged.data)).toBe(xml);
    expect(unchanged.patchCount).toBe(0);
    for (const source of [
      '<brk/>',
      '<brk id="01"/>',
      '<brk id="1" min="bad"/>',
      '<brk id="1" man="bad"/>',
      '<brk id="1" pt="bad"/>',
    ]) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(
              `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><rowBreaks>${source}</rowBreaks></worksheet>`,
            ),
            [
              {
                count: 1,
                index: 1,
                kind: 'insert-rows',
                operationId: 'bad-break',
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic.message,
      ).toMatch(/^XLSX structural page-break (?:flag|value) is invalid$/u);
    }
  });

  it('selects only owned protected ranges and rejects malformed ranges', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:foreign"><s:sheetData/><wrapper><s:protectedRanges><s:protectedRange name="Nested" sqref="Z9"/></s:protectedRanges><s:protectedRange name="Before" sqref="Z8"/></wrapper><x:protectedRanges><x:protectedRange sqref="Z9"/></x:protectedRanges><s:protectedRanges><wrapper><s:protectedRange name="NestedEntry" sqref="Z7"/></wrapper><x:protectedRange name="Foreign" sqref="Z9"/><s:protectedRange name="Input" sqref=" A1  B2 "/></s:protectedRanges><wrapper><s:protectedRange name="After" sqref="Z6"/></wrapper></s:worksheet>`;
    const unchanged = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 5,
          kind: 'insert-rows',
          operationId: 'protected-no-op',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(unchanged.data)).toBe(xml);
    expect(unchanged.patchCount).toBe(0);
    const changed = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-rows',
          operationId: 'protected-owned',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const output = new TextDecoder().decode(changed.data);
    expect(output).toContain('<s:protectedRange name="Input" sqref="A2 B3"/>');
    expect(output).toContain(
      '<wrapper><s:protectedRanges><s:protectedRange name="Nested" sqref="Z9"/></s:protectedRanges>',
    );
    expect(output).toContain('<x:protectedRange sqref="Z9"/>');
    expect(output).toContain(
      '<wrapper><s:protectedRange name="NestedEntry" sqref="Z7"/></wrapper>',
    );
    expect(output).toContain(
      '<wrapper><s:protectedRange name="After" sqref="Z6"/></wrapper>',
    );
    for (const source of [
      '<protectedRange name="Input"/>',
      '<protectedRange name="Input" sqref=""/>',
      '<protectedRange name="Input" sqref="bad"/>',
    ]) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(
              `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><protectedRanges>${source}</protectedRanges></worksheet>`,
            ),
            [
              {
                count: 1,
                index: 1,
                kind: 'insert-rows',
                operationId: 'bad-protected-range',
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic.message,
      ).toBe('XLSX structural protected range is invalid');
    }
  });

  it('transforms formula-free conditional-format ranges', () => {
    const xml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><conditionalFormatting sqref="A1:A5 C2:C3"><cfRule type="top10" priority="1" rank="1"/></conditionalFormatting><conditionalFormatting sqref="D2:D3"><cfRule type="uniqueValues" priority="2"/></conditionalFormatting></worksheet>`;
    const result = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'formats',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><conditionalFormatting sqref="A1:A3"><cfRule type="top10" priority="1" rank="1"/></conditionalFormatting></worksheet>`,
    );
    const mixed = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><conditionalFormatting sqref="A1 C2"><cfRule type="uniqueValues" priority="1"/></conditionalFormatting></worksheet>`,
      ),
      [
        {
          count: 1,
          index: 2,
          kind: 'insert-rows',
          operationId: 'mixed-format',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(mixed.data)).toContain(
      '<conditionalFormatting sqref="A1 C3">',
    );
    const removedRange = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><conditionalFormatting sqref="A1 C2:C3"><cfRule type="uniqueValues" priority="1"/></conditionalFormatting></worksheet>`,
      ),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'remove-format-range',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removedRange.data)).toContain(
      '<conditionalFormatting sqref="A1">',
    );
  });

  it('selects only owned conditional formats and preserves lexical no-ops', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:foreign"><s:sheetData/><wrapper><s:conditionalFormatting sqref="Z9"><s:cfRule type="uniqueValues" priority="1"/></s:conditionalFormatting><s:cfRule type="expression" priority="9"><s:formula/></s:cfRule></wrapper><x:conditionalFormatting sqref="Z9"/><s:conditionalFormatting sqref=" A1  B2 "><wrapper><s:formula/></wrapper><wrapper><s:cfRule type="expression" priority="8"><s:formula/></s:cfRule></wrapper><x:cfRule><s:formula/></x:cfRule><s:cfRule type="uniqueValues" priority="1"><other></other><wrapper><s:formula/></wrapper><x:formula/><s:cfvo type="num"/><x:cfvo type="formula"/></s:cfRule><wrapper><s:formula/></wrapper></s:conditionalFormatting><wrapper><s:conditionalFormatting sqref="Z8"/><s:cfRule type="expression" priority="7"><s:formula/></s:cfRule></wrapper></s:worksheet>`;
    const unchanged = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 5,
          kind: 'insert-rows',
          operationId: 'format-no-op',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(unchanged.data)).toBe(xml);
    expect(unchanged.patchCount).toBe(0);
    const changed = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-rows',
          operationId: 'format-owned',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const output = new TextDecoder().decode(changed.data);
    expect(output).toContain('<s:conditionalFormatting sqref="A2 B3">');
    expect(output).toContain(
      '<wrapper><s:conditionalFormatting sqref="Z9"><s:cfRule type="uniqueValues" priority="1"/></s:conditionalFormatting>',
    );
    expect(output).toContain('<x:conditionalFormatting sqref="Z9"/>');
    expect(output).toContain('<wrapper><s:conditionalFormatting sqref="Z8"/>');
    expect(output).toContain('<wrapper><s:formula/></wrapper>');
    expect(output).toContain('<x:formula/>');
    expect(output).toContain('<s:cfvo type="num"/>');
    expect(output).toContain('<x:cfvo type="formula"/>');
    expect(output).toContain(
      '<s:cfRule type="expression" priority="9"><s:formula/></s:cfRule>',
    );
    expect(output).toContain(
      '<s:cfRule type="expression" priority="7"><s:formula/></s:cfRule>',
    );
  });

  it('rejects formula-bearing and malformed structural conditional formats', () => {
    for (const rule of [
      '<cfRule type="expression" priority="1"><formula/></cfRule>',
      '<cfRule type="colorScale" priority="1"><colorScale><cfvo type="formula"/></colorScale></cfRule>',
    ]) {
      const error = capture(() =>
        patchXlsxWorksheetStructure(
          bytes(
            `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><conditionalFormatting sqref="A1">${rule}</conditionalFormatting></worksheet>`,
          ),
          [
            {
              count: 1,
              index: 1,
              kind: 'insert-rows',
              operationId: 'formula-format',
            },
          ],
          defaultXlsxWriteLimits(),
          PART,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        featureClass: 'conditional-format-formula-reference',
        message:
          'XLSX structural conditional-format formula cannot be preserved',
        operationId: 'formula-format',
      });
    }
    for (const source of [
      '<conditionalFormatting/>',
      '<conditionalFormatting sqref=""/>',
      '<conditionalFormatting sqref="bad"/>',
    ]) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(
              `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/>${source}</worksheet>`,
            ),
            [
              {
                count: 1,
                index: 1,
                kind: 'insert-rows',
                operationId: 'bad-format',
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic.message,
      ).toBe('XLSX structural conditional-format range is invalid');
    }
  });

  it('transforms formula-free data-validation ranges and exact counts', () => {
    const xml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><dataValidations count="2" disablePrompts="1"><dataValidation sqref="A1:A5 C2:C3"/><dataValidation sqref="B2:B3"/></dataValidations></worksheet>`;
    const result = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'validations',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><dataValidations count="1" disablePrompts="1"><dataValidation sqref="A1:A3"/></dataValidations></worksheet>`,
    );

    const removed = patchXlsxWorksheetStructure(
      bytes(
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><dataValidations count="1" disablePrompts="1"><dataValidation sqref="A2:B3"/></dataValidations></worksheet>`,
      ),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'remove-validations',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removed.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/></worksheet>`,
    );
  });

  it('selects only owned data-validation nodes and preserves lexical no-ops', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:foreign"><s:sheetData/><wrapper><s:dataValidations count="1"><s:dataValidation sqref="Z9"/></s:dataValidations><s:dataValidation sqref="Z6"/></wrapper><x:dataValidations><x:dataValidation sqref="Z9"/></x:dataValidations><s:dataValidations count="1"><wrapper><s:formula1/></wrapper><wrapper><s:dataValidation sqref="Z8"/></wrapper><x:dataValidation sqref="Z9"/><s:dataValidation sqref=" A1  B2 "><other></other><wrapper><s:formula1/></wrapper><x:formula1/></s:dataValidation><wrapper><s:formula2/></wrapper></s:dataValidations><wrapper><s:dataValidation sqref="Z7"/></wrapper></s:worksheet>`;
    const unchanged = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 5,
          kind: 'insert-rows',
          operationId: 'validation-no-op',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(unchanged.data)).toBe(xml);
    expect(unchanged.patchCount).toBe(0);
    const changed = patchXlsxWorksheetStructure(
      bytes(xml),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-rows',
          operationId: 'validation-owned',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const output = new TextDecoder().decode(changed.data);
    expect(output).toContain('<s:dataValidation sqref="A2 B3">');
    expect(output).toContain(
      '<wrapper><s:dataValidations count="1"><s:dataValidation sqref="Z9"/></s:dataValidations><s:dataValidation sqref="Z6"/></wrapper>',
    );
    expect(output).toContain('<x:dataValidation sqref="Z9"/>');
    expect(output).toContain(
      '<wrapper><s:dataValidation sqref="Z8"/></wrapper>',
    );
    expect(output).toContain(
      '<wrapper><s:dataValidation sqref="Z7"/></wrapper>',
    );
    expect(output).toContain('<wrapper><s:formula1/></wrapper>');
    expect(output).toContain('<x:formula1/>');
    expect(output).toContain('<wrapper><s:formula2/></wrapper>');
  });

  it('rejects formula-bearing and malformed structural data validations', () => {
    for (const formulaName of ['formula1', 'formula2']) {
      const formula = capture(() =>
        patchXlsxWorksheetStructure(
          bytes(
            `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><dataValidations><dataValidation sqref="A1"><${formulaName}/></dataValidation></dataValidations></worksheet>`,
          ),
          [
            {
              count: 1,
              index: 1,
              kind: 'insert-rows',
              operationId: 'formula-validation',
            },
          ],
          defaultXlsxWriteLimits(),
          PART,
        ),
      );
      expect(formula.diagnostic).toMatchObject({
        featureClass: 'data-validation-formula-reference',
        message: 'XLSX structural data-validation formula cannot be preserved',
        operationId: 'formula-validation',
      });
    }
    for (const source of [
      '<dataValidation/>',
      '<dataValidation sqref=""/>',
      '<dataValidation sqref="bad"/>',
    ]) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(
              `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><dataValidations>${source}</dataValidations></worksheet>`,
            ),
            [
              {
                count: 1,
                index: 1,
                kind: 'insert-rows',
                operationId: 'bad-validation',
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic.message,
      ).toBe('XLSX structural data-validation range is invalid');
    }
  });

  it('selects only direct layout nodes and avoids no-op patches', () => {
    const xml = `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:foreign"><wrapper><s:dimension ref="Z9"/></wrapper><x:dimension ref="Z9"/><s:dimension ref="A1:B2"/><s:sheetData/><wrapper><s:hyperlinks><s:hyperlink ref="Z9" location="Nested!A1"/></s:hyperlinks><s:hyperlink ref="Z8"/></wrapper><x:hyperlinks><x:hyperlink ref="Z9"/></x:hyperlinks><s:hyperlinks><other ref="Z9"/><wrapper><s:hyperlink ref="Z9"/></wrapper><x:hyperlink ref="Z9"/><s:hyperlink ref="A1:B2" location="Real!A1"/></s:hyperlinks><wrapper><s:hyperlink ref="Z9"/></wrapper><wrapper><s:mergeCells count="1"><s:mergeCell ref="Z9"/></s:mergeCells><s:mergeCell ref="Z9"/></wrapper><s:mergeCells count="1"><other ref="Z9"/><wrapper><s:mergeCell ref="Z9"/></wrapper><x:mergeCell ref="Z9"/><s:mergeCell ref="A1:B2"/></s:mergeCells><wrapper><s:mergeCell ref="Z9"/></wrapper></s:worksheet>`;
    const unchanged = patchXlsxWorksheetStructure(
      bytes(xml),
      [{ count: 1, index: 5, kind: 'insert-rows', operationId: 'unchanged' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(unchanged.data)).toBe(xml);
    expect(unchanged.patchCount).toBe(0);
    const changed = patchXlsxWorksheetStructure(
      bytes(xml),
      [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'changed' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    const output = new TextDecoder().decode(changed.data);
    expect(output).toContain('<s:dimension ref="A2:B3"/>');
    expect(output).toContain('<s:mergeCell ref="A2:B3"/>');
    expect(output).toContain('<s:hyperlink ref="A2:B3" location="Real!A1"/>');
    expect(output).toContain(
      '<wrapper><s:hyperlinks><s:hyperlink ref="Z9" location="Nested!A1"/></s:hyperlinks><s:hyperlink ref="Z8"/></wrapper>',
    );
    expect(output).toContain('<x:hyperlink ref="Z9"/>');
    expect(output).toContain('<wrapper><s:dimension ref="Z9"/></wrapper>');
    expect(output).toContain('<wrapper><s:mergeCell ref="Z9"/></wrapper>');
    expect(output).toContain(
      '<wrapper><s:mergeCells count="1"><s:mergeCell ref="Z9"/></s:mergeCells><s:mergeCell ref="Z9"/></wrapper>',
    );
    expect(output).toContain('<x:mergeCell ref="Z9"/>');
  });

  it('rejects malformed dimension and merge references exactly', () => {
    for (const [xml, message] of [
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><dimension ref="bad"/><sheetData/></worksheet>`,
        'XLSX structural dimension reference is invalid',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><mergeCells><mergeCell ref="bad"/></mergeCells></worksheet>`,
        'XLSX structural merged range is invalid',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><hyperlinks><hyperlink ref="bad"/></hyperlinks></worksheet>`,
        'XLSX structural hyperlink range is invalid',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="bad"/></worksheet>`,
        'XLSX structural auto-filter range is invalid',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A1"><sortState ref="bad"/></autoFilter></worksheet>`,
        'XLSX structural sort range is invalid',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/><autoFilter ref="A1"><sortState ref="A1"><sortCondition ref="bad"/></sortState></autoFilter></worksheet>`,
        'XLSX structural sort-condition range is invalid',
      ],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(xml),
            [
              {
                count: 1,
                index: 1,
                kind: 'insert-rows',
                operationId: 'bad-layout',
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic.message,
      ).toBe(message);
    }
  });

  it('rejects unsafe roots, missing sheetData, and malformed references', () => {
    for (const [xml, message] of [
      [
        `<outer>${source()}</outer>`,
        'XLSX worksheet root cannot patch structure',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"/>`,
        'XLSX worksheet sheetData cannot patch structure',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row><c r="A1"/></row></sheetData></worksheet>`,
        'XLSX structural target row reference is invalid',
      ],
      [
        `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c/></row></sheetData></worksheet>`,
        'XLSX structural target cell reference is invalid',
      ],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(xml),
            [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'bad' }],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ message, part: PART });
    }
    expect(
      capture(() =>
        patchXlsxWorksheetStructure(
          bytes(
            `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row/></sheetData></worksheet>`,
          ),
          [
            {
              count: 2,
              index: 3,
              kind: 'insert-rows',
              operationId: 'bad-range',
            },
          ],
          defaultXlsxWriteLimits(),
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({
      featureClass: 'worksheet-structure-xml',
      operationId: 'bad-range',
      range: '3:4',
    });
  });

  it('enforces patch count, patch bytes, and generated bytes exactly', () => {
    const request = [
      {
        count: 1,
        index: 1,
        kind: 'insert-rows' as const,
        operationId: 'insert',
      },
    ];
    const successful = patchXlsxWorksheetStructure(
      bytes(source()),
      request,
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(successful.patchBytes).toBe(33);
    expect(() =>
      patchXlsxWorksheetStructure(
        bytes(source()),
        request,
        {
          ...defaultXlsxWriteLimits(),
          maxGeneratedXmlBytes: successful.data.byteLength,
          maxPatchBytes: successful.patchBytes,
          maxPatchCount: successful.patchCount,
        },
        PART,
      ),
    ).not.toThrow();
    for (const [limitName, limit] of [
      ['maxGeneratedXmlBytes', successful.data.byteLength - 1],
      ['maxPatchBytes', successful.patchBytes - 1],
      ['maxPatchCount', successful.patchCount - 1],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(source()),
            request,
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ limit, limitName, part: PART });
    }

    const first = patchXlsxWorksheetStructure(
      bytes(source()),
      request,
      defaultXlsxWriteLimits(),
      PART,
    );
    const secondRequest = [
      {
        count: 1,
        index: 2,
        kind: 'insert-rows' as const,
        operationId: 'second',
      },
    ];
    const second = patchXlsxWorksheetStructure(
      first.data,
      secondRequest,
      defaultXlsxWriteLimits(),
      PART,
    );
    const aggregateBytes = first.patchBytes + second.patchBytes;
    const aggregateCount = first.patchCount + second.patchCount;
    const individualBytes = Math.max(first.patchBytes, second.patchBytes);
    const individualCount = Math.max(first.patchCount, second.patchCount);
    for (const [limitName, limit, actual] of [
      ['maxPatchBytes', individualBytes, aggregateBytes],
      ['maxPatchCount', individualCount, aggregateCount],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxWorksheetStructure(
            bytes(source()),
            [...request, ...secondRequest],
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ actual, limit, limitName, part: PART });
    }
  });
});
