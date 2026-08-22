import { describe, expect, it } from 'vitest';

import { resolveXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import { canonicalXlsxJson } from '../../src/formats/xlsx/roundtrip/canonical-json';
import { canonicalXlsxSha256 } from '../../src/formats/xlsx/roundtrip/digest';
import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import {
  replayXlsxCellOperations,
  xlsxCellTargetState,
  xlsxColumnTargetState,
  xlsxRowTargetState,
  xlsxStructuralTargetState,
} from '../../src/formats/xlsx/roundtrip/operation-planner';
import { validateXlsxCellOperations } from '../../src/formats/xlsx/roundtrip/operation-validation';
import { readXlsxRoundTrip } from '../../src/formats/xlsx/roundtrip/read-snapshot';
import type {
  XlsxEditOperation,
  XlsxRoundTripDocument,
} from '../../src/formats/xlsx/roundtrip/types';
import {
  defaultXlsxWriteLimits,
  resolveXlsxWriteLimits,
} from '../../src/formats/xlsx/roundtrip/write-limits';
import type {
  XlsxDrawing,
  XlsxSparklineGroup,
  XlsxTable,
  XlsxWorksheet,
} from '../../src/formats/xlsx/types';
import {
  createIndependentXlsx,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const writeLimits = defaultXlsxWriteLimits();
const readerLimits = resolveXlsxResourceLimits();
const ERROR_CODES = [
  '#BLOCKED!',
  '#BUSY!',
  '#CALC!',
  '#CONNECT!',
  '#DIV/0!',
  '#FIELD!',
  '#GETTING_DATA',
  '#N/A',
  '#NAME?',
  '#NULL!',
  '#NUM!',
  '#REF!',
  '#SPILL!',
  '#UNKNOWN!',
  '#VALUE!',
] as const;
const UNSUPPORTED_OPERATION_KINDS = [
  'add-worksheet',
  'delete-worksheet',
  'rename-worksheet',
] as const;

function capture(action: () => unknown): XlsxWriteError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected XLSX operation validation to fail');
}

async function captureAsync(
  action: () => Promise<unknown>,
): Promise<XlsxWriteError> {
  try {
    await action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected XLSX operation planning to fail');
}

function worksheet(document: XlsxRoundTripDocument): XlsxWorksheet & {
  key: string;
} {
  const sheet = document.sheets[0]!;
  expect(sheet.kind).toBe('worksheet');
  return sheet as XlsxWorksheet & { key: string };
}

function structuralTable(overrides: Partial<XlsxTable> = {}): XlsxTable {
  return {
    autoFilter: {
      columns: [],
      range: {
        end: { column: 2, row: 3 },
        reference: 'A1:B3',
        start: { column: 1, row: 1 },
      },
      selectionRelation: 'full-sheet',
      sort: {
        caseSensitive: false,
        columnSort: false,
        conditions: [
          {
            descending: false,
            range: {
              end: { column: 1, row: 3 },
              reference: 'A2:A3',
              start: { column: 1, row: 2 },
            },
            sortBy: 'value',
          },
        ],
        range: {
          end: { column: 2, row: 3 },
          reference: 'A1:B3',
          start: { column: 1, row: 1 },
        },
        sortMethod: 'none',
      },
    },
    columns: [
      { id: 1, name: 'A', totalsFunction: 'none' },
      { id: 2, name: 'B', totalsFunction: 'none' },
    ],
    displayName: 'Table1',
    headerRow: true,
    id: 1,
    insertRow: false,
    insertRowShift: false,
    name: 'Table1',
    published: false,
    range: {
      end: { column: 2, row: 3 },
      reference: 'A1:B3',
      start: { column: 1, row: 1 },
    },
    selectionRelation: 'full-sheet',
    tableType: 'worksheet',
    totalsRow: false,
    totalsRowShown: true,
    ...overrides,
  };
}

function structuralDrawing(overrides: Partial<XlsxDrawing> = {}): XlsxDrawing {
  return {
    extent: { height: 10, width: 10 },
    from: { column: 12, columnOffset: 0, row: 2, rowOffset: 0 },
    kind: 'one-cell',
    object: {
      geometry: { kind: 'preset', preset: 'rect' },
      hidden: false,
      id: 1,
      kind: 'shape',
      name: 'Shape 1',
      transform: {
        flipHorizontal: false,
        flipVertical: false,
        height: 10,
        rotation: 0,
        width: 10,
        x: 0,
        y: 0,
      },
    },
    selectionRelation: 'full-sheet',
    ...overrides,
  };
}

function structuralSparklineGroup(
  overrides: Partial<XlsxSparklineGroup> = {},
): XlsxSparklineGroup {
  return {
    colors: {},
    dateAxis: false,
    displayEmptyCellsAs: 'zero',
    displayHidden: false,
    displayXAxis: false,
    first: false,
    high: false,
    last: false,
    low: false,
    markers: false,
    maximumAxisType: 'individual',
    minimumAxisType: 'individual',
    negative: false,
    rightToLeft: false,
    sparklines: [
      {
        dataFormula: 'A1:A3',
        location: 'L10',
        selectionRelation: 'full-sheet',
      },
    ],
    type: 'line',
    ...overrides,
  };
}

function cellOperation(
  document: XlsxRoundTripDocument,
  overrides: Partial<Extract<XlsxEditOperation, { kind: 'set-cell' }>> = {},
): Extract<XlsxEditOperation, { kind: 'set-cell' }> {
  return {
    cell: 'A1',
    content: { kind: 'value', value: { kind: 'text', text: 'updated' } },
    kind: 'set-cell',
    operationId: 'edit-1',
    sheetKey: worksheet(document).key,
    ...overrides,
  };
}

describe('XLSX cell operation validation', () => {
  it('normalizes every supported scalar payload and formula', () => {
    const common = {
      cell: 'A1',
      operationId: 'agent:edit_1',
      sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
    };
    const operations = validateXlsxCellOperations(
      [
        {
          ...common,
          content: { kind: 'value', value: { kind: 'number', value: -0 } },
          kind: 'set-cell',
        },
        {
          ...common,
          content: { kind: 'value', value: { kind: 'boolean', value: false } },
          kind: 'set-cell',
          operationId: 'boolean',
        },
        {
          ...common,
          content: { kind: 'value', value: { code: '#N/A', kind: 'error' } },
          kind: 'set-cell',
          operationId: 'error',
        },
        {
          ...common,
          content: { kind: 'formula', expression: 'SUM(A2:A3)' },
          ifMatch: 'b'.repeat(64),
          kind: 'set-cell',
          operationId: 'formula',
        },
        { ...common, kind: 'clear-cell', operationId: 'clear' },
        {
          ...common,
          ifMatch: 'c'.repeat(64),
          kind: 'set-hyperlink',
          operationId: 'internal-link',
          target: { kind: 'internal', location: "'Sheet 2'!A1" },
        },
        {
          ...common,
          kind: 'set-hyperlink',
          operationId: 'external-link',
          target: {
            kind: 'external',
            location: 'Section',
            url: 'https://example.invalid/path',
          },
        },
        {
          ...common,
          kind: 'set-hyperlink',
          operationId: 'remove-link',
          target: null,
        },
      ],
      writeLimits,
      readerLimits,
    );
    expect(operations).toHaveLength(8);
    expect(operations[0]).toMatchObject({
      content: { value: { kind: 'number', value: 0 } },
    });
    expect(
      Object.is(
        (operations[0] as { content: { value: { value: number } } }).content
          .value.value,
        -0,
      ),
    ).toBe(false);
    expect(operations[3]).toMatchObject({
      content: { expression: 'SUM(A2:A3)', kind: 'formula' },
      ifMatch: 'b'.repeat(64),
    });
    expect(operations.slice(5)).toEqual([
      expect.objectContaining({
        ifMatch: 'c'.repeat(64),
        kind: 'set-hyperlink',
        target: { kind: 'internal', location: "'Sheet 2'!A1" },
      }),
      expect.objectContaining({
        kind: 'set-hyperlink',
        target: {
          kind: 'external',
          location: 'Section',
          url: 'https://example.invalid/path',
        },
      }),
      expect.objectContaining({ kind: 'set-hyperlink', target: null }),
    ]);
  });

  it.each([
    [null, 'XLSX round-trip operations must be an array'],
    [[null], 'XLSX operation shape is invalid'],
    [[[]], 'XLSX operation shape is invalid'],
    [[Object.create(null)], 'XLSX operation shape is invalid'],
    [[{ kind: 'wat', operationId: 'one' }], 'XLSX operation kind is invalid'],
    [
      [
        {
          cell: 'a1',
          kind: 'clear-cell',
          operationId: 'one',
          sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
        },
      ],
      'XLSX operation cell reference is invalid',
    ],
    [
      [
        {
          cell: 'A1',
          extra: true,
          kind: 'clear-cell',
          operationId: 'one',
          sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
        },
      ],
      'XLSX clear-cell operation shape is invalid',
    ],
    [
      [
        {
          cell: 'A1',
          kind: 'clear-cell',
          operationId: '-bad',
          sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
        },
      ],
      'XLSX operation ID is invalid',
    ],
    [
      [
        {
          cell: 'A1',
          kind: 'clear-cell',
          operationId: 'one',
          sheetKey: 'Sheet1',
        },
      ],
      'XLSX operation sheet key is invalid',
    ],
    [
      [
        {
          cell: 'A1',
          ifMatch: 'A'.repeat(64),
          kind: 'clear-cell',
          operationId: 'one',
          sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
        },
      ],
      'XLSX operation precondition hash is invalid',
    ],
  ] as const)('rejects invalid operation contract %#', (value, message) => {
    expect(
      capture(() =>
        validateXlsxCellOperations(value, writeLimits, readerLimits),
      ).diagnostic.message,
    ).toBe(message);
  });

  it.each([
    ['one!', 'XLSX operation ID is invalid'],
    [`x${'a'.repeat(64)}`, 'XLSX operation precondition hash is invalid'],
    [`${'a'.repeat(64)}x`, 'XLSX operation precondition hash is invalid'],
    [`xxlsx:sheet:${'a'.repeat(32)}`, 'XLSX operation sheet key is invalid'],
    [`xlsx:sheet:${'a'.repeat(32)}x`, 'XLSX operation sheet key is invalid'],
  ] as const)(
    'requires whole-string operation identities %#',
    (value, message) => {
      const operation: Record<string, unknown> = {
        cell: 'A1',
        kind: 'clear-cell',
        operationId: 'one',
        sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
      };
      if (message.includes('ID')) operation.operationId = value;
      else if (message.includes('precondition')) operation.ifMatch = value;
      else operation.sheetKey = value;
      expect(
        capture(() =>
          validateXlsxCellOperations([operation], writeLimits, readerLimits),
        ).diagnostic.message,
      ).toBe(message);
    },
  );

  it('validates structural row and column operations at exact grid bounds', () => {
    const sheetKey = `xlsx:sheet:${'a'.repeat(32)}`;
    expect(
      validateXlsxCellOperations(
        [
          {
            count: 1,
            ifMatch: 'b'.repeat(64),
            index: readerLimits.maxRowsPerWorksheet,
            kind: 'delete-rows',
            operationId: 'rows',
            sheetKey,
          },
          {
            count: readerLimits.maxColumnsPerWorksheet,
            index: 1,
            kind: 'insert-columns',
            operationId: 'columns',
            sheetKey,
          },
        ],
        writeLimits,
        readerLimits,
      ),
    ).toEqual([
      {
        count: 1,
        ifMatch: 'b'.repeat(64),
        index: readerLimits.maxRowsPerWorksheet,
        kind: 'delete-rows',
        operationId: 'rows',
        sheetKey,
      },
      {
        count: readerLimits.maxColumnsPerWorksheet,
        index: 1,
        kind: 'insert-columns',
        operationId: 'columns',
        sheetKey,
      },
    ]);
    for (const [fields, message] of [
      [
        { count: 1, kind: 'insert-rows' },
        'XLSX structural operation shape is invalid',
      ],
      [
        { count: 1, index: 0, kind: 'insert-rows' },
        'XLSX structural operation index is invalid',
      ],
      [
        { count: 0, index: 1, kind: 'delete-columns' },
        'XLSX structural operation count is invalid',
      ],
      [
        {
          count: readerLimits.maxColumnsPerWorksheet + 1,
          index: 1,
          kind: 'insert-columns',
        },
        'XLSX structural operation count is invalid',
      ],
      [
        {
          count: 2,
          index: readerLimits.maxRowsPerWorksheet,
          kind: 'delete-rows',
        },
        'XLSX structural operation range is invalid',
      ],
    ] as const) {
      expect(
        capture(() =>
          validateXlsxCellOperations(
            [{ ...fields, operationId: 'structural', sheetKey }],
            writeLimits,
            readerLimits,
          ),
        ).diagnostic.message,
      ).toBe(message);
    }
  });

  it.each(ERROR_CODES)('accepts typed error value %s', (code) => {
    const operation = validateXlsxCellOperations(
      [
        {
          cell: 'A1',
          content: { kind: 'value', value: { code, kind: 'error' } },
          kind: 'set-cell',
          operationId: 'one',
          sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
        },
      ],
      writeLimits,
      readerLimits,
    )[0];
    expect(operation).toMatchObject({
      content: { value: { code, kind: 'error' } },
    });
  });

  it.each(UNSUPPORTED_OPERATION_KINDS)(
    'reports recognized unsupported operation %s',
    (kind) => {
      const error = capture(() =>
        validateXlsxCellOperations(
          [
            {
              cell: 'A1',
              kind,
              operationId: 'one',
              sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
            },
          ],
          writeLimits,
          readerLimits,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        code: 'unsupported-edit-operation',
        featureClass: kind,
        message: `XLSX operation ${kind} is not supported by this profile`,
        operationId: 'one',
      });
    },
  );

  it('validates and normalizes row and column property operations', () => {
    const sheetKey = `xlsx:sheet:${'a'.repeat(32)}`;
    expect(
      validateXlsxCellOperations(
        [
          {
            height: -0,
            hidden: false,
            ifMatch: 'b'.repeat(64),
            kind: 'set-row',
            operationId: 'row',
            row: readerLimits.maxRowsPerWorksheet,
            sheetKey,
          },
          {
            end: readerLimits.maxColumnsPerWorksheet,
            hidden: true,
            kind: 'set-column',
            operationId: 'column',
            sheetKey,
            start: 1,
            width: 255,
          },
        ],
        writeLimits,
        readerLimits,
      ),
    ).toEqual([
      {
        height: 0,
        hidden: false,
        ifMatch: 'b'.repeat(64),
        kind: 'set-row',
        operationId: 'row',
        row: readerLimits.maxRowsPerWorksheet,
        sheetKey,
      },
      {
        end: readerLimits.maxColumnsPerWorksheet,
        hidden: true,
        kind: 'set-column',
        operationId: 'column',
        sheetKey,
        start: 1,
        width: 255,
      },
    ]);
  });

  it.each([
    [{ kind: 'set-row', row: 1 }, 'XLSX set-row operation shape is invalid'],
    [
      { hidden: 'yes', kind: 'set-row', row: 1 },
      'XLSX set-row hidden value is invalid',
    ],
    [
      { height: 410, kind: 'set-row', row: 1 },
      'XLSX set-row height is invalid',
    ],
    [{ height: -1, kind: 'set-row', row: 1 }, 'XLSX set-row height is invalid'],
    [
      { hidden: true, kind: 'set-row', row: 0 },
      'XLSX set-row index is invalid',
    ],
    [
      { end: 2, kind: 'set-column', start: 1 },
      'XLSX set-column operation shape is invalid',
    ],
    [
      { end: 2, hidden: 'yes', kind: 'set-column', start: 1 },
      'XLSX set-column hidden value is invalid',
    ],
    [
      { end: 2, kind: 'set-column', start: 1, width: 256 },
      'XLSX set-column width is invalid',
    ],
    [
      { end: 2, kind: 'set-column', start: 1, width: -1 },
      'XLSX set-column width is invalid',
    ],
    [
      { end: 1, hidden: true, kind: 'set-column', start: 2 },
      'XLSX set-column range is invalid',
    ],
    [
      { end: 1, hidden: true, kind: 'set-column', start: 0 },
      'XLSX set-column start is invalid',
    ],
    [
      {
        end: readerLimits.maxColumnsPerWorksheet + 1,
        hidden: true,
        kind: 'set-column',
        start: 1,
      },
      'XLSX set-column end is invalid',
    ],
  ] as const)(
    'rejects invalid row or column operation %#',
    (fields, message) => {
      expect(
        capture(() =>
          validateXlsxCellOperations(
            [
              {
                ...fields,
                operationId: 'property',
                sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
              },
            ],
            writeLimits,
            readerLimits,
          ),
        ).diagnostic.message,
      ).toBe(message);
    },
  );

  it('requires every operation field while allowing only ifMatch as optional', () => {
    const clear: Record<string, unknown> = {
      cell: 'A1',
      kind: 'clear-cell',
      operationId: 'one',
      sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
    };
    for (const key of ['cell', 'sheetKey']) {
      const operation = { ...clear };
      delete operation[key];
      expect(
        capture(() =>
          validateXlsxCellOperations([operation], writeLimits, readerLimits),
        ).diagnostic.message,
      ).toBe('XLSX clear-cell operation shape is invalid');
    }
    const set = {
      ...clear,
      content: { kind: 'value', value: { kind: 'number', value: 1 } },
      kind: 'set-cell',
    };
    for (const key of ['cell', 'content', 'sheetKey']) {
      const operation = { ...set } as Record<string, unknown>;
      delete operation[key];
      expect(
        capture(() =>
          validateXlsxCellOperations([operation], writeLimits, readerLimits),
        ).diagnostic.message,
      ).toBe('XLSX set-cell operation shape is invalid');
    }
    const style = { ...clear, kind: 'set-cell-style', style: {} };
    for (const key of ['cell', 'sheetKey', 'style']) {
      const operation = { ...style } as Record<string, unknown>;
      delete operation[key];
      expect(
        capture(() =>
          validateXlsxCellOperations([operation], writeLimits, readerLimits),
        ).diagnostic.message,
      ).toBe('XLSX set-cell-style operation shape is invalid');
    }
    const withMatch = validateXlsxCellOperations(
      [{ ...clear, ifMatch: 'b'.repeat(64) }],
      writeLimits,
      readerLimits,
    )[0];
    const withoutMatch = validateXlsxCellOperations(
      [clear],
      writeLimits,
      readerLimits,
    )[0];
    expect(withMatch).toHaveProperty('ifMatch', 'b'.repeat(64));
    expect(withoutMatch).not.toHaveProperty('ifMatch');
    expect(
      validateXlsxCellOperations(
        [{ ...style, ifMatch: 'c'.repeat(64) }],
        writeLimits,
        readerLimits,
      )[0],
    ).toHaveProperty('ifMatch', 'c'.repeat(64));
  });

  it.each([
    [
      { kind: 'value', value: { kind: 'text', runs: [], text: '' } },
      'rich-text',
    ],
    [
      {
        kind: 'value',
        value: {
          kind: 'date',
          normalized: '2024-01-01',
          precision: 'date',
          source: { kind: 'iso', value: '2024-01-01' },
        },
      },
      'date-value',
    ],
  ] as const)(
    'blocks valid but unsupported set-cell payload %#',
    (content, featureClass) => {
      const error = capture(() =>
        validateXlsxCellOperations(
          [
            {
              cell: 'A1',
              content,
              kind: 'set-cell',
              operationId: 'one',
              sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
            },
          ],
          writeLimits,
          readerLimits,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        code: 'unsupported-edit-operation',
        featureClass,
        operationId: 'one',
      });
      expect(error.diagnostic.message).toBe(
        featureClass === 'rich-text'
          ? 'XLSX cell editing does not yet support rich text runs'
          : 'XLSX cell editing does not yet support date values',
      );
    },
  );

  it('distinguishes malformed cell and style payloads', () => {
    const base = {
      cell: 'A1',
      kind: 'set-cell',
      operationId: 'one',
      sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
    };
    for (const content of [
      null,
      {},
      { kind: 'formula', expression: '' },
      { kind: 'formula', expression: '=A1' },
      { kind: 'value' },
    ]) {
      expect(
        capture(() =>
          validateXlsxCellOperations(
            [{ ...base, content }],
            writeLimits,
            readerLimits,
          ),
        ).diagnostic.code,
      ).toBe('invalid-roundtrip-json');
    }
    for (const style of [null, [], { extra: true }]) {
      expect(
        capture(() =>
          validateXlsxCellOperations(
            [
              {
                cell: 'A1',
                kind: 'set-cell-style',
                operationId: 'one',
                sheetKey: base.sheetKey,
                style,
              },
            ],
            writeLimits,
            readerLimits,
          ),
        ).diagnostic,
      ).toMatchObject({
        code: 'invalid-roundtrip-json',
        message: 'XLSX set-cell-style style shape is invalid',
        operationId: 'one',
      });
    }
    expect(
      validateXlsxCellOperations(
        [
          {
            cell: 'A1',
            kind: 'set-cell-style',
            operationId: 'one',
            sheetKey: base.sheetKey,
            style: {},
          },
        ],
        writeLimits,
        readerLimits,
      ),
    ).toEqual([
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'one',
        sheetKey: base.sheetKey,
        style: {},
      },
    ]);
  });

  it.each([
    [undefined, 'XLSX set-hyperlink target shape is invalid'],
    [{}, 'XLSX hyperlink target kind is invalid'],
    [
      { kind: 'internal', location: '' },
      'XLSX internal hyperlink target is invalid',
    ],
    [
      { extra: true, kind: 'internal', location: 'A1' },
      'XLSX internal hyperlink target is invalid',
    ],
    [
      { kind: 'external', url: 'javascript:alert(1)' },
      'XLSX external hyperlink protocol or lexical form is unsafe',
    ],
    [
      { kind: 'external', url: ' https://example.invalid/' },
      'XLSX external hyperlink protocol or lexical form is unsafe',
    ],
    [
      { kind: 'external', url: 'https://user:secret@example.invalid/' },
      'XLSX external hyperlink credentials are not allowed',
    ],
    [
      { kind: 'external', url: 'https://user@example.invalid/' },
      'XLSX external hyperlink credentials are not allowed',
    ],
    [
      { kind: 'external', url: 'https://:secret@example.invalid/' },
      'XLSX external hyperlink credentials are not allowed',
    ],
    [
      { kind: 'external', url: 'https://example.invalid' },
      'XLSX external hyperlink URL must be canonical',
    ],
    [
      { kind: 'external', location: 1, url: 'https://example.invalid/' },
      'XLSX external hyperlink target is invalid',
    ],
  ] as const)('rejects invalid hyperlink target %#', (target, message) => {
    const error = capture(() =>
      validateXlsxCellOperations(
        [
          {
            cell: 'A1',
            kind: 'set-hyperlink',
            operationId: 'link',
            sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
            target,
          },
        ],
        writeLimits,
        readerLimits,
      ),
    );
    expect(error.diagnostic).toMatchObject({
      message,
      operationId: 'link',
      ...(message.includes('protocol')
        ? { featureClass: 'hyperlink-protocol' }
        : message.includes('credentials')
          ? { featureClass: 'hyperlink-credentials' }
          : {}),
    });
  });

  it.each([
    [
      {
        cell: 'A1',
        kind: 'set-hyperlink',
        operationId: 'link',
        sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
      },
    ],
    [
      {
        cell: 'A1',
        extra: true,
        kind: 'set-hyperlink',
        operationId: 'link',
        sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
        target: null,
      },
    ],
  ])('rejects invalid set-hyperlink operation shape %#', (operation) => {
    expect(
      capture(() =>
        validateXlsxCellOperations([operation], writeLimits, readerLimits),
      ).diagnostic.message,
    ).toBe('XLSX set-hyperlink operation shape is invalid');
  });

  it.each([
    [null, 'XLSX cell value shape is invalid'],
    [undefined, 'XLSX cell value shape is invalid'],
    [1, 'XLSX cell value shape is invalid'],
    [[], 'XLSX cell value shape is invalid'],
    [Object.create(null) as unknown, 'XLSX cell value shape is invalid'],
    [{ kind: 'text' }, 'XLSX text cell value shape is invalid'],
    [
      { extra: true, kind: 'text', text: 'x' },
      'XLSX text cell value shape is invalid',
    ],
    [{ kind: 'text', text: 1 }, 'XLSX text cell value is invalid'],
    [{ kind: 'number' }, 'XLSX number cell value is invalid'],
    [{ kind: 'number', value: '1' }, 'XLSX number cell value is invalid'],
    [{ kind: 'number', value: Infinity }, 'XLSX number cell value is invalid'],
    [{ kind: 'boolean' }, 'XLSX boolean cell value is invalid'],
    [{ kind: 'boolean', value: 0 }, 'XLSX boolean cell value is invalid'],
    [{ code: '#BAD!', kind: 'error' }, 'XLSX error cell value is invalid'],
    [{ kind: 'error' }, 'XLSX error cell value is invalid'],
    [{ kind: 'unknown' }, 'XLSX cell value kind is invalid'],
  ] as const)('rejects malformed cell value %#', (value, message) => {
    expect(
      capture(() =>
        validateXlsxCellOperations(
          [
            {
              cell: 'A1',
              content: { kind: 'value', value },
              kind: 'set-cell',
              operationId: 'one',
              sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
            },
          ],
          writeLimits,
          readerLimits,
        ),
      ).diagnostic.message,
    ).toBe(message);
  });

  it.each([
    [undefined, 'XLSX set-cell content shape is invalid'],
    [1, 'XLSX set-cell content shape is invalid'],
    [[], 'XLSX set-cell content shape is invalid'],
    [Object.create(null) as unknown, 'XLSX set-cell content shape is invalid'],
    [{ kind: 'formula' }, 'XLSX set-cell formula is invalid'],
    [{ expression: 1, kind: 'formula' }, 'XLSX set-cell formula is invalid'],
    [
      { expression: 'A1', extra: true, kind: 'formula' },
      'XLSX set-cell formula is invalid',
    ],
    [
      { extra: true, kind: 'value', value: { kind: 'number', value: 1 } },
      'XLSX set-cell value content shape is invalid',
    ],
    [
      { kind: 'value', other: true },
      'XLSX set-cell value content shape is invalid',
    ],
    [{ kind: 'other' }, 'XLSX set-cell content kind is invalid'],
  ] as const)('rejects malformed set-cell content %#', (content, message) => {
    expect(
      capture(() =>
        validateXlsxCellOperations(
          [
            {
              cell: 'A1',
              content,
              kind: 'set-cell',
              operationId: 'one',
              sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
            },
          ],
          writeLimits,
          readerLimits,
        ),
      ).diagnostic.message,
    ).toBe(message);
  });

  it('rejects duplicate operation IDs before replay', () => {
    const operation = {
      cell: 'A1',
      kind: 'clear-cell',
      operationId: 'same',
      sheetKey: `xlsx:sheet:${'a'.repeat(32)}`,
    };
    const error = capture(() =>
      validateXlsxCellOperations(
        [operation, operation],
        writeLimits,
        readerLimits,
      ),
    );
    expect(error.diagnostic).toMatchObject({
      code: 'invalid-roundtrip-json',
      operationId: 'same',
    });
    expect(error.diagnostic.message).toBe('XLSX operation IDs must be unique');
  });

  it('enforces operation, formula, and text budgets at exact boundaries', () => {
    const sheetKey = `xlsx:sheet:${'a'.repeat(32)}`;
    const clear = {
      cell: 'A1',
      kind: 'clear-cell',
      operationId: 'one',
      sheetKey,
    };
    const exactBytes = new TextEncoder().encode(
      JSON.stringify(clear, Object.keys(clear).sort()),
    ).byteLength;
    expect(() =>
      validateXlsxCellOperations(
        [clear],
        { ...writeLimits, maxOperationBytes: exactBytes, maxOperations: 1 },
        readerLimits,
      ),
    ).not.toThrow();
    expect(
      capture(() =>
        validateXlsxCellOperations(
          [clear],
          { ...writeLimits, maxOperationBytes: exactBytes - 1 },
          readerLimits,
        ),
      ).diagnostic,
    ).toMatchObject({
      limitName: 'maxOperationBytes',
      message: 'XLSX operation exceeds its byte limit',
    });
    expect(
      capture(() =>
        validateXlsxCellOperations(
          [clear, { ...clear, operationId: 'two' }],
          { ...writeLimits, maxOperations: 1 },
          readerLimits,
        ),
      ).diagnostic.limitName,
    ).toBe('maxOperations');
    const secondClear = { ...clear, operationId: 'two' };
    const secondBytes = new TextEncoder().encode(
      canonicalXlsxJson(secondClear),
    ).byteLength;
    expect(() =>
      validateXlsxCellOperations(
        [clear, secondClear],
        {
          ...writeLimits,
          maxTotalOperationBytes: exactBytes + secondBytes,
        },
        readerLimits,
      ),
    ).not.toThrow();
    const totalByteError = capture(() =>
      validateXlsxCellOperations(
        [clear, secondClear],
        {
          ...writeLimits,
          maxTotalOperationBytes: exactBytes + secondBytes - 1,
        },
        readerLimits,
      ),
    );
    expect(totalByteError.diagnostic).toMatchObject({
      actual: exactBytes + secondBytes,
      code: 'resource-limit-exceeded',
      limit: exactBytes + secondBytes - 1,
      limitName: 'maxTotalOperationBytes',
      message: 'XLSX operations exceed their total byte limit',
      operationId: 'two',
    });

    const formula = {
      ...clear,
      content: { kind: 'formula', expression: 'AB' },
      kind: 'set-cell',
    };
    expect(() =>
      validateXlsxCellOperations([formula], writeLimits, {
        ...readerLimits,
        maxFormulaCharacters: 2,
      }),
    ).not.toThrow();
    expect(
      capture(() =>
        validateXlsxCellOperations([formula], writeLimits, {
          ...readerLimits,
          maxFormulaCharacters: 1,
        }),
      ).diagnostic,
    ).toMatchObject({
      limitName: 'maxFormulaCharacters',
      message: 'XLSX operation formula exceeds its character limit',
    });
    const formulaTwo = { ...formula, operationId: 'two' };
    expect(() =>
      validateXlsxCellOperations([formula, formulaTwo], writeLimits, {
        ...readerLimits,
        maxFormulaCharacters: 2,
        maxTotalFormulaCharacters: 4,
      }),
    ).not.toThrow();
    const totalFormulaError = capture(() =>
      validateXlsxCellOperations([formula, formulaTwo], writeLimits, {
        ...readerLimits,
        maxFormulaCharacters: 2,
        maxTotalFormulaCharacters: 3,
      }),
    );
    expect(totalFormulaError.diagnostic).toMatchObject({
      actual: 4,
      limit: 3,
      limitName: 'maxTotalFormulaCharacters',
      message: 'XLSX operations exceed their total formula character limit',
      operationId: 'two',
    });

    const text = {
      ...clear,
      content: { kind: 'value', value: { kind: 'text', text: 'éé' } },
      kind: 'set-cell',
    };
    expect(() =>
      validateXlsxCellOperations([text], writeLimits, {
        ...readerLimits,
        maxTextCharacters: 2,
      }),
    ).not.toThrow();
    expect(
      capture(() =>
        validateXlsxCellOperations([text], writeLimits, {
          ...readerLimits,
          maxTextCharacters: 1,
        }),
      ).diagnostic.limitName,
    ).toBe('maxTextCharacters');
    const singleCharacterText = {
      ...text,
      content: { kind: 'value', value: { kind: 'text', text: 'x' } },
    };
    const secondText = { ...singleCharacterText, operationId: 'two' };
    expect(() =>
      validateXlsxCellOperations(
        [singleCharacterText, secondText],
        writeLimits,
        { ...readerLimits, maxTextCharacters: 2 },
      ),
    ).not.toThrow();
    const totalTextError = capture(() =>
      validateXlsxCellOperations(
        [singleCharacterText, secondText],
        writeLimits,
        { ...readerLimits, maxTextCharacters: 1 },
      ),
    );
    expect(totalTextError.diagnostic).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxTextCharacters',
      message: 'XLSX operations exceed their text character limit',
      operationId: 'two',
    });
    const hyperlink = {
      ...clear,
      kind: 'set-hyperlink',
      target: { kind: 'internal', location: 'ab' },
    };
    expect(() =>
      validateXlsxCellOperations([hyperlink], writeLimits, {
        ...readerLimits,
        maxTextCharacters: 2,
      }),
    ).not.toThrow();
    expect(
      capture(() =>
        validateXlsxCellOperations([hyperlink], writeLimits, {
          ...readerLimits,
          maxTextCharacters: 1,
        }),
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxTextCharacters',
      message: 'XLSX hyperlink location exceeds its text limit',
    });
    const firstLink = {
      ...hyperlink,
      target: { kind: 'internal', location: 'x' },
    };
    const secondLink = { ...firstLink, operationId: 'two' };
    expect(
      capture(() =>
        validateXlsxCellOperations([firstLink, secondLink], writeLimits, {
          ...readerLimits,
          maxTextCharacters: 1,
        }),
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      limit: 1,
      message: 'XLSX operations exceed their text character limit',
      operationId: 'two',
    });
    const externalLink = {
      ...hyperlink,
      target: {
        kind: 'external',
        location: 'xy',
        url: 'https://e.co/',
      },
    };
    const externalCharacters = 'https://e.co/'.length + 2;
    expect(() =>
      validateXlsxCellOperations([externalLink], writeLimits, {
        ...readerLimits,
        maxTextCharacters: externalCharacters,
      }),
    ).not.toThrow();
    expect(
      capture(() =>
        validateXlsxCellOperations([externalLink], writeLimits, {
          ...readerLimits,
          maxTextCharacters: externalCharacters - 1,
        }),
      ).diagnostic,
    ).toMatchObject({
      actual: externalCharacters,
      limit: externalCharacters - 1,
      message: 'XLSX hyperlink target exceeds its text limit',
    });
    const externalTwo = { ...externalLink, operationId: 'two' };
    expect(
      capture(() =>
        validateXlsxCellOperations([externalLink, externalTwo], writeLimits, {
          ...readerLimits,
          maxTextCharacters: externalCharacters,
        }),
      ).diagnostic,
    ).toMatchObject({
      actual: externalCharacters * 2,
      limit: externalCharacters,
      operationId: 'two',
    });
  });
});

describe('XLSX cell operation planner', () => {
  it('replays ordered set and clear operations without mutating inputs', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const before = JSON.stringify(snapshot.document);
    const operations: XlsxEditOperation[] = [
      cellOperation(snapshot.document, {
        content: { kind: 'value', value: { kind: 'number', value: 7 } },
      }),
      {
        cell: 'A1',
        kind: 'clear-cell',
        operationId: 'edit-2',
        sheetKey: worksheet(snapshot.document).key,
      },
      cellOperation(snapshot.document, {
        content: { kind: 'formula', expression: '1+2' },
        operationId: 'edit-3',
      }),
    ];
    const plan = await replayXlsxCellOperations(
      snapshot.document,
      operations,
      writeLimits,
      readerLimits,
    );
    expect(JSON.stringify(snapshot.document)).toBe(before);
    expect(plan.operations).not.toBe(operations);
    expect(plan.impacts).toEqual([
      {
        cell: 'A1',
        kind: 'set-cell',
        operationId: 'edit-1',
        sheetKey: worksheet(snapshot.document).key,
      },
      {
        cell: 'A1',
        kind: 'clear-cell',
        operationId: 'edit-2',
        sheetKey: worksheet(snapshot.document).key,
      },
      {
        cell: 'A1',
        kind: 'set-cell',
        operationId: 'edit-3',
        sheetKey: worksheet(snapshot.document).key,
      },
    ]);
    expect(worksheet(plan.document).rows[0]?.cells[0]?.content).toEqual({
      cached: { kind: 'missing' },
      formula: { expression: '1+2', kind: 'normal' },
      kind: 'formula',
    });
    expect(plan.stateHash).toBe(await canonicalXlsxSha256(plan.document));
  });

  it('replays row and column properties with target-specific preconditions', async () => {
    const snapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><cols><col min="1" max="2" width="10" customWidth="1"/></cols><sheetData><row r="1" ht="12" customHeight="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
      }),
    );
    const source = worksheet(snapshot.document);
    const rowMatch = await canonicalXlsxSha256(
      xlsxRowTargetState(source, source.rows[0]!),
    );
    const columnMatch = await canonicalXlsxSha256(
      xlsxColumnTargetState(source, source.columns[0]!),
    );
    const plan = await replayXlsxCellOperations(
      snapshot.document,
      [
        {
          height: 20,
          hidden: true,
          ifMatch: rowMatch,
          kind: 'set-row',
          operationId: 'row',
          row: 1,
          sheetKey: source.key,
        },
        {
          end: 2,
          hidden: false,
          ifMatch: columnMatch,
          kind: 'set-column',
          operationId: 'column',
          sheetKey: source.key,
          start: 1,
          width: 25,
        },
      ],
      writeLimits,
      readerLimits,
    );
    expect(source.rows[0]).toMatchObject({ height: 12, index: 1 });
    expect(source.columns[0]).toMatchObject({ start: 1, width: 10 });
    expect(worksheet(plan.document).rows[0]).toMatchObject({
      height: 20,
      hidden: true,
      index: 1,
    });
    expect(worksheet(plan.document).columns[0]).toMatchObject({
      end: 2,
      hidden: false,
      start: 1,
      width: 25,
    });
    expect(plan.impacts).toEqual([
      {
        kind: 'set-row',
        operationId: 'row',
        range: '1',
        sheetKey: source.key,
      },
      {
        kind: 'set-column',
        operationId: 'column',
        range: '1:2',
        sheetKey: source.key,
      },
    ]);
    expect(plan.stateHash).toBe(await canonicalXlsxSha256(plan.document));

    for (const operation of [
      {
        hidden: true,
        kind: 'set-row' as const,
        operationId: 'missing-row',
        row: 2,
        sheetKey: source.key,
      },
      {
        end: 1,
        hidden: true,
        kind: 'set-column' as const,
        operationId: 'missing-column',
        sheetKey: source.key,
        start: 1,
      },
    ]) {
      const error = await captureAsync(() =>
        replayXlsxCellOperations(
          snapshot.document,
          [operation],
          writeLimits,
          readerLimits,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        code: 'preservation-conflict',
        featureClass:
          operation.kind === 'set-row' ? 'missing-row' : 'missing-column-range',
        message:
          operation.kind === 'set-row'
            ? 'XLSX set-row operation requires an existing explicit source row'
            : 'XLSX set-column operation requires an existing exact source range',
        operationId: operation.operationId,
        range: operation.kind === 'set-row' ? '2' : '1:1',
      });
    }
    const shiftedColumns = structuredClone(snapshot.document);
    worksheet(shiftedColumns).columns = [{ end: 2, start: 2 }];
    expect(
      (
        await captureAsync(() =>
          replayXlsxCellOperations(
            shiftedColumns,
            [
              {
                end: 2,
                hidden: true,
                kind: 'set-column',
                operationId: 'wrong-start',
                sheetKey: source.key,
                start: 1,
              },
            ],
            writeLimits,
            readerLimits,
          ),
        )
      ).diagnostic.featureClass,
    ).toBe('missing-column-range');
    const mismatch = await captureAsync(() =>
      replayXlsxCellOperations(
        snapshot.document,
        [
          {
            hidden: true,
            ifMatch: '0'.repeat(64),
            kind: 'set-row',
            operationId: 'row-precondition',
            row: 1,
            sheetKey: source.key,
          },
        ],
        writeLimits,
        readerLimits,
      ),
    );
    expect(mismatch.diagnostic).toMatchObject({
      code: 'operation-precondition-failed',
      message: 'XLSX operation precondition does not match the target row',
      range: '1',
    });
    const columnMismatch = await captureAsync(() =>
      replayXlsxCellOperations(
        snapshot.document,
        [
          {
            end: 2,
            hidden: true,
            ifMatch: '0'.repeat(64),
            kind: 'set-column',
            operationId: 'column-precondition',
            sheetKey: source.key,
            start: 1,
          },
        ],
        writeLimits,
        readerLimits,
      ),
    );
    expect(columnMismatch.diagnostic).toMatchObject({
      code: 'operation-precondition-failed',
      message:
        'XLSX operation precondition does not match the target column range',
      range: '1:2',
    });
  });

  it('replays structural shifts only for a reference-free worksheet', async () => {
    const snapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="3"><c r="C3"><v>3</v></c></row></sheetData></worksheet>`,
      }),
    );
    const source = worksheet(snapshot.document);
    const ifMatch = await canonicalXlsxSha256(
      xlsxStructuralTargetState(source),
    );
    const operations: XlsxEditOperation[] = [
      {
        count: 2,
        ifMatch,
        index: 2,
        kind: 'insert-rows',
        operationId: 'insert-rows',
        sheetKey: source.key,
      },
      {
        count: 1,
        index: 1,
        kind: 'delete-rows',
        operationId: 'delete-rows',
        sheetKey: source.key,
      },
      {
        count: 1,
        index: 2,
        kind: 'insert-columns',
        operationId: 'insert-columns',
        sheetKey: source.key,
      },
      {
        count: 1,
        index: 1,
        kind: 'delete-columns',
        operationId: 'delete-columns',
        sheetKey: source.key,
      },
    ];
    const plan = await replayXlsxCellOperations(
      snapshot.document,
      operations,
      writeLimits,
      readerLimits,
    );
    expect(source.rows.map((row) => row.index)).toEqual([1, 3]);
    expect(worksheet(plan.document).rows).toEqual([
      expect.objectContaining({
        cells: [expect.objectContaining({ address: 'C4', column: 3 })],
        index: 4,
      }),
    ]);
    expect(
      plan.impacts.map((impact) => [
        impact.kind,
        'range' in impact ? impact.range : impact.cell,
      ]),
    ).toEqual([
      ['insert-rows', '2:3'],
      ['delete-rows', '1:1'],
      ['insert-columns', '2:2'],
      ['delete-columns', '1:1'],
    ]);
    await expect(
      replayXlsxCellOperations(
        snapshot.document,
        operations,
        { ...writeLimits, maxReferenceUpdates: 9 },
        readerLimits,
      ),
    ).resolves.toMatchObject({ impacts: plan.impacts });
    expect(
      (
        await captureAsync(() =>
          replayXlsxCellOperations(
            snapshot.document,
            operations,
            { ...writeLimits, maxReferenceUpdates: 8 },
            readerLimits,
          ),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 9,
      limit: 8,
      limitName: 'maxReferenceUpdates',
    });

    const referencedDocument: XlsxRoundTripDocument = {
      ...snapshot.document,
      sheets: [
        {
          ...source,
          declaredDimension: {
            end: { column: 3, row: 3 },
            reference: 'A1:C3',
            start: { column: 1, row: 1 },
          },
          mergedRanges: [
            {
              end: { column: 2, row: 3 },
              reference: 'A1:B3',
              start: { column: 1, row: 1 },
            },
          ],
          hyperlinks: [
            {
              range: {
                end: { column: 2, row: 3 },
                reference: 'A1:B3',
                start: { column: 1, row: 1 },
              },
              selectionRelation: 'full-sheet',
              target: { kind: 'internal', location: 'Sheet1!A1' },
            },
          ],
          autoFilter: {
            columns: [],
            range: {
              end: { column: 3, row: 3 },
              reference: 'A1:C3',
              start: { column: 1, row: 1 },
            },
            selectionRelation: 'full-sheet',
            sort: {
              caseSensitive: false,
              columnSort: false,
              conditions: [
                {
                  descending: false,
                  range: {
                    end: { column: 2, row: 3 },
                    reference: 'B2:B3',
                    start: { column: 2, row: 2 },
                  },
                  sortBy: 'value',
                },
              ],
              range: {
                end: { column: 3, row: 3 },
                reference: 'A1:C3',
                start: { column: 1, row: 1 },
              },
              sortMethod: 'none',
            },
          },
          dataValidationSettings: { disablePrompts: true },
          dataValidations: [
            {
              allowBlank: false,
              errorStyle: 'stop',
              imeMode: 'no-control',
              operator: 'between',
              ranges: [
                {
                  end: { column: 1, row: 3 },
                  reference: 'A1:A3',
                  start: { column: 1, row: 1 },
                },
                {
                  end: { column: 3, row: 3 },
                  reference: 'C2:C3',
                  start: { column: 3, row: 2 },
                },
              ],
              selectionRelation: 'full-sheet',
              showDropDown: false,
              showErrorMessage: false,
              showInputMessage: false,
              type: 'none',
            },
          ],
          conditionalFormattings: [
            {
              pivot: false,
              ranges: [
                {
                  end: { column: 4, row: 3 },
                  reference: 'D1:D3',
                  start: { column: 4, row: 1 },
                },
                {
                  end: { column: 5, row: 3 },
                  reference: 'E2:E3',
                  start: { column: 5, row: 2 },
                },
              ],
              rules: [
                {
                  formulas: [],
                  priority: 1,
                  stopIfTrue: false,
                  type: 'top',
                },
                {
                  colorScale: {
                    stops: [
                      {
                        color: {} as never,
                        threshold: {
                          greaterThanOrEqual: true,
                          kind: 'minimum',
                        },
                      },
                    ],
                  },
                  formulas: [],
                  priority: 2,
                  stopIfTrue: false,
                  type: 'color-scale',
                },
                {
                  dataBar: {
                    color: {} as never,
                    maximumLength: 100,
                    minimumLength: 0,
                    showValue: true,
                    thresholds: [
                      { greaterThanOrEqual: true, kind: 'minimum' },
                      { greaterThanOrEqual: true, kind: 'maximum' },
                    ],
                  },
                  formulas: [],
                  priority: 3,
                  stopIfTrue: false,
                  type: 'data-bar',
                },
                {
                  formulas: [],
                  iconSet: {
                    iconSet: '3Arrows',
                    percent: false,
                    reverse: false,
                    showValue: true,
                    thresholds: [
                      {
                        greaterThanOrEqual: true,
                        kind: 'number',
                        value: 0,
                      },
                    ],
                  },
                  priority: 4,
                  stopIfTrue: false,
                  type: 'icon-set',
                },
              ],
              selectionRelation: 'full-sheet',
            },
          ],
          comments: [
            {
              author: 'Author',
              kind: 'note',
              reference: 'J10',
              selectionRelation: 'full-sheet',
              text: 'Note',
              visible: false,
            },
            {
              id: 'thread',
              kind: 'threaded',
              personId: 'person',
              reference: 'K11',
              selectionRelation: 'full-sheet',
              text: 'Thread',
              timestamp: '2024-01-01T00:00:00Z',
            },
          ],
          drawings: [
            structuralDrawing({
              from: {
                column: 12,
                columnOffset: 0,
                row: 20,
                rowOffset: 0,
              },
            }),
            structuralDrawing({
              from: {
                column: 13,
                columnOffset: 0,
                row: 1,
                rowOffset: 0,
              },
              kind: 'two-cell',
              to: {
                column: 14,
                columnOffset: 0,
                row: 20,
                rowOffset: 0,
              },
            }),
          ],
          sparklineGroups: [structuralSparklineGroup()],
          protectedRanges: [
            {
              name: 'Input',
              ranges: [
                {
                  end: { column: 6, row: 3 },
                  reference: 'F1:F3',
                  start: { column: 6, row: 1 },
                },
                {
                  end: { column: 7, row: 3 },
                  reference: 'G2:G3',
                  start: { column: 7, row: 2 },
                },
              ],
              selectionRelation: 'full-sheet',
            },
          ],
          print: {
            columnBreaks: [
              {
                end: 2,
                manual: false,
                pivot: true,
                position: 3,
                start: 0,
              },
            ],
            rowBreaks: [
              {
                end: 2,
                manual: true,
                pivot: false,
                position: 3,
                start: 0,
              },
            ],
          },
          tables: [structuralTable()],
          views: [
            {
              kind: 'normal',
              rightToLeft: false,
              selections: [
                {
                  activeCell: 'I2',
                  activeCellId: 1,
                  pane: 'top-left',
                  ranges: [
                    {
                      end: { column: 8, row: 3 },
                      reference: 'H1:H3',
                      start: { column: 8, row: 1 },
                    },
                    {
                      end: { column: 9, row: 3 },
                      reference: 'I2:I3',
                      start: { column: 9, row: 2 },
                    },
                  ],
                },
              ],
              showGridLines: true,
              showRowColumnHeaders: true,
              tabSelected: false,
              topLeftCell: 'A2',
              workbookViewId: 0,
              zoomScale: 100,
            },
          ],
        },
      ],
    };
    const layoutOperation: XlsxEditOperation = {
      count: 2,
      index: 2,
      kind: 'insert-rows',
      operationId: 'insert-layout-rows',
      sheetKey: source.key,
    };
    expect(
      worksheet(
        (
          await replayXlsxCellOperations(
            referencedDocument,
            [layoutOperation],
            writeLimits,
            readerLimits,
          )
        ).document,
      ).declaredDimension?.reference,
    ).toBe('A1:C5');
    expect(
      worksheet(
        (
          await replayXlsxCellOperations(
            referencedDocument,
            [layoutOperation],
            writeLimits,
            readerLimits,
          )
        ).document,
      ).mergedRanges[0]?.reference,
    ).toBe('A1:B5');
    const transformedLayout = worksheet(
      (
        await replayXlsxCellOperations(
          referencedDocument,
          [layoutOperation],
          writeLimits,
          readerLimits,
        )
      ).document,
    );
    expect(transformedLayout.autoFilter?.range.reference).toBe('A1:C5');
    expect(transformedLayout.autoFilter?.sort?.range.reference).toBe('A1:C5');
    expect(
      transformedLayout.autoFilter?.sort?.conditions[0]?.range.reference,
    ).toBe('B4:B5');
    expect(
      transformedLayout.dataValidations[0]?.ranges.map(
        (range) => range.reference,
      ),
    ).toEqual(['A1:A5', 'C4:C5']);
    expect(
      transformedLayout.conditionalFormattings[0]?.ranges.map(
        (range) => range.reference,
      ),
    ).toEqual(['D1:D5', 'E4:E5']);
    expect(
      transformedLayout.protectedRanges[0]?.ranges.map(
        (range) => range.reference,
      ),
    ).toEqual(['F1:F5', 'G4:G5']);
    expect(transformedLayout.print?.rowBreaks?.[0]).toMatchObject({
      end: 2,
      position: 5,
      start: 0,
    });
    expect(transformedLayout.print?.columnBreaks?.[0]).toMatchObject({
      end: 4,
      position: 3,
      start: 0,
    });
    expect(transformedLayout.views[0]?.topLeftCell).toBe('A4');
    expect(transformedLayout.views[0]?.selections[0]).toMatchObject({
      activeCell: 'I4',
      activeCellId: 1,
      ranges: [{ reference: 'H1:H5' }, { reference: 'I4:I5' }],
    });
    expect(transformedLayout.tables[0]?.range.reference).toBe('A1:B5');
    expect(transformedLayout.tables[0]?.autoFilter?.range.reference).toBe(
      'A1:B5',
    );
    expect(
      transformedLayout.tables[0]?.autoFilter?.sort?.conditions[0]?.range
        .reference,
    ).toBe('A4:A5');
    expect(
      transformedLayout.comments.map((comment) => comment.reference),
    ).toEqual(['J12', 'K13']);
    expect(transformedLayout.drawings[0]?.from).toMatchObject({ row: 22 });
    expect(transformedLayout.drawings[1]).toMatchObject({
      from: { column: 13, row: 1 },
      to: { column: 14, row: 22 },
    });
    expect(transformedLayout.sparklineGroups?.[0]?.sparklines[0]).toMatchObject(
      {
        dataFormula: 'A1:A5',
        location: 'L12',
      },
    );
    const deletedSortDocument = structuredClone(referencedDocument);
    const deletedSortSheet = worksheet(deletedSortDocument);
    deletedSortSheet.autoFilter!.sort!.range = {
      end: { column: 3, row: 3 },
      reference: 'A2:C3',
      start: { column: 1, row: 2 },
    };
    const deletedSort = await replayXlsxCellOperations(
      deletedSortDocument,
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'delete-sort-range',
          sheetKey: source.key,
        },
      ],
      writeLimits,
      readerLimits,
    );
    expect(worksheet(deletedSort.document).autoFilter).not.toHaveProperty(
      'sort',
    );
    const deletedFilterDocument = structuredClone(referencedDocument);
    worksheet(deletedFilterDocument).autoFilter!.range = {
      end: { column: 3, row: 3 },
      reference: 'A2:C3',
      start: { column: 1, row: 2 },
    };
    const deletedFilter = await replayXlsxCellOperations(
      deletedFilterDocument,
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'delete-filter-range',
          sheetKey: source.key,
        },
      ],
      writeLimits,
      readerLimits,
    );
    expect(worksheet(deletedFilter.document)).not.toHaveProperty('autoFilter');
    const deletedValidationDocument = structuredClone(referencedDocument);
    const deletedValidationSheet = worksheet(deletedValidationDocument);
    deletedValidationSheet.dataValidations[0]!.ranges = [
      {
        end: { column: 1, row: 3 },
        reference: 'A2:A3',
        start: { column: 1, row: 2 },
      },
    ];
    const deletedValidation = await replayXlsxCellOperations(
      deletedValidationDocument,
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'delete-validation-range',
          sheetKey: source.key,
        },
      ],
      writeLimits,
      readerLimits,
    );
    expect(worksheet(deletedValidation.document).dataValidations).toEqual([]);
    expect(worksheet(deletedValidation.document)).not.toHaveProperty(
      'dataValidationSettings',
    );
    const deletedFormatDocument = structuredClone(referencedDocument);
    worksheet(deletedFormatDocument).conditionalFormattings[0]!.ranges = [
      {
        end: { column: 4, row: 3 },
        reference: 'D2:D3',
        start: { column: 4, row: 2 },
      },
    ];
    const deletedFormat = await replayXlsxCellOperations(
      deletedFormatDocument,
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'delete-format-range',
          sheetKey: source.key,
        },
      ],
      writeLimits,
      readerLimits,
    );
    expect(worksheet(deletedFormat.document).conditionalFormattings).toEqual(
      [],
    );
    const deletedProtectedDocument = structuredClone(referencedDocument);
    worksheet(deletedProtectedDocument).protectedRanges[0]!.ranges = [
      {
        end: { column: 6, row: 3 },
        reference: 'F2:F3',
        start: { column: 6, row: 2 },
      },
    ];
    const deletedProtected = await replayXlsxCellOperations(
      deletedProtectedDocument,
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'delete-protected-range',
          sheetKey: source.key,
        },
      ],
      writeLimits,
      readerLimits,
    );
    expect(worksheet(deletedProtected.document).protectedRanges).toEqual([]);
    const deletedBreakDocument = structuredClone(referencedDocument);
    const deletedBreakSheet = worksheet(deletedBreakDocument);
    deletedBreakSheet.print!.rowBreaks![0]!.position = 2;
    deletedBreakSheet.print!.columnBreaks![0]!.start = 1;
    deletedBreakSheet.print!.columnBreaks![0]!.end = 1;
    const deletedBreaks = await replayXlsxCellOperations(
      deletedBreakDocument,
      [
        {
          count: 1,
          index: 2,
          kind: 'delete-rows',
          operationId: 'delete-page-breaks',
          sheetKey: source.key,
        },
      ],
      writeLimits,
      readerLimits,
    );
    expect(worksheet(deletedBreaks.document).print).toEqual({});
    const emptyValidationDocument = structuredClone(snapshot.document);
    worksheet(emptyValidationDocument).dataValidationSettings = {
      disablePrompts: true,
    };
    worksheet(emptyValidationDocument).print = { rowBreaks: [] };
    expect(
      worksheet(
        (
          await replayXlsxCellOperations(
            emptyValidationDocument,
            [layoutOperation],
            writeLimits,
            readerLimits,
          )
        ).document,
      ).dataValidationSettings,
    ).toEqual({ disablePrompts: true });
    expect(
      worksheet(
        (
          await replayXlsxCellOperations(
            emptyValidationDocument,
            [layoutOperation],
            writeLimits,
            readerLimits,
          )
        ).document,
      ).print,
    ).toEqual({ rowBreaks: [] });
    expect(
      worksheet(
        (
          await replayXlsxCellOperations(
            referencedDocument,
            [layoutOperation],
            writeLimits,
            readerLimits,
          )
        ).document,
      ).hyperlinks[0]?.range.reference,
    ).toBe('A1:B5');
    await expect(
      replayXlsxCellOperations(
        referencedDocument,
        [layoutOperation],
        { ...writeLimits, maxReferenceUpdates: 33 },
        readerLimits,
      ),
    ).resolves.toBeDefined();
    expect(
      (
        await captureAsync(() =>
          replayXlsxCellOperations(
            referencedDocument,
            [layoutOperation],
            { ...writeLimits, maxReferenceUpdates: 32 },
            readerLimits,
          ),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 33,
      limit: 32,
      limitName: 'maxReferenceUpdates',
    });
    const viewOnlyDocument = structuredClone(snapshot.document);
    worksheet(viewOnlyDocument).views = [
      {
        kind: 'normal',
        rightToLeft: false,
        selections: [
          {
            pane: 'top-left',
            ranges: [
              {
                end: { column: 1, row: 1 },
                reference: 'A1',
                start: { column: 1, row: 1 },
              },
            ],
          },
        ],
        showGridLines: true,
        showRowColumnHeaders: true,
        tabSelected: false,
        workbookViewId: 0,
        zoomScale: 100,
      },
    ];
    const viewOnlyOperation: XlsxEditOperation = {
      count: 1,
      index: 4,
      kind: 'insert-rows',
      operationId: 'view-reference-limit',
      sheetKey: source.key,
    };
    await expect(
      replayXlsxCellOperations(
        viewOnlyDocument,
        [viewOnlyOperation],
        { ...writeLimits, maxReferenceUpdates: 1 },
        readerLimits,
      ),
    ).resolves.toBeDefined();
    expect(
      (
        await captureAsync(() =>
          replayXlsxCellOperations(
            viewOnlyDocument,
            [viewOnlyOperation],
            { ...writeLimits, maxReferenceUpdates: 0 },
            readerLimits,
          ),
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 1,
      limit: 0,
      limitName: 'maxReferenceUpdates',
    });
    expect(
      (
        await captureAsync(() =>
          replayXlsxCellOperations(
            snapshot.document,
            [{ ...operations[0]!, ifMatch: '0'.repeat(64) }],
            writeLimits,
            readerLimits,
          ),
        )
      ).diagnostic,
    ).toMatchObject({
      code: 'operation-precondition-failed',
      message:
        'XLSX operation precondition does not match the target worksheet',
    });
    expect(
      (
        await captureAsync(() =>
          replayXlsxCellOperations(
            snapshot.document,
            [
              operations[0]!,
              {
                hidden: true,
                kind: 'set-row',
                operationId: 'mixed-row',
                row: 1,
                sheetKey: source.key,
              },
            ],
            writeLimits,
            readerLimits,
          ),
        )
      ).diagnostic.featureClass,
    ).toBe('mixed-operation-closure');
  });

  it('keeps every structural transform boundary distinct in preview state', async () => {
    const snapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="3"><c r="C3"><v>3</v></c></row></sheetData></worksheet>`,
      }),
    );
    const sheetKey = snapshot.document.sheets[0]!.key;
    const replay = (operation: XlsxEditOperation) =>
      replayXlsxCellOperations(
        snapshot.document,
        [operation],
        writeLimits,
        readerLimits,
      );
    const rowsAndCells = (document: XlsxRoundTripDocument) =>
      worksheet(document).rows.map((row) => [
        row.index,
        row.cells.map((cell) => [
          cell.address,
          cell.content.kind === 'value' && cell.content.value.kind === 'number'
            ? cell.content.value.value
            : null,
        ]),
      ]);
    expect(
      rowsAndCells(
        (
          await replay({
            count: 2,
            index: 2,
            kind: 'insert-rows',
            operationId: 'insert-rows',
            sheetKey,
          })
        ).document,
      ),
    ).toEqual([
      [
        1,
        [
          ['A1', 1],
          ['B1', 2],
        ],
      ],
      [5, [['C5', 3]]],
    ]);
    expect(
      rowsAndCells(
        (
          await replay({
            count: 1,
            index: 1,
            kind: 'delete-rows',
            operationId: 'delete-rows',
            sheetKey,
          })
        ).document,
      ),
    ).toEqual([[2, [['C2', 3]]]]);
    expect(
      rowsAndCells(
        (
          await replay({
            count: 1,
            index: 2,
            kind: 'insert-columns',
            operationId: 'insert-columns',
            sheetKey,
          })
        ).document,
      ),
    ).toEqual([
      [
        1,
        [
          ['A1', 1],
          ['C1', 2],
        ],
      ],
      [3, [['D3', 3]]],
    ]);
    expect(
      rowsAndCells(
        (
          await replay({
            count: 1,
            index: 1,
            kind: 'delete-columns',
            operationId: 'delete-columns',
            sheetKey,
          })
        ).document,
      ),
    ).toEqual([
      [1, [['A1', 2]]],
      [3, [['B3', 3]]],
    ]);
    expect(
      rowsAndCells(
        (
          await replay({
            count: 1,
            index: 2,
            kind: 'delete-rows',
            operationId: 'delete-empty-row',
            sheetKey,
          })
        ).document,
      ),
    ).toEqual([
      [
        1,
        [
          ['A1', 1],
          ['B1', 2],
        ],
      ],
      [2, [['C2', 3]]],
    ]);
    expect(
      rowsAndCells(
        (
          await replay({
            count: 1,
            index: 2,
            kind: 'delete-columns',
            operationId: 'delete-middle-column',
            sheetKey,
          })
        ).document,
      ),
    ).toEqual([
      [1, [['A1', 1]]],
      [3, [['B3', 3]]],
    ]);
    await expect(
      replayXlsxCellOperations(
        snapshot.document,
        [
          {
            count: 2,
            index: 2,
            kind: 'insert-rows',
            operationId: 'row-reference-limit',
            sheetKey,
          },
        ],
        { ...writeLimits, maxReferenceUpdates: 2 },
        readerLimits,
      ),
    ).resolves.toBeDefined();
    expect(
      (
        await captureAsync(() =>
          replayXlsxCellOperations(
            snapshot.document,
            [
              {
                count: 1,
                index: 2,
                kind: 'insert-columns',
                operationId: 'column-reference-limit',
                sheetKey,
              },
            ],
            { ...writeLimits, maxReferenceUpdates: 1 },
            readerLimits,
          ),
        )
      ).diagnostic,
    ).toMatchObject({ actual: 2, limit: 1, limitName: 'maxReferenceUpdates' });

    const rowOverflow = structuredClone(snapshot.document);
    const row = worksheet(rowOverflow).rows[1]!;
    row.index = readerLimits.maxRowsPerWorksheet;
    row.cells[0]!.address = `C${readerLimits.maxRowsPerWorksheet}`;
    expect(
      (
        await captureAsync(() =>
          replayXlsxCellOperations(
            rowOverflow,
            [
              {
                count: 1,
                index: readerLimits.maxRowsPerWorksheet,
                kind: 'insert-rows',
                operationId: 'row-overflow',
                sheetKey,
              },
            ],
            writeLimits,
            readerLimits,
          ),
        )
      ).diagnostic,
    ).toMatchObject({
      featureClass: 'grid-overflow',
      message: 'XLSX row insertion would move an authored row outside the grid',
    });
    row.index = readerLimits.maxRowsPerWorksheet - 1;
    row.cells[0]!.address = `C${readerLimits.maxRowsPerWorksheet - 1}`;
    await expect(
      replayXlsxCellOperations(
        rowOverflow,
        [
          {
            count: 1,
            index: readerLimits.maxRowsPerWorksheet - 1,
            kind: 'insert-rows',
            operationId: 'row-boundary',
            sheetKey,
          },
        ],
        writeLimits,
        readerLimits,
      ),
    ).resolves.toBeDefined();
    row.index = readerLimits.maxRowsPerWorksheet;
    row.cells[0]!.address = `C${readerLimits.maxRowsPerWorksheet}`;
    await expect(
      replayXlsxCellOperations(
        rowOverflow,
        [
          {
            count: 1,
            index: readerLimits.maxRowsPerWorksheet,
            kind: 'delete-rows',
            operationId: 'delete-row-boundary',
            sheetKey,
          },
        ],
        writeLimits,
        readerLimits,
      ),
    ).resolves.toBeDefined();
    const columnOverflow = structuredClone(snapshot.document);
    const cell = worksheet(columnOverflow).rows[1]!.cells[0]!;
    cell.column = readerLimits.maxColumnsPerWorksheet;
    cell.address = `XFD3`;
    expect(
      (
        await captureAsync(() =>
          replayXlsxCellOperations(
            columnOverflow,
            [
              {
                count: 1,
                index: readerLimits.maxColumnsPerWorksheet,
                kind: 'insert-columns',
                operationId: 'column-overflow',
                sheetKey,
              },
            ],
            writeLimits,
            readerLimits,
          ),
        )
      ).diagnostic,
    ).toMatchObject({
      featureClass: 'grid-overflow',
      message:
        'XLSX column insertion would move an authored cell outside the grid',
    });
    cell.column = readerLimits.maxColumnsPerWorksheet - 1;
    cell.address = 'XFC3';
    await expect(
      replayXlsxCellOperations(
        columnOverflow,
        [
          {
            count: 1,
            index: readerLimits.maxColumnsPerWorksheet - 1,
            kind: 'insert-columns',
            operationId: 'column-boundary',
            sheetKey,
          },
        ],
        writeLimits,
        readerLimits,
      ),
    ).resolves.toBeDefined();
    cell.column = readerLimits.maxColumnsPerWorksheet;
    cell.address = 'XFD3';
    await expect(
      replayXlsxCellOperations(
        columnOverflow,
        [
          {
            count: 1,
            index: readerLimits.maxColumnsPerWorksheet,
            kind: 'delete-columns',
            operationId: 'delete-column-boundary',
            sheetKey,
          },
        ],
        writeLimits,
        readerLimits,
      ),
    ).resolves.toBeDefined();
    for (const [table, operation] of [
      [
        structuralTable(),
        {
          count: 1,
          index: 1,
          kind: 'insert-rows',
          operationId: 'table-shift',
          sheetKey,
        },
      ],
      [
        structuralTable({ totalsRow: true }),
        {
          count: 1,
          index: 2,
          kind: 'delete-rows',
          operationId: 'table-data-row',
          sheetKey,
        },
      ],
      [
        structuralTable({ headerRow: false }),
        {
          count: 1,
          index: 1,
          kind: 'delete-rows',
          operationId: 'table-no-header',
          sheetKey,
        },
      ],
      [
        structuralTable(),
        {
          count: 1,
          index: 3,
          kind: 'delete-rows',
          operationId: 'table-no-totals',
          sheetKey,
        },
      ],
      [
        structuralTable({
          autoFilter: {
            columns: [],
            range: {
              end: { column: 2, row: 5 },
              reference: 'A3:B5',
              start: { column: 1, row: 3 },
            },
            selectionRelation: 'full-sheet',
          },
          range: {
            end: { column: 2, row: 5 },
            reference: 'A3:B5',
            start: { column: 1, row: 3 },
          },
        }),
        {
          count: 1,
          index: 1,
          kind: 'delete-rows',
          operationId: 'table-header-after-delete',
          sheetKey,
        },
      ],
      [
        structuralTable({ totalsRow: true }),
        {
          count: 1,
          index: 5,
          kind: 'delete-rows',
          operationId: 'table-totals-before-delete',
          sheetKey,
        },
      ],
    ] as const) {
      const document = structuredClone(snapshot.document);
      worksheet(document).tables = [table];
      await expect(
        replayXlsxCellOperations(
          document,
          [operation],
          writeLimits,
          readerLimits,
        ),
      ).resolves.toBeDefined();
    }
  });

  it('blocks every reference-bearing structural closure domain', async () => {
    const snapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
      }),
    );
    const sheetKey = snapshot.document.sheets[0]!.key;
    const cases: Array<
      readonly [
        string,
        (document: XlsxRoundTripDocument) => void,
        ('delete-columns' | 'delete-rows' | 'insert-columns' | 'insert-rows')?,
      ]
    > = [
      [
        'defined-name-reference',
        (document) => void document.workbook.definedNames.push({} as never),
      ],
      [
        'calculation-chain-reference',
        (document) => void (document.workbook.calculation.chain = []),
      ],
      [
        'formula-reference',
        (document) => {
          worksheet(document).rows[0]!.cells[0]!.content = {
            cached: { kind: 'missing' },
            formula: { expression: '1+1', kind: 'normal' },
            kind: 'formula',
          };
        },
      ],
      [
        'conditional-format-formula-reference',
        (document) =>
          void setSheetField(document, 'conditionalFormattings', [
            { rules: [{ formulas: ['A1'] }, { formulas: [] }] },
          ]),
      ],
      [
        'conditional-format-formula-reference',
        (document) =>
          void setSheetField(document, 'conditionalFormattings', [
            {
              rules: [
                {
                  colorScale: {
                    stops: [
                      { threshold: { kind: 'formula' } },
                      { threshold: { kind: 'minimum' } },
                    ],
                  },
                  formulas: [],
                },
              ],
            },
          ]),
      ],
      [
        'conditional-format-formula-reference',
        (document) =>
          void setSheetField(document, 'conditionalFormattings', [
            {
              rules: [
                {
                  dataBar: {
                    thresholds: [{ kind: 'formula' }, { kind: 'minimum' }],
                  },
                  formulas: [],
                },
              ],
            },
          ]),
      ],
      [
        'conditional-format-formula-reference',
        (document) =>
          void setSheetField(document, 'conditionalFormattings', [
            {
              rules: [
                {
                  formulas: [],
                  iconSet: {
                    thresholds: [{ kind: 'formula' }, { kind: 'minimum' }],
                  },
                },
              ],
            },
          ]),
      ],
      [
        'data-validation-formula-reference',
        (document) =>
          void setSheetField(document, 'dataValidations', [{ formula1: 'A1' }]),
      ],
      [
        'data-validation-formula-reference',
        (document) =>
          void setSheetField(document, 'dataValidations', [{ formula2: 'A1' }]),
      ],
      [
        'drawing-chart-reference',
        (document) =>
          void setSheetField(document, 'drawings', [
            structuralDrawing({ object: { kind: 'chart' } as never }),
          ]),
      ],
      [
        'drawing-chart-reference',
        (document) =>
          void setSheetField(document, 'drawings', [
            structuralDrawing({
              object: {
                children: [{ kind: 'chart' }, structuralDrawing().object],
                kind: 'group',
              } as never,
            }),
          ]),
      ],
      [
        'protection-reference',
        (document) => void setSheetField(document, 'protection', {}),
      ],
      [
        'pivot-reference',
        (document) => void setSheetField(document, 'pivotTables', []),
      ],
      [
        'query-table-reference',
        (document) => void setSheetField(document, 'queryTables', []),
      ],
      [
        'slicer-reference',
        (document) => void setSheetField(document, 'slicers', []),
      ],
      [
        'sparkline-formula-reference',
        (document) =>
          void setSheetField(document, 'sparklineGroups', [
            structuralSparklineGroup({
              sparklines: [
                {
                  dataFormula: 'SUM(A1:A3)',
                  location: 'B1',
                  selectionRelation: 'full-sheet',
                },
              ],
            }),
          ]),
      ],
      [
        'table-formula-reference',
        (document) =>
          void setSheetField(document, 'tables', [
            structuralTable({
              columns: [
                {
                  calculatedFormula: { array: false, expression: 'A1' },
                  id: 1,
                  name: 'A',
                  totalsFunction: 'none',
                },
                { id: 2, name: 'B', totalsFunction: 'none' },
              ],
            }),
          ]),
      ],
      [
        'table-formula-reference',
        (document) =>
          void setSheetField(document, 'tables', [
            structuralTable({
              columns: [
                {
                  id: 1,
                  name: 'A',
                  totalsFormula: { array: false, expression: 'A1' },
                  totalsFunction: 'custom',
                },
                { id: 2, name: 'B', totalsFunction: 'none' },
              ],
            }),
          ]),
      ],
      [
        'timeline-reference',
        (document) => void setSheetField(document, 'timelines', []),
      ],
      [
        'view-pane-reference',
        (document) => void setSheetField(document, 'views', [{ pane: {} }]),
      ],
      [
        'cell-metadata-reference',
        (document) => {
          worksheet(document).rows[0]!.cells[0]!.metadata = {};
        },
      ],
      [
        'column-definition',
        (document) =>
          void worksheet(document).columns.push({ end: 1, start: 1 }),
        'insert-columns',
      ],
      [
        'column-definition',
        (document) =>
          void worksheet(document).columns.push({ end: 1, start: 1 }),
        'delete-columns',
      ],
    ];
    function setSheetField(
      document: XlsxRoundTripDocument,
      field: string,
      value: unknown,
    ): void {
      (worksheet(document) as unknown as Record<string, unknown>)[field] =
        value;
    }
    for (const [featureClass, mutate, kind = 'insert-rows'] of cases) {
      const document = structuredClone(snapshot.document);
      mutate(document);
      const error = await captureAsync(() =>
        replayXlsxCellOperations(
          document,
          [
            {
              count: 1,
              index: 1,
              kind,
              operationId: `block-${featureClass}`,
              sheetKey,
            },
          ],
          writeLimits,
          readerLimits,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        code: 'unsupported-edit-operation',
        featureClass,
        message:
          'XLSX structural edit requires a reference-free worksheet closure',
      });
    }
    const commentDocument = structuredClone(snapshot.document);
    worksheet(commentDocument).comments = [
      {
        author: 'Author',
        kind: 'note',
        reference: 'A1',
        selectionRelation: 'full-sheet',
        text: 'Note',
        visible: false,
      },
    ];
    expect(
      (
        await captureAsync(() =>
          replayXlsxCellOperations(
            commentDocument,
            [
              {
                count: 1,
                index: 1,
                kind: 'delete-rows',
                operationId: 'delete-comment-anchor',
                sheetKey,
              },
            ],
            writeLimits,
            readerLimits,
          ),
        )
      ).diagnostic,
    ).toMatchObject({
      code: 'unsupported-edit-operation',
      featureClass: 'comment-anchor-deletion',
    });
    const drawingDocument = structuredClone(snapshot.document);
    worksheet(drawingDocument).drawings = [
      structuralDrawing({
        from: { column: 1, columnOffset: 0, row: 1, rowOffset: 0 },
        kind: 'two-cell',
        to: { column: 1, columnOffset: 0, row: 1, rowOffset: 0 },
      }),
    ];
    expect(
      (
        await captureAsync(() =>
          replayXlsxCellOperations(
            drawingDocument,
            [
              {
                count: 1,
                index: 1,
                kind: 'delete-rows',
                operationId: 'delete-drawing-anchor',
                sheetKey,
              },
            ],
            writeLimits,
            readerLimits,
          ),
        )
      ).diagnostic,
    ).toMatchObject({
      code: 'unsupported-edit-operation',
      featureClass: 'drawing-anchor-deletion',
    });
    for (const [featureClass, dataFormula, location, count] of [
      ['sparkline-source-deletion', 'A1:A2', 'B3', 2],
      ['sparkline-location-deletion', 'A1:A3', 'B1', 1],
    ] as const) {
      const sparklineDocument = structuredClone(snapshot.document);
      worksheet(sparklineDocument).sparklineGroups = [
        structuralSparklineGroup({
          sparklines: [
            {
              dataFormula,
              location,
              selectionRelation: 'full-sheet',
            },
          ],
        }),
      ];
      expect(
        (
          await captureAsync(() =>
            replayXlsxCellOperations(
              sparklineDocument,
              [
                {
                  count,
                  index: 1,
                  kind: 'delete-rows',
                  operationId: `delete-${featureClass}`,
                  sheetKey,
                },
              ],
              writeLimits,
              readerLimits,
            ),
          )
        ).diagnostic,
      ).toMatchObject({
        code: 'unsupported-edit-operation',
        featureClass,
      });
    }
    const drawingBudgetDocument = structuredClone(snapshot.document);
    const drawingBudgetSheet = worksheet(drawingBudgetDocument);
    drawingBudgetSheet.rows = [];
    delete drawingBudgetSheet.declaredDimension;
    const absoluteDrawing = structuralDrawing({ kind: 'absolute' });
    delete absoluteDrawing.from;
    drawingBudgetSheet.drawings = [absoluteDrawing, structuralDrawing()];
    await expect(
      replayXlsxCellOperations(
        drawingBudgetDocument,
        [
          {
            count: 1,
            index: 1,
            kind: 'insert-rows',
            operationId: 'drawing-reference-budget',
            sheetKey,
          },
        ],
        { ...writeLimits, maxReferenceUpdates: 1 },
        readerLimits,
      ),
    ).resolves.toBeDefined();
    expect(
      (
        await captureAsync(() =>
          replayXlsxCellOperations(
            drawingBudgetDocument,
            [
              {
                count: 1,
                index: 1,
                kind: 'insert-rows',
                operationId: 'drawing-reference-budget',
                sheetKey,
              },
            ],
            { ...writeLimits, maxReferenceUpdates: 0 },
            readerLimits,
          ),
        )
      ).diagnostic,
    ).toMatchObject({ actual: 1, limit: 0, limitName: 'maxReferenceUpdates' });
    const rowWithColumns = structuredClone(snapshot.document);
    worksheet(rowWithColumns).columns.push({ end: 1, start: 1 });
    await expect(
      replayXlsxCellOperations(
        rowWithColumns,
        [
          {
            count: 1,
            index: 1,
            kind: 'insert-rows',
            operationId: 'row-with-columns',
            sheetKey,
          },
        ],
        writeLimits,
        readerLimits,
      ),
    ).resolves.toMatchObject({
      impacts: [expect.objectContaining({ kind: 'insert-rows' })],
    });
  });

  it('blocks table structural edits that cannot preserve table shape', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const sheetKey = snapshot.document.sheets[0]!.key;
    const cases: Array<
      readonly [
        string,
        XlsxTable,
        Extract<
          XlsxEditOperation,
          {
            kind:
              | 'delete-columns'
              | 'delete-rows'
              | 'insert-columns'
              | 'insert-rows';
          }
        >,
      ]
    > = [
      [
        'table-column-structure',
        structuralTable(),
        {
          count: 1,
          index: 2,
          kind: 'insert-columns',
          operationId: 'table-column',
          sheetKey,
        },
      ],
      [
        'table-range-deletion',
        structuralTable(),
        {
          count: 3,
          index: 1,
          kind: 'delete-rows',
          operationId: 'table-delete',
          sheetKey,
        },
      ],
      [
        'table-row-structure',
        structuralTable({
          range: {
            end: { column: 2, row: 2 },
            reference: 'A1:B2',
            start: { column: 1, row: 1 },
          },
          totalsRow: true,
        }),
        {
          count: 1,
          index: 1,
          kind: 'delete-rows',
          operationId: 'table-short',
          sheetKey,
        },
      ],
      [
        'table-header-row',
        structuralTable(),
        {
          count: 1,
          index: 1,
          kind: 'delete-rows',
          operationId: 'table-header',
          sheetKey,
        },
      ],
      [
        'table-totals-row',
        structuralTable({ totalsRow: true }),
        {
          count: 1,
          index: 3,
          kind: 'delete-rows',
          operationId: 'table-totals',
          sheetKey,
        },
      ],
    ];
    for (const [featureClass, table, operation] of cases) {
      const document = structuredClone(snapshot.document);
      worksheet(document).tables = [table];
      expect(
        (
          await captureAsync(() =>
            replayXlsxCellOperations(
              document,
              [operation],
              writeLimits,
              readerLimits,
            ),
          )
        ).diagnostic,
      ).toMatchObject({
        code: 'unsupported-edit-operation',
        featureClass,
      });
    }
    const outside = structuredClone(snapshot.document);
    worksheet(outside).tables = [structuralTable()];
    await expect(
      replayXlsxCellOperations(
        outside,
        [
          {
            count: 1,
            index: 3,
            kind: 'insert-columns',
            operationId: 'table-outside',
            sheetKey,
          },
        ],
        writeLimits,
        readerLimits,
      ),
    ).resolves.toBeDefined();
  });

  it('applies existing styles, appends deterministically, and bounds style growth', async () => {
    const styles = `<styleSheet xmlns="${XLSX_SPREADSHEET_NS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    const snapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({ 'xl/styles.xml': styles }),
    );
    const sheet = worksheet(snapshot.document);
    const targetStyle = snapshot.document.styles[1]!;
    const operation = {
      cell: 'A1',
      kind: 'set-cell-style' as const,
      operationId: 'style-1',
      sheetKey: sheet.key,
      style: targetStyle,
    };
    const plan = await replayXlsxCellOperations(
      snapshot.document,
      [operation],
      writeLimits,
      readerLimits,
    );
    expect(sheet.rows[0]!.cells[0]!.style).toBeUndefined();
    expect(worksheet(plan.document).rows[0]!.cells[0]!.style).toBe(1);
    expect(plan.impacts).toEqual([
      {
        cell: 'A1',
        kind: 'set-cell-style',
        operationId: 'style-1',
        sheetKey: sheet.key,
      },
    ]);
    const defaultPlan = await replayXlsxCellOperations(
      snapshot.document,
      [
        {
          ...operation,
          operationId: 'style-0',
          style: snapshot.document.styles[0]!,
        },
      ],
      writeLimits,
      readerLimits,
    );
    expect(worksheet(defaultPlan.document).rows[0]!.cells[0]!.style).toBe(0);
    const appended = await replayXlsxCellOperations(
      snapshot.document,
      [{ ...operation, style: { font: { italic: true } } }],
      writeLimits,
      readerLimits,
    );
    expect(appended.document.styles).toEqual([
      ...snapshot.document.styles,
      { font: { italic: true } },
    ]);
    expect(worksheet(appended.document).rows[0]!.cells[0]!.style).toBe(2);
    const error = await captureAsync(() =>
      replayXlsxCellOperations(
        snapshot.document,
        [{ ...operation, style: { checkbox: true } }],
        writeLimits,
        readerLimits,
      ),
    );
    expect(error.diagnostic).toMatchObject({
      cell: 'A1',
      code: 'unsupported-edit-operation',
      featureClass: 'append-checkbox-style',
      message:
        'XLSX cannot append a checkbox style without a feature-property-bag edit',
      operationId: 'style-1',
      sheetKey: sheet.key,
    });
    await expect(
      replayXlsxCellOperations(
        snapshot.document,
        [{ ...operation, style: { font: { italic: true } } }],
        writeLimits,
        {
          ...readerLimits,
          maxStyles: snapshot.document.styles.length + 1,
        },
      ),
    ).resolves.toMatchObject({
      document: {
        styles: [...snapshot.document.styles, { font: { italic: true } }],
      },
    });
    const limitError = await captureAsync(() =>
      replayXlsxCellOperations(
        snapshot.document,
        [{ ...operation, style: { font: { italic: true } } }],
        writeLimits,
        { ...readerLimits, maxStyles: snapshot.document.styles.length },
      ),
    );
    expect(limitError.diagnostic).toMatchObject({
      actual: snapshot.document.styles.length + 1,
      code: 'resource-limit-exceeded',
      limit: snapshot.document.styles.length,
      limitName: 'maxStyles',
      message: 'XLSX edited normalized styles exceed their reader limit',
      operationId: 'style-1',
    });
  });

  it('replays internal hyperlink targets with sequential preconditions and range guards', async () => {
    const snapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>2</v></c><c r="D2"><v>2</v></c></row><row r="3"><c r="C3"><v>3</v></c></row><row r="4"><c r="B4"><v>4</v></c></row></sheetData><hyperlinks><hyperlink ref="A1" location="Old!A1" display="Old" tooltip="Tip"/><hyperlink ref="B2:C3" location="Range!A1"/></hyperlinks></worksheet>`,
      }),
    );
    const sourceSheet = worksheet(snapshot.document);
    const sourceCell = sourceSheet.rows[0]!.cells[0]!;
    const sourceMatch = await canonicalXlsxSha256(
      xlsxCellTargetState(sourceSheet, sourceCell),
    );
    expect(xlsxCellTargetState(sourceSheet, sourceCell)).toMatchObject({
      hyperlink: {
        display: 'Old',
        target: { kind: 'internal', location: 'Old!A1' },
      },
    });
    const noExactLinkCell = sourceSheet.rows
      .flatMap((row) => row.cells)
      .find((cell) => cell.address === 'B2')!;
    expect(
      xlsxCellTargetState(sourceSheet, noExactLinkCell),
    ).not.toHaveProperty('hyperlink');
    const update = {
      cell: 'A1',
      ifMatch: sourceMatch,
      kind: 'set-hyperlink' as const,
      operationId: 'update-link',
      sheetKey: sourceSheet.key,
      target: { kind: 'internal' as const, location: 'New!B2' },
    };
    const updated = await replayXlsxCellOperations(
      snapshot.document,
      [update],
      writeLimits,
      readerLimits,
    );
    expect(worksheet(updated.document).hyperlinks[0]).toMatchObject({
      display: 'Old',
      target: { kind: 'internal', location: 'New!B2' },
      tooltip: 'Tip',
    });
    const updatedSheet = worksheet(updated.document);
    const updatedMatch = await canonicalXlsxSha256(
      xlsxCellTargetState(updatedSheet, updatedSheet.rows[0]!.cells[0]!),
    );
    const removed = await replayXlsxCellOperations(
      snapshot.document,
      [
        update,
        {
          cell: 'A1',
          ifMatch: updatedMatch,
          kind: 'set-hyperlink',
          operationId: 'remove-link',
          sheetKey: sourceSheet.key,
          target: null,
        },
      ],
      writeLimits,
      readerLimits,
    );
    expect(worksheet(removed.document).hyperlinks).toHaveLength(1);
    expect(worksheet(removed.document).hyperlinks[0]!.range.reference).toBe(
      'B2:C3',
    );
    const outside = await replayXlsxCellOperations(
      snapshot.document,
      ['B1', 'A2', 'D2', 'B4'].map((cell, index) => ({
        cell,
        kind: 'set-hyperlink' as const,
        operationId: `outside-${index}`,
        sheetKey: sourceSheet.key,
        target: null,
      })),
      writeLimits,
      readerLimits,
    );
    expect(worksheet(outside.document).hyperlinks).toHaveLength(2);
    for (const cell of ['B2', 'C3']) {
      const conflict = await captureAsync(() =>
        replayXlsxCellOperations(
          snapshot.document,
          [
            {
              cell,
              kind: 'set-hyperlink',
              operationId: `overlap-${cell}`,
              sheetKey: sourceSheet.key,
              target: { kind: 'internal', location: 'Other!A1' },
            },
          ],
          writeLimits,
          readerLimits,
        ),
      );
      expect(conflict.diagnostic).toMatchObject({
        cell,
        code: 'preservation-conflict',
        featureClass: 'hyperlink-range',
        message:
          'XLSX hyperlink operation overlaps a multi-cell hyperlink range',
        operationId: `overlap-${cell}`,
      });
    }
  });

  it('checks ifMatch against the sequential target state', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const sourceSheet = worksheet(snapshot.document);
    const sourceCell = sourceSheet.rows[0]!.cells[0]!;
    const sourceMatch = await canonicalXlsxSha256(
      xlsxCellTargetState(sourceSheet, sourceCell),
    );
    const first = cellOperation(snapshot.document, { ifMatch: sourceMatch });
    const firstPlan = await replayXlsxCellOperations(
      snapshot.document,
      [first],
      writeLimits,
      readerLimits,
    );
    const updatedSheet = worksheet(firstPlan.document);
    const updatedMatch = await canonicalXlsxSha256(
      xlsxCellTargetState(updatedSheet, updatedSheet.rows[0]!.cells[0]!),
    );
    const plan = await replayXlsxCellOperations(
      snapshot.document,
      [
        first,
        cellOperation(snapshot.document, {
          content: { kind: 'value', value: { kind: 'boolean', value: true } },
          ifMatch: updatedMatch,
          operationId: 'edit-2',
        }),
      ],
      writeLimits,
      readerLimits,
    );
    expect(worksheet(plan.document).rows[0]?.cells[0]?.content).toEqual({
      kind: 'value',
      value: { kind: 'boolean', value: true },
    });

    const error = await captureAsync(() =>
      replayXlsxCellOperations(
        snapshot.document,
        [first, { ...first, operationId: 'edit-2' }],
        writeLimits,
        readerLimits,
      ),
    );
    expect(error.diagnostic).toMatchObject({
      cell: 'A1',
      code: 'operation-precondition-failed',
      operationId: 'edit-2',
      sheetKey: sourceSheet.key,
    });
    expect(error.diagnostic.message).toBe(
      'XLSX operation precondition does not match the target cell',
    );
  });

  it('blocks missing targets and resolves row and column coordinates exactly', async () => {
    const snapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}">
          <sheetData>
            <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
            <row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>4</v></c></row>
          </sheetData>
        </worksheet>`,
      }),
    );
    const cases = [
      cellOperation(snapshot.document, {
        sheetKey: `xlsx:sheet:${'f'.repeat(32)}`,
      }),
      cellOperation(snapshot.document, { cell: 'A2' }),
      cellOperation(snapshot.document, { cell: 'C1' }),
    ];
    const features = ['worksheet', 'missing-cell', 'missing-cell'];
    const messages = [
      'XLSX operation sheet key does not exist in the snapshot',
      'XLSX cell operation requires an existing explicit source cell',
      'XLSX cell operation requires an existing explicit source cell',
    ];
    for (const [index, operation] of cases.entries()) {
      const error = await captureAsync(() =>
        replayXlsxCellOperations(
          snapshot.document,
          [operation],
          writeLimits,
          readerLimits,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        code: 'preservation-conflict',
        featureClass: features[index],
        operationId: 'edit-1',
      });
      expect(error.diagnostic.message).toBe(messages[index]);
    }
  });

  it('applies merged-range geometry only inside the range and only at its anchor', async () => {
    const snapshot = await readXlsxRoundTrip(
      await createIndependentXlsx({
        'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}">
          <sheetData>
            <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>1</v></c><c r="D1"><v>1</v></c></row>
            <row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>1</v></c><c r="C2"><v>1</v></c><c r="D2"><v>1</v></c></row>
            <row r="3"><c r="A3"><v>1</v></c><c r="B3"><v>1</v></c><c r="C3"><v>1</v></c><c r="D3"><v>1</v></c></row>
            <row r="4"><c r="A4"><v>1</v></c><c r="B4"><v>1</v></c><c r="D4"><v>1</v></c></row>
          </sheetData>
          <mergeCells count="1"><mergeCell ref="B2:C3"/></mergeCells>
        </worksheet>`,
      }),
    );
    for (const cell of ['A2', 'D2', 'B1', 'B4', 'B2']) {
      await expect(
        replayXlsxCellOperations(
          snapshot.document,
          [cellOperation(snapshot.document, { cell })],
          writeLimits,
          readerLimits,
        ),
      ).resolves.toMatchObject({
        impacts: [{ cell, operationId: 'edit-1' }],
      });
    }
    for (const cell of ['C2', 'B3']) {
      const error = await captureAsync(() =>
        replayXlsxCellOperations(
          snapshot.document,
          [cellOperation(snapshot.document, { cell })],
          writeLimits,
          readerLimits,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        cell,
        code: 'preservation-conflict',
        featureClass: 'merged-cell',
        message: 'XLSX cell operation cannot target a non-anchor merged cell',
      });
    }
  });

  it('rejects chart-sheet targets and duplicate snapshot keys', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const chartDocument: XlsxRoundTripDocument = {
      ...snapshot.document,
      sheets: [
        {
          index: 0,
          key: worksheet(snapshot.document).key,
          kind: 'chart-sheet',
          name: 'Chart',
          payload: 'full-sheet',
          state: 'visible',
        },
      ],
    };
    const chartError = await captureAsync(() =>
      replayXlsxCellOperations(
        chartDocument,
        [cellOperation(snapshot.document)],
        writeLimits,
        readerLimits,
      ),
    );
    expect(chartError.diagnostic.featureClass).toBe('chart-sheet');
    expect(chartError.diagnostic.message).toBe(
      'XLSX cell operation cannot target a chart sheet',
    );
    const chartRowError = await captureAsync(() =>
      replayXlsxCellOperations(
        chartDocument,
        [
          {
            hidden: true,
            kind: 'set-row',
            operationId: 'chart-row',
            row: 1,
            sheetKey: chartDocument.sheets[0]!.key,
          },
        ],
        writeLimits,
        readerLimits,
      ),
    );
    expect(chartRowError.diagnostic).toMatchObject({
      featureClass: 'chart-sheet',
      message: 'XLSX row or column operation cannot target a chart sheet',
      range: '1',
    });

    const duplicate: XlsxRoundTripDocument = {
      ...snapshot.document,
      sheets: [
        snapshot.document.sheets[0]!,
        { ...snapshot.document.sheets[0]!, index: 1 },
      ],
    };
    const duplicateError = await captureAsync(() =>
      replayXlsxCellOperations(duplicate, [], writeLimits, readerLimits),
    );
    expect(duplicateError.diagnostic).toMatchObject({
      code: 'snapshot-integrity-failed',
      message: 'XLSX snapshot sheet keys must be unique',
      objectKey: worksheet(snapshot.document).key,
    });
  });

  it('returns an isolated literal-equal state for an empty operation list', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    const plan = await replayXlsxCellOperations(
      snapshot.document,
      [],
      resolveXlsxWriteLimits(undefined),
      readerLimits,
    );
    expect(plan.document).toEqual(snapshot.document);
    expect(plan.document).not.toBe(snapshot.document);
    expect(plan.impacts).toEqual([]);
    expect(plan.stateHash).toBe(snapshot.stateHash);
  });
});
