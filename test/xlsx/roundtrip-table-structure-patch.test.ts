import { describe, expect, it } from 'vitest';

import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import { patchXlsxTableStructure } from '../../src/formats/xlsx/roundtrip/table-structure-patch';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import { XLSX_SPREADSHEET_NS } from '../black-box/xlsx-package';

const PART = 'xl/tables/table1.xml';

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
  throw new Error('Expected table structural patch to fail');
}

function table(inner = ''): string {
  return `<table xmlns="${XLSX_SPREADSHEET_NS}" id="1" name="Table1" displayName="Table1" ref="A1:C5">${inner}<tableColumns count="3"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/><tableColumn id="3" name="C"/></tableColumns></table>`;
}

describe('XLSX table structural patching', () => {
  it('transforms table filter, sort, and condition ranges', () => {
    const source = table(
      '<autoFilter ref="A1:C5"><sortState ref="A1:C5"><sortCondition ref="A2:A3"/><sortCondition ref="B1:B5"/></sortState></autoFilter>',
    );
    const result = patchXlsxTableStructure(
      bytes(source),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'table-rows',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<table xmlns="${XLSX_SPREADSHEET_NS}" id="1" name="Table1" displayName="Table1" ref="A1:C3"><autoFilter ref="A1:C3"><sortState ref="A1:C3"><sortCondition ref="B1:B3"/></sortState></autoFilter><tableColumns count="3"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/><tableColumn id="3" name="C"/></tableColumns></table>`,
    );
    expect(result.patchCount).toBe(5);
    const attributeBytes = new TextEncoder().encode(' ref="A1:C3"').byteLength;
    const conditionBytes = new TextEncoder().encode(' ref="B1:B3"').byteLength;
    expect(result.patchBytes).toBe(attributeBytes * 3 + conditionBytes);
  });

  it('removes deleted table filter and sort subtrees', () => {
    const filter = patchXlsxTableStructure(
      bytes(
        table('<autoFilter ref="A2:C3"><sortState ref="A2:C3"/></autoFilter>'),
      ),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'remove-filter',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(filter.data)).toBe(
      `<table xmlns="${XLSX_SPREADSHEET_NS}" id="1" name="Table1" displayName="Table1" ref="A1:C3"><tableColumns count="3"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/><tableColumn id="3" name="C"/></tableColumns></table>`,
    );
    const sort = patchXlsxTableStructure(
      bytes(
        table(
          '<autoFilter ref="A1:C5"><sortState ref="A2:C3"><sortCondition ref="A2:A3"/></sortState></autoFilter>',
        ),
      ),
      [
        {
          count: 2,
          index: 2,
          kind: 'delete-rows',
          operationId: 'remove-sort',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(sort.data)).toContain(
      '<autoFilter ref="A1:C3"></autoFilter>',
    );
  });

  it('preserves exact bytes for no-op transforms and empty requests', () => {
    const source = table('<autoFilter ref="A1:C5"/>');
    const noOp = patchXlsxTableStructure(
      bytes(source),
      [
        {
          count: 1,
          index: 8,
          kind: 'insert-rows',
          operationId: 'table-no-op',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(noOp.data)).toBe(source);
    expect(noOp.patchCount).toBe(0);
    const input = bytes(source);
    const empty = patchXlsxTableStructure(
      input,
      [],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(empty.data).toEqual(input);
    expect(empty.data).not.toBe(input);
  });

  it('selects only owned prefixed table nodes', () => {
    const source = `<s:table xmlns:s="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:foreign" ref="A1:C3"><wrapper><s:autoFilter ref="Z9"><s:sortState ref="Z9"><s:sortCondition ref="Z9"/></s:sortState></s:autoFilter><s:sortState ref="Z6"/></wrapper><x:autoFilter ref="Z9"/><s:autoFilter ref="A1:C3"><wrapper><s:sortState ref="Z8"/></wrapper><x:sortState ref="Z9"/><wrapper><s:sortCondition ref="Z6"/></wrapper><s:sortState ref="A1:C3"><wrapper><s:sortCondition ref="Z7"/></wrapper><x:sortCondition ref="Z9"/><s:sortCondition ref="A1:A3"/></s:sortState><wrapper><s:sortCondition ref="Z5"/></wrapper></s:autoFilter><wrapper><s:sortState ref="Z4"/></wrapper></s:table>`;
    const result = patchXlsxTableStructure(
      bytes(source),
      [
        {
          count: 1,
          index: 1,
          kind: 'insert-rows',
          operationId: 'owned-table',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const output = new TextDecoder().decode(result.data);
    expect(output).toContain('<s:table xmlns:s=');
    expect(output).toContain('ref="A2:C4">');
    expect(output).toContain('<s:autoFilter ref="A2:C4">');
    expect(output).toContain('<s:sortState ref="A2:C4">');
    expect(output).toContain('<s:sortCondition ref="A2:A4"/>');
    expect(output).toContain('<wrapper><s:autoFilter ref="Z9">');
    expect(output).toContain('<s:sortState ref="Z6"/>');
    expect(output).toContain('<wrapper><s:sortState ref="Z8"/></wrapper>');
    expect(output).toContain('<wrapper><s:sortCondition ref="Z7"/></wrapper>');
    expect(output).toContain('<wrapper><s:sortCondition ref="Z5"/></wrapper>');
    expect(output).toContain('<wrapper><s:sortState ref="Z4"/></wrapper>');
    expect(output).toContain('<x:autoFilter ref="Z9"/>');
  });

  it('rejects malformed and deleted table structures exactly', () => {
    for (const [source, message] of [
      ['<table/>', 'XLSX structural table range is invalid'],
      ['<table ref="bad"/>', 'XLSX structural table range is invalid'],
      ['<wrong ref="A1"/>', 'XLSX table root cannot patch structure'],
      [
        '<wrapper><table ref="A1"/></wrapper>',
        'XLSX table root cannot patch structure',
      ],
      [
        '<table ref="A1"><autoFilter ref="bad"/></table>',
        'XLSX structural table auto-filter range is invalid',
      ],
      [
        '<table ref="A1"><autoFilter ref="A1"><sortState ref="bad"/></autoFilter></table>',
        'XLSX structural table sort range is invalid',
      ],
      [
        '<table ref="A1"><autoFilter ref="A1"><sortState ref="A1"><sortCondition ref="bad"/></sortState></autoFilter></table>',
        'XLSX structural table sort-condition range is invalid',
      ],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxTableStructure(
            bytes(source),
            [
              {
                count: 1,
                index: 1,
                kind: 'insert-rows',
                operationId: 'bad-table',
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic.message,
      ).toBe(message);
    }
    expect(
      capture(() =>
        patchXlsxTableStructure(
          bytes('<table ref="A2:C3"/>'),
          [
            {
              count: 2,
              index: 2,
              kind: 'delete-rows',
              operationId: 'delete-table',
            },
          ],
          defaultXlsxWriteLimits(),
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({
      featureClass: 'table-structure-xml',
      message: 'XLSX structural table range cannot be deleted',
      operationId: 'delete-table',
      part: PART,
      range: '2:3',
    });
  });

  it('enforces table patch resource limits at exact boundaries', () => {
    const source = table('<autoFilter ref="A1:C5"/>');
    const request = {
      count: 1,
      index: 2,
      kind: 'insert-rows' as const,
      operationId: 'table-limits',
    };
    const baseline = patchXlsxTableStructure(
      bytes(source),
      [request],
      defaultXlsxWriteLimits(),
      PART,
    );
    const replacement = new TextEncoder().encode(' ref="A1:C6"').byteLength;
    expect(baseline.patchBytes).toBe(replacement * 2);
    expect(baseline.patchCount).toBe(2);
    expect(() =>
      patchXlsxTableStructure(
        bytes(source),
        [request],
        {
          ...defaultXlsxWriteLimits(),
          maxGeneratedXmlBytes: baseline.data.byteLength,
          maxPatchBytes: baseline.patchBytes,
          maxPatchCount: baseline.patchCount,
        },
        PART,
      ),
    ).not.toThrow();
    for (const [limitName, limit] of [
      ['maxGeneratedXmlBytes', baseline.data.byteLength - 1],
      ['maxPatchBytes', baseline.patchBytes - 1],
      ['maxPatchCount', baseline.patchCount - 1],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxTableStructure(
            bytes(source),
            [request],
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ limitName });
    }
  });
});
