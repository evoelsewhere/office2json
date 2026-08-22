import { describe, expect, it } from 'vitest';

import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import { patchXlsxWorksheetProperties } from '../../src/formats/xlsx/roundtrip/worksheet-properties-patch';
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
  throw new Error('Expected worksheet property patch to fail');
}

function source(): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><cols><col min="1" max="2" width="10" customWidth="1" hidden="0" style="3" extra="kept"/><col min="3" max="3"/></cols><sheetData><row r="1" ht="12" customHeight="1" hidden="0" spans="1:2"><c r="A1"><v>1</v></c></row><row r="2"/></sheetData></worksheet>`;
}

describe('XLSX worksheet row and column property patching', () => {
  it('patches existing rows and exact column ranges while preserving unrelated XML', () => {
    const result = patchXlsxWorksheetProperties(
      bytes(source()),
      [
        {
          end: 2,
          hidden: true,
          kind: 'set-column',
          operationId: 'column-1',
          start: 1,
          width: 25,
        },
        {
          height: 20,
          hidden: true,
          kind: 'set-row',
          operationId: 'row-1',
          row: 1,
        },
        {
          end: 3,
          hidden: false,
          kind: 'set-column',
          operationId: 'column-3',
          start: 3,
        },
        {
          hidden: false,
          kind: 'set-row',
          operationId: 'row-2',
          row: 2,
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><cols><col min="1" max="2" style="3" extra="kept" width="25" customWidth="1" hidden="1"/><col min="3" max="3" hidden="0"/></cols><sheetData><row r="1" spans="1:2" ht="20" customHeight="1" hidden="1"><c r="A1"><v>1</v></c></row><row r="2" hidden="0"/></sheetData></worksheet>`,
    );
    expect(result.patchCount).toBe(4);
    expect(result.patchBytes).toBe(198);
  });

  it('returns an owned byte copy for an empty patch list', () => {
    const input = bytes(source());
    const result = patchXlsxWorksheetProperties(
      input,
      [],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(result).toEqual({ data: input, patchBytes: 0, patchCount: 0 });
    expect(result.data).not.toBe(input);
    const malformed = bytes('<not-xml');
    expect(
      patchXlsxWorksheetProperties(
        malformed,
        [],
        defaultXlsxWriteLimits(),
        PART,
      ).data,
    ).toEqual(malformed);
  });

  it('preserves a property when only the other property is requested', () => {
    const result = patchXlsxWorksheetProperties(
      bytes(source()),
      [
        { height: 30, kind: 'set-row', operationId: 'row', row: 1 },
        {
          end: 2,
          kind: 'set-column',
          operationId: 'column',
          start: 1,
          width: 30,
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const output = new TextDecoder().decode(result.data);
    expect(output).toContain(
      '<col min="1" max="2" hidden="0" style="3" extra="kept" width="30" customWidth="1"/>',
    );
    expect(output).toContain(
      '<row r="1" hidden="0" spans="1:2" ht="30" customHeight="1">',
    );
  });

  it('patches prefixed UTF-16LE XML', () => {
    const xml = source()
      .replaceAll(
        /<(\/)?(worksheet|cols|col|sheetData|row|c|v)(?=[\s/>])/gu,
        '<$1s:$2',
      )
      .replace('xmlns=', 'xmlns:s=');
    const encoded = new Uint8Array(2 + xml.length * 2);
    encoded.set([0xff, 0xfe]);
    for (let index = 0; index < xml.length; index += 1) {
      const code = xml.charCodeAt(index);
      encoded[2 + index * 2] = code & 0xff;
      encoded[3 + index * 2] = code >>> 8;
    }
    const result = patchXlsxWorksheetProperties(
      encoded,
      [
        { height: 0, kind: 'set-row', operationId: 'row', row: 1 },
        {
          end: 2,
          kind: 'set-column',
          operationId: 'column',
          start: 1,
          width: 0,
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect([...result.data.slice(0, 2)]).toEqual([0xff, 0xfe]);
    const output = new TextDecoder('utf-16le').decode(result.data.slice(2));
    expect(output).toContain('<s:col min="1" max="2"');
    expect(output).toContain('width="0" customWidth="1"');
    expect(output).toContain('<s:row r="1"');
    expect(output).toContain('ht="0" customHeight="1"');
  });

  it('requires unique direct row and column spans with structured provenance', () => {
    const cases = [
      {
        expected: {
          featureClass: 'worksheet-property-xml',
          message: 'XLSX worksheet root cannot patch properties',
        },
        request: {
          hidden: true,
          kind: 'set-row' as const,
          operationId: 'root',
          row: 1,
        },
        xml: `<outer>${source()}</outer>`,
      },
      {
        expected: {
          featureClass: 'missing-row-span',
          message: 'XLSX target row has no safe explicit XML span',
          operationId: 'row',
          range: '1',
        },
        request: {
          hidden: true,
          kind: 'set-row' as const,
          operationId: 'row',
          row: 1,
        },
        xml: `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><cols/></worksheet>`,
      },
      {
        expected: {
          featureClass: 'missing-column-span',
          message: 'XLSX target column range has no safe explicit XML span',
          operationId: 'column',
          range: '1:2',
        },
        request: {
          end: 2,
          hidden: true,
          kind: 'set-column' as const,
          operationId: 'column',
          start: 1,
        },
        xml: `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/></worksheet>`,
      },
    ];
    for (const item of cases) {
      expect(
        capture(() =>
          patchXlsxWorksheetProperties(
            bytes(item.xml),
            [item.request],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ ...item.expected, part: PART });
    }

    expect(
      capture(() =>
        patchXlsxWorksheetProperties(
          bytes(source()),
          [
            { hidden: true, kind: 'set-row', operationId: 'one', row: 1 },
            { height: 2, kind: 'set-row', operationId: 'two', row: 1 },
          ],
          defaultXlsxWriteLimits(),
          PART,
        ),
      ).diagnostic.message,
    ).toBe('XLSX worksheet property patch targets must be unique');
  });

  it('ignores nested and foreign lexical lookalikes', () => {
    const xml = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:foreign"><wrapper><cols><col min="1" max="2"/></cols></wrapper><x:cols><x:col min="1" max="2"/></x:cols><cols><wrapper><col min="1" max="2"/></wrapper><other min="1" max="2"/><col min="1" max="9"/><col min="9" max="2"/><col min="1" max="2"/></cols><wrapper><row r="1"/></wrapper><sheetData><wrapper><row r="1"/></wrapper><other r="1"/><row r="1"><c r="A1"/></row></sheetData><wrapper><row r="1"/></wrapper></worksheet>`;
    const result = patchXlsxWorksheetProperties(
      bytes(xml),
      [
        { hidden: true, kind: 'set-row', operationId: 'row', row: 1 },
        {
          end: 2,
          hidden: true,
          kind: 'set-column',
          operationId: 'column',
          start: 1,
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const output = new TextDecoder().decode(result.data);
    expect(output.match(/hidden="1"/gu)).toHaveLength(2);
    expect(output).toContain(
      '<wrapper><cols><col min="1" max="2"/></cols></wrapper>',
    );
    expect(output).toContain('<col min="1" max="2" hidden="1"/></cols>');
    expect(output).toContain('<wrapper><row r="1"/></wrapper>');
    expect(output).toContain('<x:col min="1" max="2"/>');
  });

  it('rejects missing and ambiguous spans inside present collections', () => {
    for (const item of [
      {
        expected: {
          featureClass: 'missing-row-span',
          message: 'XLSX target row has no unique safe XML span',
          range: '2',
        },
        request: {
          hidden: true,
          kind: 'set-row' as const,
          operationId: 'row',
          row: 2,
        },
        xml: `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"/></sheetData></worksheet>`,
      },
      {
        expected: {
          featureClass: 'missing-column-span',
          message: 'XLSX target column range has no unique safe XML span',
          range: '2:2',
        },
        request: {
          end: 2,
          hidden: true,
          kind: 'set-column' as const,
          operationId: 'column',
          start: 2,
        },
        xml: `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><cols><col min="1" max="1"/></cols><sheetData/></worksheet>`,
      },
      {
        expected: {
          featureClass: 'missing-row-span',
          message: 'XLSX target row has no unique safe XML span',
          range: '1',
        },
        request: {
          hidden: true,
          kind: 'set-row' as const,
          operationId: 'row',
          row: 1,
        },
        xml: `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"/><row r="1"/></sheetData></worksheet>`,
      },
      {
        expected: {
          featureClass: 'missing-column-span',
          message: 'XLSX target column range has no unique safe XML span',
          range: '1:1',
        },
        request: {
          end: 1,
          hidden: true,
          kind: 'set-column' as const,
          operationId: 'column',
          start: 1,
        },
        xml: `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><cols><col min="1" max="1"/><col min="1" max="1"/></cols><sheetData/></worksheet>`,
      },
    ]) {
      expect(
        capture(() =>
          patchXlsxWorksheetProperties(
            bytes(item.xml),
            [item.request],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({
        ...item.expected,
        operationId: item.request.operationId,
      });
    }
  });

  it('enforces generated and aggregate patch byte limits exactly', () => {
    const request = [
      {
        height: 20,
        hidden: true,
        kind: 'set-row' as const,
        operationId: 'row',
        row: 1,
      },
      {
        end: 2,
        hidden: true,
        kind: 'set-column' as const,
        operationId: 'column',
        start: 1,
        width: 20,
      },
    ];
    const successful = patchXlsxWorksheetProperties(
      bytes(source()),
      request,
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(() =>
      patchXlsxWorksheetProperties(
        bytes(source()),
        request,
        {
          ...defaultXlsxWriteLimits(),
          maxGeneratedXmlBytes: successful.data.byteLength,
          maxPatchCount: successful.patchCount,
          maxPatchBytes: successful.patchBytes,
        },
        PART,
      ),
    ).not.toThrow();
    expect(
      capture(() =>
        patchXlsxWorksheetProperties(
          bytes(source()),
          request,
          {
            ...defaultXlsxWriteLimits(),
            maxPatchCount: successful.patchCount - 1,
          },
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({
      actual: successful.patchCount,
      limit: successful.patchCount - 1,
      limitName: 'maxPatchCount',
      part: PART,
    });
    for (const [limitName, limit] of [
      ['maxGeneratedXmlBytes', successful.data.byteLength - 1],
      ['maxPatchBytes', successful.patchBytes - 1],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxWorksheetProperties(
            bytes(source()),
            request,
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ limit, limitName, part: PART });
    }
  });
});
