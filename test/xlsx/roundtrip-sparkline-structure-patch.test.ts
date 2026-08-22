import { describe, expect, it } from 'vitest';

import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import { patchXlsxSparklineStructure } from '../../src/formats/xlsx/roundtrip/sparkline-structure-patch';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';

const PART = 'xl/worksheets/sheet1.xml';
const WORKSHEET_NAMESPACE =
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const X14_NAMESPACE =
  'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM_NAMESPACE = 'http://schemas.microsoft.com/office/excel/2006/main';
const URI = '{05c60535-1f16-4fd2-b633-f4f36f0b64e0}';

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
  throw new Error('Expected sparkline structural patch to fail');
}

function sparkline(formula: string, location: string): string {
  return `<s:sparkline><m:f>${formula}</m:f><m:sqref>${location}</m:sqref></s:sparkline>`;
}

function extension(entries: string, uri = URI, attributes = ''): string {
  return `<w:ext${attributes} uri="${uri}"><s:sparklineGroups xmlns:s="${X14_NAMESPACE}"><s:sparklineGroup><s:sparklines xmlns:m="${XM_NAMESPACE}">${entries}</s:sparklines></s:sparklineGroup></s:sparklineGroups></w:ext>`;
}

function worksheet(entries: string, extra = '', extraExtensions = ''): string {
  return `<w:worksheet xmlns:w="${WORKSHEET_NAMESPACE}"><w:sheetData/>${extra}<w:extLst>${extraExtensions}${extension(entries)}</w:extLst></w:worksheet>`;
}

describe('XLSX sparkline structural patching', () => {
  it('rewrites source ranges and locations through inherited namespaces', () => {
    const source = worksheet(
      `${sparkline("'R&amp;D'!$A$1:$A$3", 'B2')}${sparkline('Other!C1:C3', 'D2')}`,
    );
    const result = patchXlsxSparklineStructure(
      bytes(source),
      [
        { count: 1, index: 2, kind: 'insert-rows', operationId: 'rows' },
        {
          count: 2,
          index: 2,
          kind: 'insert-columns',
          operationId: 'columns',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
      'R&D',
    );
    const output = new TextDecoder().decode(result.data);
    expect(output).toContain(sparkline("'R&amp;D'!$A$1:$A$4", 'D3'));
    expect(output).toContain(sparkline('Other!C1:C3', 'F3'));
    expect(result.patchCount).toBe(5);
    expect(result.patchBytes).toBe(
      ["'R&amp;D'!$A$1:$A$4", 'B3', 'D3', 'D3', 'F3'].reduce(
        (total, value) => total + new TextEncoder().encode(value).byteLength,
        0,
      ),
    );
  });

  it('selects the exact extension URI and preserves unrelated entries', () => {
    const unrelated = `<w:ext uri="urn:other"><foreign xmlns="urn:foreign"/></w:ext><w:ext other="first"/>`;
    const source = worksheet(sparkline('A1:A3', 'B1'), '', unrelated).replace(
      `uri="${URI}"`,
      `other="x" uri="${URI.toUpperCase()}"`,
    );
    const output = new TextDecoder().decode(
      patchXlsxSparklineStructure(
        bytes(source),
        [{ count: 1, index: 2, kind: 'insert-rows', operationId: 'uri' }],
        defaultXlsxWriteLimits(),
        PART,
        'Sheet1',
      ).data,
    );
    expect(output).toContain(unrelated);
    expect(output).toContain(sparkline('A1:A4', 'B1'));

    for (const noSparkline of [
      `<w:worksheet xmlns:w="${WORKSHEET_NAMESPACE}"><w:sheetData/></w:worksheet>`,
      `<w:worksheet xmlns:w="${WORKSHEET_NAMESPACE}"><w:sheetData/><w:extLst>${unrelated}</w:extLst></w:worksheet>`,
    ]) {
      expect(
        new TextDecoder().decode(
          patchXlsxSparklineStructure(
            bytes(noSparkline),
            [
              {
                count: 1,
                index: 1,
                kind: 'insert-rows',
                operationId: 'no-sparkline',
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
            'Sheet1',
          ).data,
        ),
      ).toBe(noSparkline);
    }
  });

  it('re-escapes transformed sheet qualifiers as XML text', () => {
    const source = worksheet(sparkline("'A&lt;B&gt;&#13;C'!A1:A3", 'D1'));
    const output = new TextDecoder().decode(
      patchXlsxSparklineStructure(
        bytes(source),
        [{ count: 1, index: 2, kind: 'insert-rows', operationId: 'entities' }],
        defaultXlsxWriteLimits(),
        PART,
        'A<B>\rC',
      ).data,
    );
    expect(output).toContain(sparkline("'A&lt;B&gt;&#13;C'!A1:A4", 'D1'));
  });

  it('ignores nested and foreign-namespace lookalikes', () => {
    const foreign = `<x:sparkline xmlns:x="urn:foreign"><m:f xmlns:m="${XM_NAMESPACE}">A1:A3</m:f><m:sqref xmlns:m="${XM_NAMESPACE}">B1</m:sqref></x:sparkline>`;
    const nested = `<s:wrapper xmlns:s="${X14_NAMESPACE}">${sparkline('A1:A3', 'C1')}</s:wrapper>`;
    const source = worksheet(`${sparkline('A1:A3', 'B1')}${nested}`, foreign);
    const output = new TextDecoder().decode(
      patchXlsxSparklineStructure(
        bytes(source),
        [{ count: 1, index: 2, kind: 'insert-rows', operationId: 'owned' }],
        defaultXlsxWriteLimits(),
        PART,
        'Sheet1',
      ).data,
    );
    expect(output).toContain(sparkline('A1:A4', 'B1'));
    expect(output).toContain(foreign);
    expect(output).toContain(nested);
  });

  it('fails closed for unsupported formulas and deleted dependencies', () => {
    for (const [formula, location, request, featureClass, code] of [
      [
        'SUM(A1:A3)',
        'B1',
        { count: 1, index: 1, kind: 'insert-rows', operationId: 'formula' },
        'sparkline-formula-reference',
        'formula-rewrite-unsupported',
      ],
      [
        'A2',
        'B1',
        { count: 1, index: 2, kind: 'delete-rows', operationId: 'source' },
        'sparkline-source-deletion',
        'formula-rewrite-unsupported',
      ],
      [
        'A1:A3',
        'B2',
        { count: 1, index: 2, kind: 'delete-rows', operationId: 'location' },
        'sparkline-location-deletion',
        'preservation-conflict',
      ],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxSparklineStructure(
            bytes(worksheet(sparkline(formula, location))),
            [request],
            defaultXlsxWriteLimits(),
            PART,
            'Sheet1',
          ),
        ).diagnostic,
      ).toMatchObject({ code, featureClass, operationId: request.operationId });
      expect(
        capture(() =>
          patchXlsxSparklineStructure(
            bytes(worksheet(sparkline(formula, location))),
            [request],
            defaultXlsxWriteLimits(),
            PART,
            'Sheet1',
          ),
        ).diagnostic,
      ).toMatchObject({
        message:
          featureClass === 'sparkline-formula-reference'
            ? 'XLSX structural sparkline formula is unsupported'
            : featureClass === 'sparkline-source-deletion'
              ? 'XLSX structural edit would delete a sparkline source'
              : 'XLSX structural edit would delete a sparkline location',
        range: `${request.index}:${request.index + request.count - 1}`,
      });
    }
  });

  it('rejects malformed roots, namespace spoofing, and graph cardinality', () => {
    for (const source of [
      '<wrong/>',
      `<wrapper><w:worksheet xmlns:w="${WORKSHEET_NAMESPACE}"><w:sheetData/></w:worksheet></wrapper>`,
      `<worksheet><sheetData/></worksheet>`,
      `<w:worksheet xmlns:w="${WORKSHEET_NAMESPACE}"><w:sheetData/><w:extLst/><w:extLst/></w:worksheet>`,
      worksheet(
        sparkline('A1:A3', 'B1'),
        '',
        extension(sparkline('C1:C3', 'D1')),
      ),
      worksheet(''),
      worksheet('<s:sparkline/>'),
      worksheet(
        `<s:sparkline><m:f/><m:sqref xmlns:m="${XM_NAMESPACE}">B1</m:sqref></s:sparkline>`,
      ),
      worksheet(
        `<s:sparkline><m:f xmlns:m="${XM_NAMESPACE}"><m:nested/></m:f><m:sqref xmlns:m="${XM_NAMESPACE}">B1</m:sqref></s:sparkline>`,
      ),
      worksheet(
        `<s:sparkline><x:f xmlns:x="urn:foreign">A1:A3</x:f><m:sqref xmlns:m="${XM_NAMESPACE}">B1</m:sqref></s:sparkline>`,
      ),
    ]) {
      expect(
        capture(() =>
          patchXlsxSparklineStructure(
            bytes(source),
            [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'bad' }],
            defaultXlsxWriteLimits(),
            PART,
            'Sheet1',
          ),
        ).diagnostic,
      ).toMatchObject({ featureClass: 'sparkline-structure-xml' });
    }
    for (const [source, message] of [
      ['<wrong/>', 'XLSX worksheet root cannot patch sparklines'],
      [
        '<worksheet><sheetData/></worksheet>',
        'XLSX worksheet namespace cannot patch sparklines',
      ],
      [
        worksheet('<s:sparkline/>'),
        'XLSX structural sparkline graph is invalid',
      ],
      [
        worksheet(
          `<s:sparkline><m:f/><m:sqref xmlns:m="${XM_NAMESPACE}">B1</m:sqref></s:sparkline>`,
        ),
        'XLSX structural sparkline text is invalid',
      ],
      [
        worksheet(
          `<s:sparkline><m:f xmlns:m="${XM_NAMESPACE}"><m:nested/></m:f><m:sqref xmlns:m="${XM_NAMESPACE}">B1</m:sqref></s:sparkline>`,
        ),
        'XLSX structural sparkline text is invalid',
      ],
      [
        `<w:worksheet xmlns:w="${WORKSHEET_NAMESPACE}"><w:sheetData/><w:extLst/><w:extLst/></w:worksheet>`,
        'XLSX structural sparkline graph is invalid',
      ],
      [
        worksheet(
          sparkline('A1:A3', 'B1'),
          '',
          extension(sparkline('C1:C3', 'D1')),
        ),
        'XLSX structural sparkline graph is invalid',
      ],
      [worksheet(''), 'XLSX structural sparkline graph is invalid'],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxSparklineStructure(
            bytes(source),
            [
              {
                count: 1,
                index: 1,
                kind: 'insert-rows',
                operationId: 'message',
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
            'Sheet1',
          ),
        ).diagnostic,
      ).toMatchObject({ message });
    }
    for (const location of ['$B$1', 'B1:B2']) {
      expect(
        capture(() =>
          patchXlsxSparklineStructure(
            bytes(worksheet(sparkline('A1:A3', location))),
            [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'bad' }],
            defaultXlsxWriteLimits(),
            PART,
            'Sheet1',
          ),
        ).diagnostic,
      ).toMatchObject({
        message: 'XLSX structural sparkline location is invalid',
      });
    }
  });

  it('preserves no-op bytes and enforces exact resource limits', () => {
    const source = worksheet(sparkline('A1:A3', 'B1'));
    const unchanged = patchXlsxSparklineStructure(
      bytes(source),
      [{ count: 1, index: 5, kind: 'insert-rows', operationId: 'outside' }],
      defaultXlsxWriteLimits(),
      PART,
      'Sheet1',
    );
    expect(new TextDecoder().decode(unchanged.data)).toBe(source);
    expect(unchanged.patchCount).toBe(0);
    const input = bytes(source);
    const empty = patchXlsxSparklineStructure(
      input,
      [],
      defaultXlsxWriteLimits(),
      PART,
      'Sheet1',
    );
    expect(empty.data).toEqual(input);
    expect(empty.data).not.toBe(input);
    const request = {
      count: 1,
      index: 2,
      kind: 'insert-rows' as const,
      operationId: 'limits',
    };
    const changed = patchXlsxSparklineStructure(
      bytes(source),
      [request],
      defaultXlsxWriteLimits(),
      PART,
      'Sheet1',
    );
    expect(changed.patchCount).toBe(1);
    expect(changed.patchBytes).toBe(
      new TextEncoder().encode('A1:A4').byteLength,
    );
    expect(() =>
      patchXlsxSparklineStructure(
        bytes(source),
        [request],
        {
          ...defaultXlsxWriteLimits(),
          maxGeneratedXmlBytes: changed.data.byteLength,
          maxPatchBytes: changed.patchBytes,
          maxPatchCount: changed.patchCount,
        },
        PART,
        'Sheet1',
      ),
    ).not.toThrow();
    for (const [limitName, limit] of [
      ['maxGeneratedXmlBytes', changed.data.byteLength - 1],
      ['maxPatchBytes', changed.patchBytes - 1],
      ['maxPatchCount', changed.patchCount - 1],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxSparklineStructure(
            bytes(source),
            [request],
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            PART,
            'Sheet1',
          ),
        ).diagnostic,
      ).toMatchObject({ limitName });
    }
    const variableSource = worksheet(
      `${sparkline('A1:A3', 'B2')}${sparkline('A10:A30', 'C20')}`,
    );
    const variable = patchXlsxSparklineStructure(
      bytes(variableSource),
      [
        {
          count: 100,
          index: 2,
          kind: 'insert-rows',
          operationId: 'variable-length',
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
      'Sheet1',
    );
    expect(new TextDecoder().decode(variable.data)).toBe(
      worksheet(
        `${sparkline('A1:A103', 'B102')}${sparkline('A110:A130', 'C120')}`,
      ),
    );
    expect(variable.patchCount).toBe(4);
    expect(variable.patchBytes).toBe(
      ['A1:A103', 'B102', 'A110:A130', 'C120'].reduce(
        (total, value) => total + new TextEncoder().encode(value).byteLength,
        0,
      ),
    );
  });
});
