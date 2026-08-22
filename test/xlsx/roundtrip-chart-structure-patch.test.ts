import { describe, expect, it } from 'vitest';

import { patchXlsxChartStructure } from '../../src/formats/xlsx/roundtrip/chart-structure-patch';
import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';

const PART = 'xl/charts/chart1.xml';
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

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
  throw new Error('Expected chart structural patch to fail');
}

function source(kind: 'numRef' | 'strRef', formula: string): string {
  const cache = kind === 'numRef' ? 'numCache' : 'strCache';
  return `<c:${kind}><c:f>${formula}</c:f><c:${cache}><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:${cache}></c:${kind}>`;
}

function multiSource(formula: string): string {
  return `<c:multiLvlStrRef><c:f>${formula}</c:f><c:multiLvlStrCache><c:ptCount val="2"/><c:lvl><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef>`;
}

function chart(content: string, namespace = CHART_NS): string {
  return `<c:chartSpace xmlns:c="${namespace}"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>${content}</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
}

describe('XLSX chart structural patching', () => {
  it('moves supported references without changing their caches', () => {
    const input = chart(
      `<c:tx>${source('strRef', 'Sheet1!$D$2')}</c:tx><c:cat>${multiSource('Sheet1!$A$2:$A$3')}</c:cat><c:val>${source('numRef', 'Sheet1!$B$2:$B$3')}</c:val><c:xVal>${source('numRef', 'Sheet1!$C$2:$C$3')}</c:xVal><c:yVal>${source('numRef', 'Sheet1!$E$2:$E$3')}</c:yVal><c:bubbleSize>${source('numRef', 'Sheet1!$F$2:$F$3')}</c:bubbleSize>`,
    ).replace(
      `xmlns:c="${CHART_NS}"`,
      `xmlns:c="${CHART_NS}" xxxxxxc="urn:wrong"`,
    );
    const result = patchXlsxChartStructure(
      bytes(input),
      [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'move-chart' }],
      defaultXlsxWriteLimits(),
      PART,
      'Sheet1',
    );
    const output = new TextDecoder().decode(result.data);
    expect(output).toContain('Sheet1!$A$3:$A$4');
    expect(output).toContain('Sheet1!$B$3:$B$4');
    expect(output).toContain('Sheet1!$D$3');
    expect(output).toContain('Sheet1!$E$3:$E$4');
    expect(output).toContain('Sheet1!$F$3:$F$4');
    expect(output).toContain('Sheet1!$C$3:$C$4');
    expect(output).toContain('<c:ptCount val="2"/>');
    expect(result.patchCount).toBe(6);
    expect(result.patchBytes).toBe(
      [
        'Sheet1!$D$3',
        'Sheet1!$A$3:$A$4',
        'Sheet1!$B$3:$B$4',
        'Sheet1!$C$3:$C$4',
        'Sheet1!$E$3:$E$4',
        'Sheet1!$F$3:$F$4',
      ].reduce(
        (total, value) => total + new TextEncoder().encode(value).byteLength,
        0,
      ),
    );
  });

  it('resolves inherited prefixes and ignores nested or foreign lookalikes', () => {
    const owned = source('strRef', 'Sheet1!A2:A3');
    const nested = `<c:wrapper>${source('strRef', 'Sheet1!C2:C3')}</c:wrapper>`;
    const foreign = `<x:strRef xmlns:x="urn:foreign"><c:f>Sheet1!D2:D3</c:f></x:strRef>`;
    const foreignFormula = `<c:strRef><x:f xmlns:x="urn:foreign">Sheet1!E2:E3</x:f></c:strRef>`;
    const foreignGrandparent = `<x:cat xmlns:x="urn:foreign">${source('strRef', 'Sheet1!F2:F3')}</x:cat>`;
    const wrongParent = `<c:cat><c:wrapper><c:f>Sheet1!G2:G3</c:f></c:wrapper></c:cat>`;
    const input = chart(
      `<c:cat>${nested}${owned}${foreign}${foreignFormula}</c:cat>${foreignGrandparent}${wrongParent}`,
    );
    const output = new TextDecoder().decode(
      patchXlsxChartStructure(
        bytes(input),
        [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'owned' }],
        defaultXlsxWriteLimits(),
        PART,
        'Sheet1',
      ).data,
    );
    expect(output).toContain('Sheet1!A3:A4');
    expect(output).toContain(nested);
    expect(output).toContain(foreign);
    expect(output).toContain(foreignFormula);
    expect(output).toContain(foreignGrandparent);
    expect(output).toContain(wrongParent);
  });

  it('blocks cache cardinality changes, source deletion, and unsupported syntax', () => {
    for (const [formula, request, featureClass, code, message] of [
      [
        'Sheet1!A2:A3',
        { count: 1, index: 3, kind: 'insert-rows', operationId: 'cache' },
        'chart-cache-cardinality',
        'preservation-conflict',
        'XLSX structural chart cache cardinality would change',
      ],
      [
        'Sheet1!A2:A3',
        { count: 2, index: 2, kind: 'delete-rows', operationId: 'delete' },
        'chart-source-deletion',
        'formula-rewrite-unsupported',
        'XLSX structural edit would delete a chart source',
      ],
      [
        'SUM(A1:A3)',
        { count: 1, index: 1, kind: 'insert-rows', operationId: 'formula' },
        'chart-formula-reference',
        'formula-rewrite-unsupported',
        'XLSX structural chart formula is unsupported',
      ],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxChartStructure(
            bytes(chart(`<c:cat>${source('strRef', formula)}</c:cat>`)),
            [request],
            defaultXlsxWriteLimits(),
            PART,
            'Sheet1',
          ),
        ).diagnostic,
      ).toMatchObject({
        code,
        featureClass,
        message,
        operationId: request.operationId,
        range: `${request.index}:${request.index + request.count - 1}`,
      });
    }
  });

  it('rejects malformed roots, namespaces, and formula text', () => {
    for (const [input, message] of [
      ['<wrong/>', 'XLSX chart root cannot patch structure'],
      [
        `<wrapper><c:chartSpace xmlns:c="${CHART_NS}"/></wrapper>`,
        'XLSX chart root cannot patch structure',
      ],
      [chart('', 'urn:wrong'), 'XLSX chart namespace cannot patch structure'],
      [
        chart('<c:cat><c:strRef><c:f/></c:strRef></c:cat>'),
        'XLSX structural chart formula is invalid',
      ],
      [
        chart('<c:cat><c:strRef><c:f><c:nested/></c:f></c:strRef></c:cat>'),
        'XLSX structural chart formula is invalid',
      ],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxChartStructure(
            bytes(input),
            [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'bad' }],
            defaultXlsxWriteLimits(),
            PART,
            'Sheet1',
          ),
        ).diagnostic,
      ).toMatchObject({ featureClass: 'chart-structure-xml', message });
    }
  });

  it('supports Strict ChartML and preserves exact no-op bytes', () => {
    const strict = 'http://purl.oclc.org/ooxml/drawingml/chart';
    const input = chart(
      `<c:cat>${source('strRef', 'Other!A1:A2')}</c:cat>`,
      strict,
    );
    const result = patchXlsxChartStructure(
      bytes(input),
      [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'strict' }],
      defaultXlsxWriteLimits(),
      PART,
      'Sheet1',
    );
    expect(new TextDecoder().decode(result.data)).toBe(input);
    expect(result.patchCount).toBe(0);
    const original = bytes(input);
    const empty = patchXlsxChartStructure(
      original,
      [],
      defaultXlsxWriteLimits(),
      PART,
      'Sheet1',
    );
    expect(empty.data).toEqual(original);
    expect(empty.data).not.toBe(original);

    const unprefixed = chart(
      `<c:cat>${source('strRef', 'Sheet1!A2:A3')}</c:cat>`,
    )
      .replaceAll('c:', '')
      .replace('xmlns:c=', 'xmlns=');
    expect(
      new TextDecoder().decode(
        patchXlsxChartStructure(
          bytes(unprefixed),
          [
            {
              count: 1,
              index: 1,
              kind: 'insert-rows',
              operationId: 'default-prefix',
            },
          ],
          defaultXlsxWriteLimits(),
          PART,
          'Sheet1',
        ).data,
      ),
    ).toContain('Sheet1!A3:A4');
    const alternate = chart(
      `<c:cat>${source('strRef', 'Sheet1!A2:A3')}</c:cat>`,
    )
      .replaceAll('c:', 'z:')
      .replace('xmlns:c=', 'xmlns:z=');
    expect(
      new TextDecoder().decode(
        patchXlsxChartStructure(
          bytes(alternate),
          [
            {
              count: 1,
              index: 1,
              kind: 'insert-rows',
              operationId: 'alternate-prefix',
            },
          ],
          defaultXlsxWriteLimits(),
          PART,
          'Sheet1',
        ).data,
      ),
    ).toContain('Sheet1!A3:A4');
  });

  it('re-escapes transformed qualifiers as ChartML text', () => {
    const input = chart(
      `<c:cat>${source('strRef', "'A&lt;B&gt;&#13;&amp;D'!A2:A3")}</c:cat>`,
    );
    const output = new TextDecoder().decode(
      patchXlsxChartStructure(
        bytes(input),
        [{ count: 1, index: 1, kind: 'insert-rows', operationId: 'entities' }],
        defaultXlsxWriteLimits(),
        PART,
        'A<B>\r&D',
      ).data,
    );
    expect(output).toContain("'A&lt;B&gt;&#13;&amp;D'!A3:A4");
  });

  it('enforces exact generated and patch resource boundaries', () => {
    const input = chart(`<c:cat>${source('strRef', 'Sheet1!A2:A3')}</c:cat>`);
    const request = {
      count: 100,
      index: 1,
      kind: 'insert-rows' as const,
      operationId: 'limits',
    };
    const changed = patchXlsxChartStructure(
      bytes(input),
      [request],
      defaultXlsxWriteLimits(),
      PART,
      'Sheet1',
    );
    expect(new TextDecoder().decode(changed.data)).toContain(
      'Sheet1!A102:A103',
    );
    expect(() =>
      patchXlsxChartStructure(
        bytes(input),
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
          patchXlsxChartStructure(
            bytes(input),
            [request],
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            PART,
            'Sheet1',
          ),
        ).diagnostic,
      ).toMatchObject({ limitName });
    }
    const variableInput = chart(
      `<c:cat>${source('strRef', 'Sheet1!A2:A3')}</c:cat><c:val>${source('numRef', 'Sheet1!B10:B30')}</c:val>`,
    );
    const variable = patchXlsxChartStructure(
      bytes(variableInput),
      [request],
      defaultXlsxWriteLimits(),
      PART,
      'Sheet1',
    );
    expect(new TextDecoder().decode(variable.data)).toBe(
      chart(
        `<c:cat>${source('strRef', 'Sheet1!A102:A103')}</c:cat><c:val>${source('numRef', 'Sheet1!B110:B130')}</c:val>`,
      ),
    );
  });
});
