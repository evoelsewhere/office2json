import { describe, expect, it } from 'vitest';

import { parseXlsx } from '../../src/formats/xlsx/parser';
import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import {
  appendXlsxStylesPart,
  xlsxAppendedStyleRecordCount,
} from '../../src/formats/xlsx/roundtrip/style-append';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import type { XlsxStyle } from '../../src/formats/xlsx/types';
import {
  createIndependentXlsx,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const PART = 'xl/styles.xml';
const BASE = `<styleSheet xmlns="${XLSX_SPREADSHEET_NS}"><numFmts count="0"/><fonts count="1"><font/></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs></styleSheet>`;

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
  throw new Error('Expected style append to fail');
}

describe('XLSX style append patching', () => {
  it('returns an isolated copy for an empty append and counts style records', () => {
    const source = bytes(BASE);
    const result = appendXlsxStylesPart(
      source,
      [],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(result).toEqual({
      cellXfIndexes: [],
      data: source,
      patchBytes: 0,
      patchCount: 0,
    });
    expect(result.data).not.toBe(source);
    expect(xlsxAppendedStyleRecordCount([{}, { numberFormat: '0.00' }])).toBe(
      8,
    );
    expect(xlsxAppendedStyleRecordCount([{ numberFormat: 'custom' }])).toBe(5);
  });

  it('appends built-in/custom formats, gradients, components, and XFs in order', async () => {
    const styles: XlsxStyle[] = [
      {
        alignment: { readingOrder: 'left-to-right' },
        fill: {
          angle: 45,
          kind: 'gradient',
          stops: [
            { color: { kind: 'automatic' }, position: 0 },
            { color: { argb: 'FFFFFFFF', kind: 'rgb' }, position: 1 },
          ],
          type: 'linear',
        },
        font: { underline: 'single' },
        numberFormat: '0.00',
        protection: { hidden: true },
      },
      {
        border: { start: { style: 'mediumDashDot' } },
        fill: {
          bottom: 1,
          kind: 'gradient',
          left: 0.25,
          right: 0.5,
          stops: [
            { color: { index: 1, kind: 'indexed' }, position: 0 },
            { color: { index: 2, kind: 'theme' }, position: 1 },
          ],
          top: 0.75,
          type: 'path',
        },
        font: { bold: true },
        numberFormat: '0.000 "x" &',
      },
    ];
    const result = appendXlsxStylesPart(
      bytes(BASE),
      styles,
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(result.cellXfIndexes).toEqual([1, 2]);
    expect(result.patchCount).toBe(9);
    expect(result.patchBytes).toBe(1064);
    const xml = new TextDecoder().decode(result.data);
    expect(xml).toContain('<numFmts count="1">');
    expect(xml).toContain('numFmtId="164"');
    expect(xml).toContain('numFmtId="2"');
    expect(xml).toContain('<fonts count="3">');
    expect(xml).toContain('<fills count="3">');
    expect(xml).toContain('<borders count="3">');
    expect(xml).toContain('<cellXfs count="3">');
    expect(xml).toContain('degree="45"');
    expect(xml).toContain('type="path"');
    const parsed = await parseXlsx(
      await createIndependentXlsx({ 'xl/styles.xml': result.data }),
      { errorMode: 'strict' },
    );
    expect(parsed.styles.slice(-2)).toEqual(styles);
  });

  it('serializes explicit clearing records against a non-empty base style', async () => {
    const source = BASE.replace(
      '<font/>',
      '<font><name val="Calibri"/></font>',
    ).replace(
      '<cellXfs count="1"><xf/></cellXfs>',
      '<cellXfs count="1"><xf fontId="0"/></cellXfs>',
    );
    const result = appendXlsxStylesPart(
      bytes(source),
      [
        {},
        { alignment: { horizontal: 'center' } },
        { fill: { kind: 'pattern', pattern: 'solid' } },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const parsed = await parseXlsx(
      await createIndependentXlsx({ 'xl/styles.xml': result.data }),
      { errorMode: 'strict' },
    );
    expect(parsed.styles).toEqual([
      { font: { name: 'Calibri' } },
      {},
      { alignment: { horizontal: 'center' } },
      { fill: { kind: 'pattern', pattern: 'solid' } },
    ]);
  });

  it('preserves prefixed UTF-16LE XML and inserts a missing numFmts collection', () => {
    const source = BASE.replaceAll('<styleSheet', '<s:styleSheet')
      .replaceAll('</styleSheet>', '</s:styleSheet>')
      .replaceAll('<numFmts count="0"/>', '')
      .replaceAll(
        /<(\/)?(fonts|font|fills|fill|patternFill|borders|border|cellStyleXfs|cellXfs|xf)(?=[\s/>])/gu,
        '<$1s:$2',
      )
      .replace('xmlns=', 'xmlns:s=');
    const encoded = new Uint8Array(2 + source.length * 2);
    encoded.set([0xff, 0xfe]);
    for (let index = 0; index < source.length; index += 1) {
      const code = source.charCodeAt(index);
      encoded[2 + index * 2] = code & 0xff;
      encoded[3 + index * 2] = code >>> 8;
    }
    const result = appendXlsxStylesPart(
      encoded,
      [{ numberFormat: 'custom-a' }, { numberFormat: 'custom-b' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect([...result.data.slice(0, 2)]).toEqual([0xff, 0xfe]);
    const xml = new TextDecoder('utf-16le').decode(result.data.slice(2));
    expect(xml).toContain('<s:numFmts count="2"><s:numFmt');
    expect(xml).toContain('formatCode="custom-a"/><s:numFmt');
    expect(xml).toContain('formatCode="custom-b"/></s:numFmts>');
    expect(xml).toContain('<s:cellXfs count="3">');
  });

  it('escapes every XML attribute character and validates XML boundaries', () => {
    const result = appendXlsxStylesPart(
      bytes(BASE),
      [
        {
          font: {
            name: `&<"\r\n\t\ud7ff\ue000\ufffd${String.fromCodePoint(0x1_0000)}😀`,
          },
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toContain(
      `val="&amp;&lt;&quot;&#13;&#10;&#9;\ud7ff\ue000\ufffd${String.fromCodePoint(0x1_0000)}😀"`,
    );
    for (const value of ['bad\u0001', 'bad\ud800', 'bad\ufffe']) {
      expect(
        capture(() =>
          appendXlsxStylesPart(
            bytes(BASE),
            [{ font: { name: value } }],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({
        code: 'invalid-roundtrip-json',
        message: 'XLSX style text contains an invalid XML character',
        part: PART,
      });
    }
  });

  it('enforces generated XML and aggregate patch bytes at exact boundaries', () => {
    const successful = appendXlsxStylesPart(
      bytes(BASE),
      [{ font: { bold: true } }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(successful.patchCount).toBe(8);
    expect(() =>
      appendXlsxStylesPart(
        bytes(BASE),
        [{ font: { bold: true } }],
        {
          ...defaultXlsxWriteLimits(),
          maxGeneratedXmlBytes: successful.data.byteLength,
          maxPatchBytes: successful.patchBytes,
        },
        PART,
      ),
    ).not.toThrow();
    for (const [limitName, limit] of [
      ['maxGeneratedXmlBytes', successful.data.byteLength - 1],
      ['maxPatchBytes', successful.patchBytes - 1],
    ] as const) {
      expect(
        capture(() =>
          appendXlsxStylesPart(
            bytes(BASE),
            [{ font: { bold: true } }],
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({
        code: 'resource-limit-exceeded',
        limit,
        limitName,
        part: PART,
      });
    }
  });

  it('rejects unsafe structure, exhausted IDs, and invalid XML characters', () => {
    expect(
      capture(() =>
        appendXlsxStylesPart(
          bytes('<notStyles/>'),
          [{}],
          defaultXlsxWriteLimits(),
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({
      code: 'preservation-conflict',
      featureClass: 'styles-part',
      message: 'XLSX styles root cannot be patched',
    });
    for (const [source, message] of [
      [
        BASE.replace(/<fonts[\s\S]*?<\/fonts>/u, ''),
        'XLSX styles fonts collection is missing',
      ],
      [
        BASE.replace(/<cellXfs[\s\S]*?<\/cellXfs>/u, ''),
        'XLSX styles cellXfs collection is missing',
      ],
      [
        BASE.replace('<fonts count="1">', '<fonts>'),
        'XLSX styles fonts count cannot be patched',
      ],
    ] as const) {
      expect(
        capture(() =>
          appendXlsxStylesPart(
            bytes(source),
            [{}],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({
        code: 'preservation-conflict',
        featureClass: 'styles-part',
        message,
      });
    }
    const exhausted = BASE.replace(
      '<numFmts count="0"/>',
      '<numFmts count="1"><numFmt numFmtId="4294967295" formatCode="x"/></numFmts>',
    );
    expect(
      capture(() =>
        appendXlsxStylesPart(
          bytes(exhausted),
          [{ numberFormat: 'custom' }],
          defaultXlsxWriteLimits(),
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({
      code: 'identifier-allocation-failed',
      featureClass: 'number-format',
      message: 'XLSX custom number-format IDs are exhausted',
      part: PART,
    });
    expect(
      capture(() =>
        appendXlsxStylesPart(
          bytes(BASE),
          [{ font: { name: 'bad\u0000name' } }],
          defaultXlsxWriteLimits(),
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({ code: 'invalid-roundtrip-json', part: PART });
  });

  it('ignores nested collection and number-format lookalikes', () => {
    const source = BASE.replace(
      '<numFmts count="0"/>',
      '<before><numFmt numFmtId="4294967295"/></before><numFmts count="3"><wrapper><numFmt numFmtId="4294967295"/></wrapper><other numFmtId="4294967295"/><numFmt/><numFmt other="x" numFmtId="165"/><numFmt numFmtId="166"/></numFmts><after><numFmt numFmtId="4294967295"/></after>',
    ).replace(
      '<fonts count="1">',
      '<x:fonts xmlns:x="urn:foreign" count="99"/><fonts count="1">',
    );
    const result = appendXlsxStylesPart(
      bytes(source),
      [{ numberFormat: 'custom' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    const xml = new TextDecoder().decode(result.data);
    expect(xml).toContain('<fonts count="2">');
    expect(xml).toContain('numFmtId="167"');
    expect(xml).toContain('<wrapper><numFmt numFmtId="4294967295"/>');
    expect(xml).toContain('<x:fonts xmlns:x="urn:foreign" count="99"/>');
  });

  it('patches self-closing collections with attribute-order and anchor traps', () => {
    const source = BASE.replace(
      '<numFmts count="0"/>',
      '<numFmts note="/>" count="0" /   >',
    );
    const result = appendXlsxStylesPart(
      bytes(source),
      [{ numberFormat: 'first' }, { numberFormat: 'second' }],
      defaultXlsxWriteLimits(),
      PART,
    );
    const xml = new TextDecoder().decode(result.data);
    expect(xml).toContain('<numFmts note="/>" count="2" >');
    expect(xml).toContain('formatCode="first"/><numFmt');
    expect(xml).toContain('formatCode="second"/></numFmts>');
  });

  it('appends after nested same-name elements using the collection depth', () => {
    const source = BASE.replace(
      '<font/>',
      '<font><fonts><font/></fonts></font>',
    );
    const result = appendXlsxStylesPart(
      bytes(source),
      [{}],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toContain(
      '<font><fonts><font/></fonts></font><font/></fonts>',
    );
  });

  it('rejects a nested styleSheet instead of treating it as the document root', () => {
    const nested = `<outer>${BASE}</outer>`;
    expect(
      capture(() =>
        appendXlsxStylesPart(
          bytes(nested),
          [{}],
          defaultXlsxWriteLimits(),
          PART,
        ),
      ).diagnostic.message,
    ).toBe('XLSX styles root cannot be patched');
  });
});
