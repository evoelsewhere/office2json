import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { parseXlsx } from '../../src/formats/xlsx/parser';
import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import type { XlsxWorksheetCellPatch } from '../../src/formats/xlsx/roundtrip/worksheet-patch';
import {
  escapeXlsxCellText,
  patchXlsxWorksheetPart,
  patchXlsxWorksheetPartWithReport,
} from '../../src/formats/xlsx/roundtrip/worksheet-patch';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import type { XlsxCell } from '../../src/formats/xlsx/types';
import {
  createIndependentXlsx,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const PART = 'xl/worksheets/sheet1.xml';
const limits = defaultXlsxWriteLimits();

function worksheetXml(cells: string): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1">${cells}</row></sheetData></worksheet>`;
}

function cell(
  address: string,
  column: number,
  content: XlsxCell['content'],
): XlsxCell {
  return { address, column, content } as XlsxCell;
}

function requested(
  target: XlsxCell,
  operationId = 'edit-1',
): XlsxWorksheetCellPatch {
  return { cell: target, operationId };
}

function capture(action: () => unknown): XlsxWriteError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected XLSX worksheet patch to fail');
}

async function patchWorkbook(
  bytes: Uint8Array,
  patches: readonly XlsxWorksheetCellPatch[],
): Promise<{ data: Uint8Array; xml: string }> {
  const archive = await JSZip.loadAsync(bytes);
  const entry = archive.file(PART)!;
  const source = await entry.async('uint8array');
  const patched = patchXlsxWorksheetPart(source, patches, limits, PART);
  archive.file(PART, patched, { date: entry.date });
  const data = await archive.generateAsync({
    compression: 'DEFLATE',
    type: 'uint8array',
  });
  return { data, xml: new TextDecoder().decode(patched) };
}

function encodeUtf16(
  text: string,
  encoding: 'utf-16be' | 'utf-16le',
  bom: boolean,
): Uint8Array {
  const offset = bom ? 2 : 0;
  const output = new Uint8Array(offset + text.length * 2);
  if (bom) {
    output[0] = encoding === 'utf-16le' ? 0xff : 0xfe;
    output[1] = encoding === 'utf-16le' ? 0xfe : 0xff;
  }
  for (let index = 0; index < text.length; index += 1) {
    const value = text.charCodeAt(index);
    const byte = offset + index * 2;
    output[byte + (encoding === 'utf-16le' ? 0 : 1)] = value & 0xff;
    output[byte + (encoding === 'utf-16le' ? 1 : 0)] = value >>> 8;
  }
  return output;
}

describe('XLSX worksheet cell patching', () => {
  it.each([
    [
      'text',
      cell('A1', 1, {
        kind: 'value',
        value: { kind: 'text', text: ' <updated> &\rnext ' },
      }),
      {
        kind: 'value',
        value: { kind: 'text', text: ' <updated> &\rnext ' },
      },
    ],
    [
      'number',
      cell('A1', 1, {
        kind: 'value',
        value: { kind: 'number', value: -12.5 },
      }),
      { kind: 'value', value: { kind: 'number', value: -12.5 } },
    ],
    [
      'boolean',
      cell('A1', 1, {
        kind: 'value',
        value: { kind: 'boolean', value: false },
      }),
      { kind: 'value', value: { kind: 'boolean', value: false } },
    ],
    [
      'error',
      cell('A1', 1, {
        kind: 'value',
        value: { code: '#REF!', kind: 'error' },
      }),
      { kind: 'value', value: { code: '#REF!', kind: 'error' } },
    ],
    [
      'formula',
      cell('A1', 1, {
        cached: { kind: 'missing' },
        formula: { expression: 'SUM(B1:C1)<3', kind: 'normal' },
        kind: 'formula',
      }),
      {
        cached: { kind: 'missing' },
        formula: { expression: 'SUM(B1:C1)<3', kind: 'normal' },
        kind: 'formula',
      },
    ],
    ['blank', cell('A1', 1, { kind: 'blank' }), { kind: 'blank' }],
  ] as const)(
    'patches and strictly reparses %s content',
    async (_name, target, expected) => {
      const bytes = await createIndependentXlsx({
        [PART]: worksheetXml('<c r="A1" t="s"><v>0</v></c>'),
      });
      const before = bytes.slice();
      const output = await patchWorkbook(bytes, [requested(target)]);
      const parsed = await parseXlsx(output.data, { errorMode: 'strict' });
      const sheet = parsed.sheets[0]!;
      expect(sheet.kind).toBe('worksheet');
      expect(
        sheet.kind === 'worksheet' ? sheet.rows[0]?.cells[0]?.content : null,
      ).toEqual(expected);
      expect(bytes).toEqual(before);
      expect(output.data).not.toBe(bytes);
    },
  );

  it('patches multiple cells in descending non-overlapping spans', async () => {
    const bytes = await createIndependentXlsx({
      [PART]: worksheetXml(
        '<c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c>',
      ),
    });
    const output = await patchWorkbook(bytes, [
      requested(
        cell('A1', 1, {
          kind: 'value',
          value: { kind: 'number', value: 10 },
        }),
        'first',
      ),
      requested(
        cell('C1', 3, {
          kind: 'value',
          value: { kind: 'number', value: 30 },
        }),
        'last',
      ),
    ]);
    expect(output.xml).toContain('<c r="A1"><v>10</v></c>');
    expect(output.xml).toContain('<c r="B1"><v>2</v></c>');
    expect(output.xml).toContain('<c r="C1"><v>30</v></c>');
  });

  it('matches the requested second cell and its r attribute exactly', () => {
    const xml = worksheetXml(
      '<other r="B1"/><c r="A1"><v>1</v></c><c s="0" r="B1"><v>2</v></c>',
    );
    const output = patchXlsxWorksheetPart(
      new TextEncoder().encode(xml),
      [
        requested(
          cell('B1', 2, {
            kind: 'value',
            value: { kind: 'number', value: 20 },
          }),
        ),
      ],
      limits,
      PART,
    );
    expect(new TextDecoder().decode(output)).toBe(
      xml.replace(
        '<c s="0" r="B1"><v>2</v></c>',
        '<c s="0" r="B1"><v>20</v></c>',
      ),
    );
  });

  it('closes a self-closing target before a later non-empty cell', () => {
    const xml = worksheetXml('<c r="A1"/><c r="B1"><v>2</v></c>');
    const output = patchXlsxWorksheetPart(
      new TextEncoder().encode(xml),
      [
        requested(
          cell('A1', 1, {
            kind: 'value',
            value: { kind: 'number', value: 10 },
          }),
        ),
      ],
      limits,
      PART,
    );
    expect(new TextDecoder().decode(output)).toBe(
      xml.replace('<c r="A1"/>', '<c r="A1"><v>10</v></c>'),
    );
  });

  it('preserves qualified element names, allowed authored attributes, and BOM', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}">
        <s:sheetData><s:row r="1"><s:c  r='A1' s = "2" cm="3" ph="0" vm="4" t = 'n'><s:v>1</s:v></s:c></s:row></s:sheetData>
      </s:worksheet>`;
    const source = new Uint8Array([
      0xef,
      0xbb,
      0xbf,
      ...new TextEncoder().encode(xml),
    ]);
    const output = patchXlsxWorksheetPart(
      source,
      [
        requested(
          cell('A1', 1, {
            kind: 'value',
            value: { kind: 'boolean', value: true },
          }),
        ),
      ],
      limits,
      PART,
    );
    expect([...output.slice(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    const text = new TextDecoder().decode(output);
    expect(text).toContain(
      `<s:c  r='A1' s = "2" cm="3" ph="0" vm="4" t="b"><s:v>1</s:v></s:c>`,
    );
  });

  it.each([
    ['utf-16le', false],
    ['utf-16le', true],
    ['utf-16be', false],
    ['utf-16be', true],
  ] as const)('preserves %s encoding with BOM=%s', (encoding, bom) => {
    const xml = `<?xml version="1.0" encoding="${encoding}"?><worksheet xmlns="${XLSX_SPREADSHEET_NS}" note="é😀"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`;
    const source = encodeUtf16(xml, encoding, bom);
    const output = patchXlsxWorksheetPart(
      source,
      [requested(cell('A1', 1, { kind: 'blank' }))],
      limits,
      PART,
    );
    expect(output).toEqual(
      encodeUtf16(
        xml.replace('<c r="A1"><v>1</v></c>', '<c r="A1"/>'),
        encoding,
        bom,
      ),
    );
    const decoded = new TextDecoder(encoding).decode(output);
    expect(decoded).toContain('<c r="A1"/>');
    expect(decoded).not.toContain('<v>1</v>');
  });

  it('returns an isolated exact copy for an empty patch list', () => {
    const source = new TextEncoder().encode(worksheetXml('<c r="A1"/>'));
    const output = patchXlsxWorksheetPart(source, [], limits, PART);
    expect(output).toEqual(source);
    expect(output).not.toBe(source);
    const malformed = new Uint8Array([0xff]);
    expect(patchXlsxWorksheetPart(malformed, [], limits, PART)).toEqual(
      malformed,
    );
    expect(
      patchXlsxWorksheetPartWithReport(malformed, [], limits, PART),
    ).toEqual({ data: malformed, patchBytes: 0, patchCount: 0 });
  });

  it('skips comments and processing instructions outside patch tokens', () => {
    const xml = `<?xml version="1.0"?><worksheet xmlns="${XLSX_SPREADSHEET_NS}">
      <!-- before --><sheetData><?inside ok?><row r="1"><c r="A1"><v>1</v></c></row></sheetData><!-- after -->
    </worksheet>`;
    const output = patchXlsxWorksheetPart(
      new TextEncoder().encode(xml),
      [requested(cell('A1', 1, { kind: 'blank' }))],
      limits,
      PART,
    );
    const text = new TextDecoder().decode(output);
    expect(text).toContain('<!-- before -->');
    expect(text).toContain('<?inside ok?>');
    expect(text).toContain('<c r="A1"/>');
  });

  it('skips a complete CDATA section before worksheet tokens', () => {
    const xml = `<![CDATA[not markup]]>${worksheetXml('<c r="A1"><v>1</v></c>')}`;
    const output = patchXlsxWorksheetPart(
      new TextEncoder().encode(xml),
      [requested(cell('A1', 1, { kind: 'blank' }))],
      limits,
      PART,
    );
    expect(new TextDecoder().decode(output)).toBe(
      xml.replace('<c r="A1"><v>1</v></c>', '<c r="A1"/>'),
    );
  });

  it('tracks quoted greater-than text and anchors self-closing detection', () => {
    const xml = `<worksheet note=">"><sheetData><row r="1"><c cm='>' ph="/>" r="A1"><v>1</v></c></row></sheetData></worksheet>`;
    const output = patchXlsxWorksheetPart(
      new TextEncoder().encode(xml),
      [requested(cell('A1', 1, { kind: 'blank' }))],
      limits,
      PART,
    );
    expect(new TextDecoder().decode(output)).toContain(
      `<c cm='>' ph="/>" r="A1"/>`,
    );
  });

  it.each([
    ['utf-16le first BOM byte only', [0xff, 0x00], 'utf-16le'],
    ['utf-16le second BOM byte only', [0x00, 0xfe], 'utf-16le'],
    ['utf-16be first BOM byte only', [0xfe, 0x00], 'utf-16be'],
    ['utf-16be second BOM byte only', [0x00, 0xff], 'utf-16be'],
  ] as const)('rejects ambiguous %s', (_name, prefix, encoding) => {
    const xml = worksheetXml('<c r="A1"/>');
    const body = encodeUtf16(xml, encoding, false);
    const source = new Uint8Array(prefix.length + body.byteLength);
    source.set(prefix);
    source.set(body, prefix.length);
    expect(
      capture(() =>
        patchXlsxWorksheetPart(
          source,
          [requested(cell('A1', 1, { kind: 'blank' }))],
          limits,
          PART,
        ),
      ).diagnostic.code,
    ).toBe('preservation-conflict');
  });

  it.each([
    ['UTF-16BE marker second byte only', [0x41, 0x3c], 'utf-16be'],
    ['UTF-8 BOM first two bytes only', [0xef, 0xbb, 0x00], 'utf-8'],
    ['UTF-8 BOM last two bytes only', [0x00, 0xbb, 0xbf], 'utf-8'],
    ['UTF-8 BOM outer bytes only', [0xef, 0x00, 0xbf], 'utf-8'],
  ] as const)('rejects partial signature %s', (_name, prefix, encoding) => {
    const xml = worksheetXml('<c r="A1"/>');
    const body =
      encoding === 'utf-8'
        ? new TextEncoder().encode(xml)
        : encodeUtf16(xml, encoding, false);
    const source = new Uint8Array(prefix.length + body.byteLength);
    source.set(prefix);
    source.set(body, prefix.length);
    expect(
      capture(() =>
        patchXlsxWorksheetPart(
          source,
          [requested(cell('A1', 1, { kind: 'blank' }))],
          limits,
          PART,
        ),
      ).diagnostic.code,
    ).toBe('preservation-conflict');
  });

  it.each([
    [new Uint8Array([0xc3, 0x28]), 'XLSX worksheet XML encoding is invalid'],
    [
      new TextEncoder().encode('<worksheet'),
      'XLSX worksheet element or attribute name is unterminated',
    ],
    [
      new TextEncoder().encode('<worksheet '),
      'XLSX worksheet XML token is unterminated',
    ],
    [
      new TextEncoder().encode('<worksheet bad="unterminated >'),
      'XLSX worksheet XML token is unterminated',
    ],
    [
      new TextEncoder().encode('<!--'),
      'XLSX worksheet comment is unterminated',
    ],
    [
      new TextEncoder().encode('<![CDATA['),
      'XLSX worksheet CDATA is unterminated',
    ],
    [
      new TextEncoder().encode('<?bad'),
      'XLSX worksheet processing instruction is unterminated',
    ],
    [
      new TextEncoder().encode('<!DOCTYPE worksheet>'),
      'XLSX worksheet declaration cannot be patched safely',
    ],
    [
      new TextEncoder().encode('<1bad/>'),
      'XLSX worksheet element name is invalid',
    ],
    [
      new TextEncoder().encode('<worksheet bad></worksheet>'),
      'XLSX worksheet attribute assignment is invalid',
    ],
    [
      new TextEncoder().encode('<worksheet ="x"></worksheet>'),
      'XLSX worksheet attribute name is invalid',
    ],
    [
      new TextEncoder().encode('<worksheet bad=x></worksheet>'),
      'XLSX worksheet attribute quote is invalid',
    ],
    [
      new TextEncoder().encode('<worksheet><row></worksheet>'),
      'XLSX worksheet element nesting is invalid',
    ],
    [
      new TextEncoder().encode('<worksheet>'),
      'XLSX worksheet element is unclosed',
    ],
    [
      new TextEncoder().encode('<notWorksheet/>'),
      'XLSX worksheet root cannot be patched safely',
    ],
    [
      new TextEncoder().encode('plain text'),
      'XLSX worksheet root cannot be patched safely',
    ],
  ] as const)('rejects unsafe XML token source %#', (source, message) => {
    expect(
      capture(() =>
        patchXlsxWorksheetPart(
          source,
          [requested(cell('A1', 1, { kind: 'blank' }))],
          limits,
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({ code: 'preservation-conflict', message, part: PART });
  });

  it.each([
    [
      '<c><v>1</v></c>',
      'missing-cell-span',
      'XLSX target cell has no safe explicit XML span',
    ],
    [
      '<c r="A&amp;1"><v>1</v></c>',
      'missing-cell-span',
      'XLSX target cell has no safe explicit XML span',
    ],
    [
      '<c r="A1"/><c r="A1"/>',
      undefined,
      'XLSX target cell reference is ambiguous',
    ],
    [
      '<c r="A1" custom="x"><v>1</v></c>',
      'cell-extension',
      'XLSX target cell contains an unsupported attribute',
    ],
    [
      '<c r="A1"><extLst/></c>',
      'cell-extension',
      'XLSX target cell contains unsupported child content',
    ],
    [
      '<c r="A1"><v custom="x">1</v></c>',
      'cell-extension',
      'XLSX target cell child contains unsupported attributes',
    ],
  ] as const)(
    'blocks unsafe target shape %#',
    (sourceCell, featureClass, message) => {
      const error = capture(() =>
        patchXlsxWorksheetPart(
          new TextEncoder().encode(worksheetXml(sourceCell)),
          [requested(cell('A1', 1, { kind: 'blank' }))],
          limits,
          PART,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        cell: 'A1',
        code: 'preservation-conflict',
        message,
        operationId: 'edit-1',
        part: PART,
      });
      if (featureClass === undefined) {
        expect(error.diagnostic).not.toHaveProperty('featureClass');
      } else {
        expect(error.diagnostic.featureClass).toBe(featureClass);
      }
    },
  );

  it.each([
    '<f t="shared" si="0" ref="A1:A2">A2</f>',
    '<f si="0">A2</f>',
    '<f ref="A1:A2">A2</f>',
    '<f t="normal">A2</f>',
    '<f ca="1" si="0">A2</f>',
  ])('blocks grouped formula patch point %s', (formula) => {
    const error = capture(() =>
      patchXlsxWorksheetPart(
        new TextEncoder().encode(
          worksheetXml(`<c r="A1">${formula}<v>1</v></c>`),
        ),
        [
          requested(
            cell('A1', 1, {
              cached: { kind: 'missing' },
              formula: { expression: '1+1', kind: 'normal' },
              kind: 'formula',
            }),
          ),
        ],
        limits,
        PART,
      ),
    );
    expect(error.diagnostic).toMatchObject({
      code: 'formula-rewrite-unsupported',
      featureClass: 'formula-group',
      message: 'XLSX grouped formula cells cannot be patched independently',
    });
  });

  it('allows a normal formula flag that does not create a formula group', () => {
    const xml = worksheetXml('<c r="A1"><f ca="1">1+1</f></c>');
    const output = patchXlsxWorksheetPart(
      new TextEncoder().encode(xml),
      [
        requested(
          cell('A1', 1, {
            cached: { kind: 'missing' },
            formula: { expression: '2+2', kind: 'normal' },
            kind: 'formula',
          }),
        ),
      ],
      limits,
      PART,
    );
    expect(new TextDecoder().decode(output)).toContain(
      '<c r="A1"><f>2+2</f></c>',
    );
  });

  it('does not inspect grouped-formula children owned by another cell', () => {
    const xml = worksheetXml(
      '<c r="A1"><f t="shared" si="0" ref="A1:A2">1+1</f></c><c r="B1"><v>2</v></c>',
    );
    const output = patchXlsxWorksheetPart(
      new TextEncoder().encode(xml),
      [
        requested(
          cell('B1', 2, {
            kind: 'value',
            value: { kind: 'number', value: 20 },
          }),
        ),
      ],
      limits,
      PART,
    );
    expect(new TextDecoder().decode(output)).toBe(
      xml.replace('<c r="B1"><v>2</v></c>', '<c r="B1"><v>20</v></c>'),
    );
  });

  it('does not inspect grouped-formula children in a later cell', () => {
    const xml = worksheetXml(
      '<c r="A1"><v>1</v></c><c r="B1"><f t="shared" si="0" ref="B1:B2">1+1</f></c>',
    );
    const output = patchXlsxWorksheetPart(
      new TextEncoder().encode(xml),
      [
        requested(
          cell('A1', 1, {
            kind: 'value',
            value: { kind: 'number', value: 10 },
          }),
        ),
      ],
      limits,
      PART,
    );
    expect(new TextDecoder().decode(output)).toBe(
      xml.replace('<c r="A1"><v>1</v></c>', '<c r="A1"><v>10</v></c>'),
    );
  });

  it('rejects duplicate requested cells atomically', () => {
    const patch = requested(cell('A1', 1, { kind: 'blank' }));
    const error = capture(() =>
      patchXlsxWorksheetPart(
        new TextEncoder().encode(worksheetXml('<c r="A1"/>')),
        [patch, { ...patch, operationId: 'edit-2' }],
        limits,
        PART,
      ),
    );
    expect(error.diagnostic).toMatchObject({
      code: 'preservation-conflict',
      message: 'XLSX worksheet patch cells must be unique',
      operationId: 'edit-2',
    });
  });

  it('rejects date output and invalid XML characters with structured errors', () => {
    const date = requested(
      cell('A1', 1, {
        kind: 'value',
        value: {
          kind: 'date',
          normalized: '2024-01-01',
          precision: 'date',
          source: { kind: 'iso', value: '2024-01-01' },
        },
      }),
    );
    expect(
      capture(() =>
        patchXlsxWorksheetPart(
          new TextEncoder().encode(worksheetXml('<c r="A1"/>')),
          [date],
          limits,
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({
      code: 'unsupported-edit-operation',
      featureClass: 'date-value',
      message: 'XLSX worksheet patch does not support date values',
    });
    const invalid = requested(
      cell('A1', 1, {
        kind: 'value',
        value: { kind: 'text', text: 'bad\u0000text' },
      }),
    );
    expect(() => escapeXlsxCellText('bad\u0000text', invalid)).toThrow(
      'XLSX cell content contains an invalid XML character',
    );
  });

  it.each([
    ['\u0009', '\u0009'],
    ['\u000a', '\u000a'],
    ['\u000d', '&#13;'],
    ['\u0020', ' '],
    ['\ud7ff', '\ud7ff'],
    ['\ue000', '\ue000'],
    ['\ufffd', '\ufffd'],
    ['\u{10000}', '\u{10000}'],
    ['\u{10ffff}', '\u{10ffff}'],
  ])('escapes valid XML boundary %s', (value, expected) => {
    const patch = requested(
      cell('A1', 1, { kind: 'value', value: { kind: 'text', text: value } }),
    );
    expect(escapeXlsxCellText(value, patch)).toBe(expected);
  });

  it('escapes every XML text delimiter including the CDATA terminator', () => {
    const patch = requested(
      cell('A1', 1, { kind: 'value', value: { kind: 'text', text: '' } }),
    );
    expect(escapeXlsxCellText('&<>]]>\r', patch)).toBe(
      '&amp;&lt;&gt;]]&gt;&#13;',
    );
  });

  it.each([
    '\u0000',
    '\u0008',
    '\u000b',
    '\u000c',
    '\u001f',
    '\ud800',
    '\udfff',
    '\ufffe',
    '\uffff',
  ])('rejects invalid XML boundary %s', (value) => {
    const patch = requested(
      cell('A1', 1, { kind: 'value', value: { kind: 'text', text: value } }),
    );
    expect(
      capture(() => escapeXlsxCellText(value, patch)).diagnostic,
    ).toMatchObject({
      cell: 'A1',
      code: 'invalid-roundtrip-json',
      message: 'XLSX cell content contains an invalid XML character',
      operationId: 'edit-1',
    });
  });

  it('accepts authored inline strings as a replaceable known child', () => {
    const source = new TextEncoder().encode(
      worksheetXml('<c r="A1" t="inlineStr"><is><t>old</t></is></c>'),
    );
    const output = patchXlsxWorksheetPart(
      source,
      [
        requested(
          cell('A1', 1, {
            kind: 'value',
            value: { kind: 'text', text: 'new' },
          }),
        ),
      ],
      limits,
      PART,
    );
    expect(new TextDecoder().decode(output)).toContain(
      '<c r="A1" t="inlineStr"><is><t xml:space="preserve">new</t></is></c>',
    );
  });

  it('enforces patch count, patch bytes, and generated XML bytes exactly', () => {
    const source = new TextEncoder().encode(worksheetXml('<c r="A1"/>'));
    const patch = requested(cell('A1', 1, { kind: 'blank' }));
    const replacementBytes = new TextEncoder().encode('<c r="A1"/>').byteLength;
    expect(
      patchXlsxWorksheetPartWithReport(source, [patch], limits, PART),
    ).toMatchObject({
      patchBytes: replacementBytes,
      patchCount: 1,
    });
    expect(() =>
      patchXlsxWorksheetPart(
        source,
        [patch],
        {
          ...limits,
          maxGeneratedXmlBytes: source.byteLength,
          maxPatchBytes: replacementBytes,
          maxPatchCount: 1,
        },
        PART,
      ),
    ).not.toThrow();
    for (const [limitName, overrides] of [
      ['maxPatchCount', { maxPatchCount: 1 }],
      ['maxPatchBytes', { maxPatchBytes: replacementBytes - 1 }],
      ['maxGeneratedXmlBytes', { maxGeneratedXmlBytes: source.byteLength - 1 }],
    ] as const) {
      const error = capture(() =>
        patchXlsxWorksheetPart(
          source,
          limitName === 'maxPatchCount'
            ? [patch, { ...patch, cell: { ...patch.cell, address: 'B1' } }]
            : [patch],
          { ...limits, ...overrides },
          PART,
        ),
      );
      expect(error.diagnostic.limitName).toBe(limitName);
    }
  });
});
