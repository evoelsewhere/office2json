import { describe, expect, it } from 'vitest';

import { parseXlsx } from '../../src/formats/xlsx/parser';
import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import {
  patchXlsxInternalHyperlinks,
  readXlsxHyperlinkRelationshipIds,
} from '../../src/formats/xlsx/roundtrip/hyperlink-patch';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import {
  createIndependentXlsx,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

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
  throw new Error('Expected hyperlink patch to fail');
}

function worksheet(hyperlinks = ''): string {
  return `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="B2"><v>2</v></c></row></sheetData>${hyperlinks}</worksheet>`;
}

describe('XLSX internal hyperlink patching', () => {
  it('reads only direct relationship IDs from the worksheet hyperlink collection', () => {
    const source = bytes(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="urn:relationships" xmlns:x="urn:foreign"><wrapper><hyperlink ref="BEFORE" r:id="before"/></wrapper><hyperlinks><other ref="OTHER" r:id="other"/><wrapper><hyperlink ref="NESTED" r:id="nested"/></wrapper><x:hyperlink ref="FOREIGN" r:id="foreign"/><hyperlink bogus="wrong" ref="REAL" r:id="real"/></hyperlinks><wrapper><hyperlink ref="AFTER" r:id="after"/></wrapper></worksheet>`,
    );
    expect(readXlsxHyperlinkRelationshipIds(source, PART)).toEqual(
      new Map([['REAL', 'real']]),
    );

    const prefixed = bytes(
      `<s:worksheet xmlns:s="${XLSX_SPREADSHEET_NS}" xmlns:r="urn:relationships"><s:hyperlinks><s:hyperlink ref="A1" r:id="rId1"/></s:hyperlinks></s:worksheet>`,
    );
    expect(readXlsxHyperlinkRelationshipIds(prefixed, PART)).toEqual(
      new Map([['A1', 'rId1']]),
    );
  });

  it('requires a depth-zero worksheet root when reading relationship IDs', () => {
    for (const source of [
      '<outer/>',
      `<outer><worksheet xmlns="${XLSX_SPREADSHEET_NS}"><hyperlinks/></worksheet></outer>`,
      '<notWorksheet><hyperlinks/></notWorksheet>',
    ]) {
      expect(
        capture(() => readXlsxHyperlinkRelationshipIds(bytes(source), PART))
          .diagnostic,
      ).toMatchObject({
        message: 'XLSX worksheet root cannot read hyperlinks',
        part: PART,
      });
    }
  });

  it('updates authored metadata and appends a new internal hyperlink', async () => {
    const result = patchXlsxInternalHyperlinks(
      bytes(
        worksheet(
          '<hyperlinks><hyperlink ref="A1" location="Old!A1" display="Display" tooltip="Tip"/></hyperlinks>',
        ),
      ),
      [
        {
          cell: 'A1',
          operationId: 'update',
          target: { kind: 'internal', location: "'A & B'!C3" },
        },
        {
          cell: 'B2',
          operationId: 'add',
          target: { kind: 'internal', location: 'Sheet2!A1' },
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(result.patchCount).toBe(2);
    expect(result.patchBytes).toBe(121);
    const xml = new TextDecoder().decode(result.data);
    expect(xml).toContain(
      '<hyperlink ref="A1" display="Display" tooltip="Tip" location="&apos;A &amp; B&apos;!C3"/>'.replaceAll(
        '&apos;',
        "'",
      ),
    );
    const parsed = await parseXlsx(
      await createIndependentXlsx({ [PART]: result.data }),
      { errorMode: 'strict' },
    );
    const sheet = parsed.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    if (sheet.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.hyperlinks.map((link) => link.target)).toEqual([
      { kind: 'internal', location: "'A & B'!C3" },
      { kind: 'internal', location: 'Sheet2!A1' },
    ]);
  });

  it('adds a missing collection and removes the final collection exactly', () => {
    const added = patchXlsxInternalHyperlinks(
      bytes(worksheet()),
      [
        {
          cell: 'A1',
          operationId: 'add',
          target: { kind: 'internal', location: 'Sheet2!A1' },
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(added.data)).toContain(
      '<hyperlinks><hyperlink ref="A1" location="Sheet2!A1"/></hyperlinks></worksheet>',
    );
    const removed = patchXlsxInternalHyperlinks(
      added.data,
      [{ cell: 'A1', operationId: 'remove', target: null }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(removed.data)).not.toContain('hyperlinks');
    const removeOne = patchXlsxInternalHyperlinks(
      bytes(
        worksheet(
          '<hyperlinks><hyperlink ref="A1" location="First!A1"/><hyperlink ref="B2" location="Second!A1"/></hyperlinks>',
        ),
      ),
      [{ cell: 'A1', operationId: 'remove-one', target: null }],
      defaultXlsxWriteLimits(),
      PART,
    );
    const removeOneXml = new TextDecoder().decode(removeOne.data);
    expect(removeOneXml).not.toContain('First!A1');
    expect(removeOneXml).not.toContain('Stryker was here!');
    expect(removeOneXml).toContain('Second!A1');
  });

  it('keeps update-only and absent-removal cardinality exact', () => {
    const source = bytes(
      worksheet(
        '<hyperlinks><hyperlink display="D" ref="A1" location="Old!A1"/></hyperlinks>',
      ),
    );
    const updated = patchXlsxInternalHyperlinks(
      source,
      [
        {
          cell: 'A1',
          operationId: 'update',
          target: { kind: 'internal', location: 'New!A1' },
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(updated.patchCount).toBe(1);
    expect(new TextDecoder().decode(updated.data)).toContain(
      '<hyperlink display="D" ref="A1" location="New!A1"/>',
    );
    const absent = patchXlsxInternalHyperlinks(
      source,
      [{ cell: 'B2', operationId: 'absent', target: null }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(absent.patchCount).toBe(0);
    expect(absent.data).toEqual(source);
    const missingCollection = bytes(worksheet());
    const missing = patchXlsxInternalHyperlinks(
      missingCollection,
      [{ cell: 'A1', operationId: 'absent', target: null }],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(missing).toEqual({
      data: missingCollection,
      patchBytes: 0,
      patchCount: 0,
    });
    const empty = patchXlsxInternalHyperlinks(
      source,
      [],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(empty).toEqual({ data: source, patchBytes: 0, patchCount: 0 });
    expect(empty.data).not.toBe(source);
  });

  it('appends multiple links without injecting a separator', () => {
    const result = patchXlsxInternalHyperlinks(
      bytes(worksheet()),
      [
        {
          cell: 'A1',
          operationId: 'first',
          target: { kind: 'internal', location: 'First!A1' },
        },
        {
          cell: 'B2',
          operationId: 'second',
          target: { kind: 'internal', location: 'Second!A1' },
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toContain(
      'location="First!A1"/><hyperlink ref="B2" location="Second!A1"/>',
    );
    const existing = patchXlsxInternalHyperlinks(
      bytes(
        worksheet(
          '<hyperlinks><hyperlink ref="C3" location="Existing!A1"/></hyperlinks>',
        ),
      ),
      [
        {
          cell: 'A1',
          operationId: 'first',
          target: { kind: 'internal', location: 'First!A1' },
        },
        {
          cell: 'B2',
          operationId: 'second',
          target: { kind: 'internal', location: 'Second!A1' },
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(existing.data)).toContain(
      'location="First!A1"/><hyperlink ref="B2" location="Second!A1"/></hyperlinks>',
    );
  });

  it('uses direct collection and hyperlink depth despite lexical lookalikes', () => {
    const source = worksheet(
      '<x:hyperlinks xmlns:x="urn:foreign"><x:hyperlink ref="A1" location="Foreign!A1"/></x:hyperlinks><foreign><hyperlinks><hyperlink ref="A1" location="Fake!A1"/></hyperlinks></foreign><hyperlinks><other ref="A1"/><wrapper><hyperlink ref="A1" location="Nested!A1"/></wrapper><hyperlink><hyperlink ref="B2" location="NestedSame!A1"/></hyperlink><hyperlink display="Real" ref="A1" location="Old!A1"></hyperlink></hyperlinks>',
    );
    const result = patchXlsxInternalHyperlinks(
      bytes(source),
      [
        {
          cell: 'A1',
          operationId: 'real',
          target: { kind: 'internal', location: 'New!A1' },
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const xml = new TextDecoder().decode(result.data);
    expect(xml).toContain('location="Fake!A1"');
    expect(xml).toContain('location="Foreign!A1"');
    expect(xml).toContain('location="Nested!A1"');
    expect(xml).toContain('location="NestedSame!A1"');
    expect(xml).toContain(
      '<hyperlink display="Real" ref="A1" location="New!A1"/>',
    );

    const nestedCollections = worksheet(
      '<hyperlinks><hyperlinks><hyperlink ref="A1" location="NestedCollection!A1"/></hyperlinks><hyperlink ref="A1" location="Real!A1"/></hyperlinks>',
    );
    const nestedResult = patchXlsxInternalHyperlinks(
      bytes(nestedCollections),
      [
        {
          cell: 'A1',
          operationId: 'outer-collection',
          target: { kind: 'internal', location: 'Updated!A1' },
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(nestedResult.data)).toContain(
      '<hyperlink ref="A1" location="Updated!A1"/></hyperlinks>',
    );

    const nestedHyperlink = worksheet(
      '<hyperlinks><hyperlink ref="A1" location="Outer!A1"><hyperlink ref="B2" location="Inner!A1"></hyperlink></hyperlink></hyperlinks>',
    );
    const nestedHyperlinkResult = patchXlsxInternalHyperlinks(
      bytes(nestedHyperlink),
      [
        {
          cell: 'A1',
          operationId: 'outer-link',
          target: { kind: 'internal', location: 'Updated!A1' },
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const nestedHyperlinkXml = new TextDecoder().decode(
      nestedHyperlinkResult.data,
    );
    expect(nestedHyperlinkXml).not.toContain('Inner!A1');
    expect(nestedHyperlinkXml).toContain('location="Updated!A1"/>');

    const priorSibling = worksheet(
      '<hyperlinks><hyperlink ref="A1" location="First!A1"></hyperlink><hyperlink ref="B2" location="Second!A1"></hyperlink></hyperlinks>',
    );
    const priorSiblingResult = patchXlsxInternalHyperlinks(
      bytes(priorSibling),
      [
        {
          cell: 'B2',
          operationId: 'second-link',
          target: { kind: 'internal', location: 'UpdatedSecond!A1' },
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    const siblingXml = new TextDecoder().decode(priorSiblingResult.data);
    expect(siblingXml).toContain('location="First!A1"></hyperlink>');
    expect(siblingXml).toContain('ref="B2" location="UpdatedSecond!A1"/>');
  });

  it('preserves prefixed UTF-16LE worksheet XML', () => {
    const source = worksheet()
      .replaceAll('<worksheet', '<s:worksheet')
      .replaceAll('</worksheet>', '</s:worksheet>')
      .replaceAll(/<(\/)?(sheetData|row|c|v)(?=[\s/>])/gu, '<$1s:$2')
      .replace('xmlns=', 'xmlns:s=');
    const encoded = new Uint8Array(2 + source.length * 2);
    encoded.set([0xff, 0xfe]);
    for (let index = 0; index < source.length; index += 1) {
      const code = source.charCodeAt(index);
      encoded[2 + index * 2] = code & 0xff;
      encoded[3 + index * 2] = code >>> 8;
    }
    const result = patchXlsxInternalHyperlinks(
      encoded,
      [
        {
          cell: 'A1',
          operationId: 'add',
          target: { kind: 'internal', location: 'Sheet2!A1' },
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect([...result.data.slice(0, 2)]).toEqual([0xff, 0xfe]);
    expect(new TextDecoder('utf-16le').decode(result.data.slice(2))).toContain(
      '<s:hyperlinks><s:hyperlink',
    );
  });

  it('escapes XML attribute characters and validates character boundaries', () => {
    const location = `&<"\r\n\t\ud7ff\ue000\ufffd${String.fromCodePoint(0x1_0000)}😀`;
    const result = patchXlsxInternalHyperlinks(
      bytes(worksheet()),
      [
        {
          cell: 'A1',
          operationId: 'escaped',
          target: { kind: 'internal', location },
        },
      ],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toContain(
      `location="&amp;&lt;&quot;&#13;&#10;&#9;\ud7ff\ue000\ufffd${String.fromCodePoint(0x1_0000)}😀"`,
    );
    for (const value of ['bad\u0001', 'bad\ud800', 'bad\ufffe']) {
      expect(
        capture(() =>
          patchXlsxInternalHyperlinks(
            bytes(worksheet()),
            [
              {
                cell: 'A1',
                operationId: 'bad-text',
                target: { kind: 'internal', location: value },
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({
        code: 'invalid-roundtrip-json',
        message: 'XLSX hyperlink location contains an invalid XML character',
        operationId: 'bad-text',
      });
    }
  });

  it.each([
    [
      '<hyperlinks><hyperlink ref="A1" r:id="link"/></hyperlinks>',
      'XLSX internal hyperlink edit cannot replace an external relationship target',
      'external-hyperlink',
    ],
    [
      '<hyperlinks><hyperlink ref="A1" custom="x"/></hyperlinks>',
      'XLSX hyperlink element contains an unsupported attribute',
      'hyperlink-xml',
    ],
    [
      '<hyperlinks><hyperlink ref="A1"/><hyperlink ref="A1"/></hyperlinks>',
      'XLSX hyperlink reference is ambiguous',
      'hyperlink-xml',
    ],
  ] as const)(
    'rejects unsafe authored hyperlink %#',
    (source, message, featureClass) => {
      expect(
        capture(() =>
          patchXlsxInternalHyperlinks(
            bytes(worksheet(source)),
            [
              {
                cell: 'A1',
                operationId: 'edit',
                target: { kind: 'internal', location: 'New!A1' },
              },
            ],
            defaultXlsxWriteLimits(),
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({
        featureClass,
        message,
        operationId: 'edit',
        part: PART,
      });
    },
  );

  it('rejects duplicate requests and invalid XML characters', () => {
    expect(
      capture(() =>
        patchXlsxInternalHyperlinks(
          bytes(`<outer>${worksheet()}</outer>`),
          [{ cell: 'A1', operationId: 'root', target: null }],
          defaultXlsxWriteLimits(),
          PART,
        ),
      ).diagnostic.message,
    ).toBe('XLSX worksheet root cannot patch hyperlinks');
    expect(
      capture(() =>
        patchXlsxInternalHyperlinks(
          bytes(worksheet()),
          [
            { cell: 'A1', operationId: 'one', target: null },
            { cell: 'A1', operationId: 'two', target: null },
          ],
          defaultXlsxWriteLimits(),
          PART,
        ),
      ).diagnostic.message,
    ).toBe('XLSX hyperlink patch cells must be unique');
    expect(
      capture(() =>
        patchXlsxInternalHyperlinks(
          bytes(worksheet()),
          [
            {
              cell: 'A1',
              operationId: 'bad-text',
              target: { kind: 'internal', location: 'bad\u0000' },
            },
          ],
          defaultXlsxWriteLimits(),
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({
      code: 'invalid-roundtrip-json',
      operationId: 'bad-text',
      part: PART,
    });
  });

  it('enforces generated XML and patch bytes at exact boundaries', () => {
    const source = bytes(worksheet());
    const request = [
      {
        cell: 'A1',
        operationId: 'add',
        target: { kind: 'internal' as const, location: 'Sheet2!A1' },
      },
    ];
    const successful = patchXlsxInternalHyperlinks(
      source,
      request,
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(() =>
      patchXlsxInternalHyperlinks(
        source,
        request,
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
          patchXlsxInternalHyperlinks(
            source,
            request,
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ limit, limitName, part: PART });
    }
  });
});
