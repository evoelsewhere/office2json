import { describe, expect, it } from 'vitest';

import {
  patchXlsxCommentAnchors,
  patchXlsxCommentVmlAnchors,
} from '../../src/formats/xlsx/roundtrip/comment-structure-patch';
import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import { XLSX_SPREADSHEET_NS } from '../black-box/xlsx-package';

const PART = 'xl/comments1.xml';
const VML_PART = 'xl/drawings/vmlDrawing1.vml';

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
  throw new Error('Expected comment structural patch to fail');
}

const OPERATIONS = [
  { count: 1, index: 2, kind: 'insert-rows' as const, operationId: 'rows' },
  {
    count: 1,
    index: 2,
    kind: 'insert-columns' as const,
    operationId: 'columns',
  },
];

describe('XLSX comment structural patching', () => {
  it('transforms legacy and threaded comment references', () => {
    const legacy = `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><t>One</t></text></comment><comment authorId="0" ref="B2"><text><t>Two</t></text></comment></commentList></comments>`;
    const legacyResult = patchXlsxCommentAnchors(
      bytes(legacy),
      OPERATIONS,
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(legacyResult.data)).toContain(
      '<comment authorId="0" ref="C3">',
    );
    expect(new TextDecoder().decode(legacyResult.data)).toContain(
      '<comment ref="A1" authorId="0">',
    );
    expect(legacyResult.patchCount).toBe(2);

    const threaded = `<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="C3" id="root"><text>Root</text></threadedComment><threadedComment ref="C3" id="reply" parentId="root"><text>Reply</text></threadedComment></ThreadedComments>`;
    const threadedResult = patchXlsxCommentAnchors(
      bytes(threaded),
      OPERATIONS,
      defaultXlsxWriteLimits(),
      'xl/threadedComments/threadedComment1.xml',
    );
    expect(new TextDecoder().decode(threadedResult.data)).toContain(
      '<threadedComment ref="D4" id="root">',
    );
    expect(new TextDecoder().decode(threadedResult.data)).toContain(
      '<threadedComment ref="D4" id="reply" parentId="root">',
    );
    expect(threadedResult.patchCount).toBe(4);
  });

  it('selects only owned prefixed legacy and threaded comments', () => {
    const legacy = `<s:comments xmlns:s="${XLSX_SPREADSHEET_NS}" xmlns:x="urn:foreign"><wrapper><s:commentList><s:comment ref="Z9"/></s:commentList></wrapper><x:commentList><x:comment ref="Z9"/></x:commentList><s:commentList><wrapper><s:comment ref="Z8"/></wrapper><x:comment ref="Z9"/><s:comment authorId="0" ref="A1"/></s:commentList><wrapper><s:comment ref="Z7"/></wrapper></s:comments>`;
    const legacyResult = patchXlsxCommentAnchors(
      bytes(legacy),
      [OPERATIONS[0]!],
      defaultXlsxWriteLimits(),
      PART,
    );
    const legacyOutput = new TextDecoder().decode(legacyResult.data);
    expect(legacyOutput).toContain('<s:comment authorId="0" ref="A1"/>');
    expect(legacyOutput).toContain('<s:comment ref="Z8"/>');
    expect(legacyOutput).toContain('<x:comment ref="Z9"/>');
    expect(legacyResult.patchCount).toBe(0);

    const threaded = `<s:ThreadedComments xmlns:s="urn:threaded" xmlns:x="urn:foreign"><wrapper><s:threadedComment ref="Z9"/></wrapper><x:threadedComment ref="Z9"/><s:threadedComment id="one" ref="A2"/><wrapper><s:threadedComment ref="Z8"/></wrapper></s:ThreadedComments>`;
    const threadedResult = patchXlsxCommentAnchors(
      bytes(threaded),
      [OPERATIONS[0]!],
      defaultXlsxWriteLimits(),
      'xl/threadedComments/threadedComment1.xml',
    );
    const threadedOutput = new TextDecoder().decode(threadedResult.data);
    expect(threadedOutput).toContain('<s:threadedComment id="one" ref="A3"/>');
    expect(threadedOutput).toContain('<s:threadedComment ref="Z8"/>');
    expect(threadedOutput).toContain('<x:threadedComment ref="Z9"/>');
    expect(threadedResult.patchCount).toBe(1);
  });

  it('transforms only owned VML note row and column anchors', () => {
    const vml = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><x:ClientData ObjectType="Note"><x:Row>6</x:Row><x:Column>6</x:Column></x:ClientData><other><x:ClientData ObjectType="Note"><x:Row>5</x:Row><x:Column>5</x:Column></x:ClientData></other><v:shape><x:Other ObjectType="Note"><x:Row>4</x:Row><x:Column>4</x:Column></x:Other><wrapper><x:ClientData ObjectType="Note"><x:Row>3</x:Row><x:Column>3</x:Column></x:ClientData></wrapper><x:ClientData ObjectType="Note"><wrapper><x:Row>7</x:Row><x:Column>7</x:Column></wrapper><x:Other>9</x:Other><x:Row> 1 </x:Row><x:Column>1</x:Column></x:ClientData></v:shape><v:shape><x:ClientData ObjectType="Button"><x:Row>8</x:Row><x:Column>8</x:Column></x:ClientData></v:shape><wrapper><v:shape><x:ClientData ObjectType="Note"><x:Row>7</x:Row><x:Column>7</x:Column></x:ClientData></v:shape></wrapper></xml>`;
    const result = patchXlsxCommentVmlAnchors(
      bytes(vml),
      OPERATIONS,
      defaultXlsxWriteLimits(),
      VML_PART,
    );
    const output = new TextDecoder().decode(result.data);
    expect(output).toContain('<x:Row>2</x:Row><x:Column>2</x:Column>');
    expect(output).toContain('<x:Row>8</x:Row><x:Column>8</x:Column>');
    expect(output).toContain('<x:Row>7</x:Row><x:Column>7</x:Column>');
    expect(output).toContain('<x:Row>6</x:Row><x:Column>6</x:Column>');
    expect(output).toContain('<x:Row>5</x:Row><x:Column>5</x:Column>');
    expect(output).toContain('<x:Row>4</x:Row><x:Column>4</x:Column>');
    expect(output).toContain('<x:Row>3</x:Row><x:Column>3</x:Column>');
    expect(result.patchCount).toBe(2);
  });

  it('fails closed when a structural edit deletes any comment anchor', () => {
    const operation = {
      count: 1,
      index: 2,
      kind: 'delete-rows' as const,
      operationId: 'delete-comment',
    };
    for (const [source, patch, part] of [
      [
        `<comments xmlns="${XLSX_SPREADSHEET_NS}"><commentList><comment ref="A2"/></commentList></comments>`,
        patchXlsxCommentAnchors,
        PART,
      ],
      [
        '<ThreadedComments><threadedComment ref="A2"/></ThreadedComments>',
        patchXlsxCommentAnchors,
        'xl/threadedComments/threadedComment1.xml',
      ],
      [
        '<xml><shape><ClientData ObjectType="Note"><Row>1</Row><Column>0</Column></ClientData></shape></xml>',
        patchXlsxCommentVmlAnchors,
        VML_PART,
      ],
    ] as const) {
      expect(
        capture(() =>
          patch(bytes(source), [operation], defaultXlsxWriteLimits(), part),
        ).diagnostic,
      ).toMatchObject({
        featureClass: 'comment-anchor-deletion',
        message: 'XLSX structural edit would delete a comment anchor',
        operationId: 'delete-comment',
        part,
        range: '2:2',
      });
    }
  });

  it('rejects malformed comment and VML structures exactly', () => {
    for (const [source, patch, message] of [
      [
        '<wrong/>',
        patchXlsxCommentAnchors,
        'XLSX comment root cannot patch structure',
      ],
      [
        '<!--empty-->',
        patchXlsxCommentAnchors,
        'XLSX comment root cannot patch structure',
      ],
      [
        `<comments xmlns="${XLSX_SPREADSHEET_NS}"/>`,
        patchXlsxCommentAnchors,
        'XLSX comment list cannot patch structure',
      ],
      [
        `<comments xmlns="${XLSX_SPREADSHEET_NS}"><commentList><comment ref="bad"/></commentList></comments>`,
        patchXlsxCommentAnchors,
        'XLSX structural comment reference is invalid',
      ],
      [
        '<xml><shape><ClientData ObjectType="Note"><Row>bad</Row><Column>0</Column></ClientData></shape></xml>',
        patchXlsxCommentVmlAnchors,
        'XLSX structural VML comment anchor is invalid',
      ],
      [
        '<xml><shape><ClientData ObjectType="Note"><Row>x1</Row><Column>0</Column></ClientData></shape></xml>',
        patchXlsxCommentVmlAnchors,
        'XLSX structural VML comment anchor is invalid',
      ],
      [
        '<xml><shape><ClientData ObjectType="Note"><Row>1x</Row><Column>0</Column></ClientData></shape></xml>',
        patchXlsxCommentVmlAnchors,
        'XLSX structural VML comment anchor is invalid',
      ],
      [
        '<xml><shape><ClientData ObjectType="Note"><Row>1.0</Row><Column>0</Column></ClientData></shape></xml>',
        patchXlsxCommentVmlAnchors,
        'XLSX structural VML comment anchor is invalid',
      ],
      [
        '<xml><shape><ClientData ObjectType="Note"><Row>0</Row></ClientData></shape></xml>',
        patchXlsxCommentVmlAnchors,
        'XLSX structural VML comment anchor is invalid',
      ],
    ] as const) {
      const error = capture(() =>
        patch(
          bytes(source),
          [OPERATIONS[0]!],
          defaultXlsxWriteLimits(),
          source.startsWith('<xml') ? VML_PART : PART,
        ),
      );
      expect(error.diagnostic).toMatchObject({
        featureClass: 'comment-structure-xml',
        message,
      });
    }
  });

  it('preserves exact no-op bytes and enforces resource limits', () => {
    const source = `<comments xmlns="${XLSX_SPREADSHEET_NS}"><commentList><comment ref="A1"/></commentList></comments>`;
    const request = {
      count: 1,
      index: 2,
      kind: 'insert-rows' as const,
      operationId: 'limits',
    };
    const noOp = patchXlsxCommentAnchors(
      bytes(source),
      [request],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(new TextDecoder().decode(noOp.data)).toBe(source);
    expect(noOp.patchCount).toBe(0);
    const input = bytes(source);
    const empty = patchXlsxCommentAnchors(
      input,
      [],
      defaultXlsxWriteLimits(),
      PART,
    );
    expect(empty.data).toEqual(input);
    expect(empty.data).not.toBe(input);
    const variable = `<comments xmlns="${XLSX_SPREADSHEET_NS}"><commentList><comment ref="A9"/><comment ref="B99"/></commentList></comments>`;
    expect(
      new TextDecoder().decode(
        patchXlsxCommentAnchors(
          bytes(variable),
          [
            {
              count: 1,
              index: 1,
              kind: 'insert-rows',
              operationId: 'variable-length',
            },
          ],
          defaultXlsxWriteLimits(),
          PART,
        ).data,
      ),
    ).toBe(
      `<comments xmlns="${XLSX_SPREADSHEET_NS}"><commentList><comment ref="A10"/><comment ref="B100"/></commentList></comments>`,
    );
    const changed = patchXlsxCommentAnchors(
      bytes(source.replace('A1', 'A2')),
      [request],
      defaultXlsxWriteLimits(),
      PART,
    );
    const replacementBytes = new TextEncoder().encode(' ref="A3"').byteLength;
    expect(changed.patchBytes).toBe(replacementBytes);
    expect(changed.patchCount).toBe(1);
    expect(() =>
      patchXlsxCommentAnchors(
        bytes(source.replace('A1', 'A2')),
        [request],
        {
          ...defaultXlsxWriteLimits(),
          maxGeneratedXmlBytes: changed.data.byteLength,
          maxPatchBytes: changed.patchBytes,
          maxPatchCount: changed.patchCount,
        },
        PART,
      ),
    ).not.toThrow();
    for (const [limitName, limit] of [
      ['maxGeneratedXmlBytes', changed.data.byteLength - 1],
      ['maxPatchBytes', changed.patchBytes - 1],
      ['maxPatchCount', 0],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxCommentAnchors(
            bytes(source.replace('A1', 'A2')),
            [request],
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ limitName });
    }
  });
});
