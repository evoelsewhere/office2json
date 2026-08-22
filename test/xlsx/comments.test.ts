import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  parseXlsxWithDiagnostics,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
  XlsxParseError,
} from '../../src/formats/xlsx';
import type { XmlLookupValue } from '../../src/common/xml/tree';
import {
  isValidXlsxCommentTimestamp,
  parseXlsxCommentVmlVisibility,
} from '../../src/formats/xlsx/internal/comments';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import {
  createIndependentXlsx,
  type XlsxBlackBoxOverrides,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const COMMENTS_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}comments`;
const VML_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}vmlDrawing`;
const THREADED_RELATIONSHIP =
  'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
const PERSON_RELATIONSHIP =
  'http://schemas.microsoft.com/office/2017/10/relationships/person';
const THREADED_NAMESPACE =
  'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';

const CONTENT_TYPES = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  <Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>
  <Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>
</Types>`;

const WORKBOOK_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
  <Relationship Id="rIdSheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdStyles" Type="${XLSX_OFFICE_REL_TYPE}styles" Target="styles.xml"/>
  <Relationship Id="rIdSharedStrings" Type="${XLSX_OFFICE_REL_TYPE}sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rIdPersons" Type="${PERSON_RELATIONSHIP}" Target="persons/person.xml"/>
</Relationships>`;

const WORKSHEET_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
  <Relationship Id="rIdComments" Type="${COMMENTS_RELATIONSHIP}" Target="../comments1.xml"/>
  <Relationship Id="rIdVml" Type="${VML_RELATIONSHIP}" Target="../drawings/vmlDrawing1.vml"/>
  <Relationship Id="rIdThreaded" Type="${THREADED_RELATIONSHIP}" Target="../threadedComments/threadedComment1.xml"/>
</Relationships>`;

const WORKSHEET = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_TYPE.slice(0, -1)}">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <legacyDrawing r:id="rIdVml"/>
</worksheet>`;

const LEGACY_COMMENTS = `<comments xmlns="${XLSX_SPREADSHEET_NS}">
  <authors><author>Alice &amp; Co</author><author>Bob</author></authors>
  <commentList>
    <comment ref="A1" authorId="0"><text><r><t xml:space="preserve">Hello </t></r><r><t>&amp; world</t></r></text></comment>
    <comment ref="B2" authorId="1"><text><t>Hidden note</t></text></comment>
  </commentList>
</comments>`;

const PERSONS = `<personList xmlns="${THREADED_NAMESPACE}">
  <person displayName="Carol" id="person-1" userId="carol@example.com" providerId="AD"/>
  <person displayName="Dan" id="person-2"/>
</personList>`;

const THREADED_COMMENTS = `<ThreadedComments xmlns="${THREADED_NAMESPACE}">
  <threadedComment ref="C3" dT="2024-01-02T03:04:05Z" personId="person-1" id="thread-1"><text>Root thread</text></threadedComment>
  <threadedComment ref="C3" dT="2024-01-02T04:05:06+01:00" personId="person-2" id="thread-2" parentId="thread-1"><text>Reply</text></threadedComment>
</ThreadedComments>`;

const VML = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape style="visibility:hidden"><x:ClientData ObjectType="Note"><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>
  <v:shape style="visibility:visible"><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>
  <v:shape style="visibility:hidden"><x:ClientData ObjectType="Note"><x:Row>2</x:Row><x:Column>2</x:Column><x:Visible/></x:ClientData></v:shape>
</xml>`;

function parts(overrides: XlsxBlackBoxOverrides = {}): XlsxBlackBoxOverrides {
  return {
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
    'xl/comments1.xml': LEGACY_COMMENTS,
    'xl/drawings/vmlDrawing1.vml': VML,
    'xl/persons/person.xml': PERSONS,
    'xl/threadedComments/threadedComment1.xml': THREADED_COMMENTS,
    'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS,
    'xl/worksheets/sheet1.xml': WORKSHEET,
    ...overrides,
  };
}

async function bytes(
  overrides: XlsxBlackBoxOverrides = {},
): Promise<Uint8Array> {
  return createIndependentXlsx(parts(overrides));
}

async function capture(
  overrides: XlsxBlackBoxOverrides,
  options: Parameters<typeof parseXlsx>[1] = { errorMode: 'strict' },
): Promise<XlsxParseError> {
  try {
    await parseXlsx(await bytes(overrides), options);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected XLSX comment parsing to fail');
}

describe('XLSX comments', () => {
  it('parses legacy notes, VML visibility, threaded comments, and persons', async () => {
    const document = await parseXlsx(await bytes(), { errorMode: 'strict' });
    expect(document.workbook.commentPersons).toEqual([
      {
        displayName: 'Carol',
        id: 'person-1',
        providerId: 'AD',
        userId: 'carol@example.com',
      },
      { displayName: 'Dan', id: 'person-2' },
    ]);
    const sheet = document.sheets[0]!;
    expect(sheet.kind).toBe('worksheet');
    expect(sheet.kind === 'worksheet' ? sheet.comments : []).toEqual([
      {
        author: 'Alice & Co',
        kind: 'note',
        reference: 'A1',
        selectionRelation: 'full-sheet',
        text: 'Hello & world',
        visible: true,
      },
      {
        author: 'Bob',
        kind: 'note',
        reference: 'B2',
        selectionRelation: 'full-sheet',
        text: 'Hidden note',
        visible: false,
      },
      {
        id: 'thread-1',
        kind: 'threaded',
        personId: 'person-1',
        reference: 'C3',
        selectionRelation: 'full-sheet',
        text: 'Root thread',
        timestamp: '2024-01-02T03:04:05Z',
      },
      {
        id: 'thread-2',
        kind: 'threaded',
        parentId: 'thread-1',
        personId: 'person-2',
        reference: 'C3',
        selectionRelation: 'full-sheet',
        text: 'Reply',
        timestamp: '2024-01-02T04:05:06+01:00',
      },
    ]);
    expect(JSON.parse(JSON.stringify(document))).toEqual(document);
  });

  it('filters comments by selection while validating complete metadata', async () => {
    const document = await parseXlsx(await bytes(), {
      errorMode: 'strict',
      selection: { ranges: { Sheet1: ['B2:C3'] } },
    });
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.comments.map((comment) => [
            comment.reference,
            comment.selectionRelation,
          ])
        : [],
    ).toEqual([
      ['B2', 'intersects-selection'],
      ['C3', 'intersects-selection'],
      ['C3', 'intersects-selection'],
    ]);
    expect(document.workbook.commentPersons).toHaveLength(2);
  });

  it('round-trips comment metadata through portable exact R0', async () => {
    const source = await bytes();
    const snapshot = await readXlsxRoundTrip(source);
    const output = await writeXlsxRoundTrip(
      await validateXlsxRoundTripJson(
        JSON.parse(JSON.stringify(snapshot)) as unknown,
      ),
    );
    expect(output.data).toEqual(source);
    expect(output.report.level).toBe('R0');
  });

  it('parses prefixed Strict legacy comments and VML ownership', async () => {
    const strictSheet = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictRelationship =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const source = await createIndependentXlsx({
      '[Content_Types].xml': CONTENT_TYPES.replace(
        'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        strictSheet,
      ),
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelationship}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${strictRelationship}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/comments1.xml': `<s:comments xmlns:s="${strictSheet}"><s:authors><s:author>Strict</s:author></s:authors><s:commentList><s:comment ref="A1" authorId="0"><s:text><s:t>Note</s:t></s:text></s:comment></s:commentList></s:comments>`,
      'xl/drawings/vmlDrawing1.vml': VML,
      'xl/persons/person.xml': null,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/threadedComments/threadedComment1.xml': null,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheet}" xmlns:r="${strictRelationship}"><s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets></s:workbook>`,
      'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="comments" Type="${strictRelationship}/comments" Target="../comments1.xml"/><Relationship Id="vml" Type="${strictRelationship}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheet}" xmlns:r="${strictRelationship}"><s:sheetData/><s:legacyDrawing r:id="vml"/></s:worksheet>`,
    });
    const document = await parseXlsx(source, { errorMode: 'strict' });
    const sheet = document.sheets[0]!;
    expect(sheet.kind === 'worksheet' ? sheet.comments : []).toMatchObject([
      { author: 'Strict', kind: 'note', text: 'Note' },
    ]);
  });

  it('recovers malformed optional comments only in tolerant mode', async () => {
    const source = await bytes({ 'xl/comments1.xml': null });
    const result = await parseXlsxWithDiagnostics(source);
    const sheet = result.document.sheets[0]!;
    expect(sheet.kind === 'worksheet' ? sheet.comments : []).toEqual([]);
    expect(result.diagnostics).toMatchObject([
      {
        code: 'missing-required-part',
        part: 'xl/comments1.xml',
        severity: 'warning',
      },
    ]);
    expect(
      (await capture({ 'xl/comments1.xml': null })).diagnostic,
    ).toMatchObject({
      code: 'missing-required-part',
      part: 'xl/comments1.xml',
      severity: 'error',
    });
  });

  it('adds tolerant warnings for parser-owned person and comment failures', async () => {
    const malformedComments = await parseXlsxWithDiagnostics(
      await bytes({
        'xl/comments1.xml': `<wrong xmlns="${XLSX_SPREADSHEET_NS}"/>`,
      }),
    );
    expect(malformedComments.diagnostics).toMatchObject([
      {
        code: 'invalid-document-structure',
        message: 'comments root is missing',
        severity: 'warning',
      },
    ]);
    const malformedPersons = await parseXlsxWithDiagnostics(
      await bytes({
        'xl/persons/person.xml': `<personList xmlns="${THREADED_NAMESPACE}"><person displayName="" id="p"/></personList>`,
      }),
    );
    expect(malformedPersons.diagnostics).toMatchObject([
      {
        message: 'Comment person display name is invalid',
        severity: 'warning',
      },
      {
        message: 'Threaded comment person is missing',
        severity: 'warning',
      },
    ]);
  });

  it('defaults legacy visibility when no VML drawing is authored', async () => {
    const document = await parseXlsx(
      await bytes({
        'xl/worksheets/sheet1.xml': WORKSHEET.replace(
          '<legacyDrawing r:id="rIdVml"/>',
          '',
        ),
      }),
      { errorMode: 'strict' },
    );
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet'
        ? sheet.comments
            .filter((comment) => comment.kind === 'note')
            .map((comment) => comment.visible)
        : [],
    ).toEqual([false, false]);
  });

  it('parses legacy-only, threaded-only, and comment-free relationship graphs', async () => {
    const removeRelation = (xml: string, id: string) =>
      xml.replace(new RegExp(`\\s*<Relationship Id="${id}"[^>]+/>`, 'u'), '');
    const legacyOnly = await parseXlsx(
      await bytes({
        'xl/_rels/workbook.xml.rels': removeRelation(
          WORKBOOK_RELS,
          'rIdPersons',
        ),
        'xl/persons/person.xml': null,
        'xl/threadedComments/threadedComment1.xml': null,
        'xl/worksheets/_rels/sheet1.xml.rels': removeRelation(
          WORKSHEET_RELS,
          'rIdThreaded',
        ),
      }),
      { errorMode: 'strict' },
    );
    expect(legacyOnly.workbook.commentPersons).toBeUndefined();
    expect(
      legacyOnly.sheets[0]!.kind === 'worksheet'
        ? legacyOnly.sheets[0]!.comments.map((comment) => comment.kind)
        : [],
    ).toEqual(['note', 'note']);

    const threadedOnlyRels = removeRelation(
      removeRelation(WORKSHEET_RELS, 'rIdComments'),
      'rIdVml',
    );
    const threadedOnly = await parseXlsx(
      await bytes({
        'xl/comments1.xml': null,
        'xl/drawings/vmlDrawing1.vml': null,
        'xl/worksheets/_rels/sheet1.xml.rels': threadedOnlyRels,
        'xl/worksheets/sheet1.xml': WORKSHEET.replace(
          '<legacyDrawing r:id="rIdVml"/>',
          '',
        ),
      }),
      { errorMode: 'strict' },
    );
    expect(
      threadedOnly.sheets[0]!.kind === 'worksheet'
        ? threadedOnly.sheets[0]!.comments.map((comment) => comment.kind)
        : [],
    ).toEqual(['threaded', 'threaded']);

    const none = await parseXlsx(
      await bytes({
        'xl/_rels/workbook.xml.rels': removeRelation(
          WORKBOOK_RELS,
          'rIdPersons',
        ),
        'xl/comments1.xml': null,
        'xl/drawings/vmlDrawing1.vml': null,
        'xl/persons/person.xml': null,
        'xl/threadedComments/threadedComment1.xml': null,
        'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"/>`,
        'xl/worksheets/sheet1.xml': WORKSHEET.replace(
          '<legacyDrawing r:id="rIdVml"/>',
          '',
        ),
      }),
      { errorMode: 'strict' },
    );
    expect(
      none.sheets[0]!.kind === 'worksheet' ? none.sheets[0]!.comments : [],
    ).toEqual([]);
  });

  it.each([
    [
      {
        'xl/persons/person.xml': `<personList xmlns="${THREADED_NAMESPACE}"><person id="person-1"/></personList>`,
      },
      'Comment person display name is invalid',
    ],
    [
      {
        'xl/persons/person.xml': `<personList xmlns="${THREADED_NAMESPACE}"><person displayName="One" id="same"/><person displayName="Two" id="same"/></personList>`,
      },
      'Comment-person list contains duplicate IDs',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><commentList><comment ref="A1" authorId="0"><text><t>x</t></text></comment></commentList></comments>`,
      },
      'Legacy comments structure is invalid',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>A</author></authors><commentList><comment ref="$A$1" authorId="0"><text><t>x</t></text></comment></commentList></comments>`,
      },
      'Comment cell reference is invalid',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="1"><text><t>x</t></text></comment></commentList></comments>`,
      },
      'Legacy comment author reference is invalid',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"/></commentList></comments>`,
      },
      'Legacy comment text is missing',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><rPr/></text></comment></commentList></comments>`,
      },
      'Comment text is missing',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><t>x</t></text></comment><comment ref="A1" authorId="0"><text><t>y</t></text></comment></commentList></comments>`,
      },
      'Legacy comments contain duplicate cell references',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"><threadedComment ref="C3" dT="bad" personId="person-1" id="thread-1"><text>x</text></threadedComment></ThreadedComments>`,
      },
      'Threaded comment timestamp is invalid',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"><threadedComment ref="C3" dT="2024-01-02T03:04:05Z" personId="missing" id="thread-1"><text>x</text></threadedComment></ThreadedComments>`,
      },
      'Threaded comment person is missing',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"><threadedComment ref="C3" dT="2024-01-02T03:04:05Z" personId="person-1" id="same"><text>x</text></threadedComment><threadedComment ref="C3" dT="2024-01-02T03:04:06Z" personId="person-2" id="same"><text>y</text></threadedComment></ThreadedComments>`,
      },
      'Threaded comments contain duplicate IDs',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"><threadedComment ref="C3" dT="2024-01-02T03:04:05Z" personId="person-1" id="thread-1" parentId="missing"><text>x</text></threadedComment></ThreadedComments>`,
      },
      'Threaded comment parent reference is invalid',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"><threadedComment ref="C3" dT="2024-01-02T03:04:05Z" personId="person-1" id="thread-1"><text>x</text></threadedComment><threadedComment ref="D3" dT="2024-01-02T03:04:06Z" personId="person-2" id="thread-2" parentId="thread-1"><text>y</text></threadedComment></ThreadedComments>`,
      },
      'Threaded comment parent reference is invalid',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"><threadedComment ref="C3" dT="2024-01-02T03:04:05Z" personId="person-1" id="thread-1"/></ThreadedComments>`,
      },
      'Threaded comment text is missing or invalid',
    ],
    [
      {
        'xl/drawings/vmlDrawing1.vml': `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape><x:ClientData ObjectType="Note"><x:Row>bad</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>`,
      },
      'Comment VML row is invalid',
    ],
    [
      {
        'xl/drawings/vmlDrawing1.vml': `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape><x:ClientData ObjectType="Note"><x:Row>0</x:Row></x:ClientData></v:shape></xml>`,
      },
      'Comment VML column is invalid',
    ],
    [
      {
        'xl/worksheets/sheet1.xml': WORKSHEET.replace(
          '</worksheet>',
          '<legacyDrawing r:id="rIdVml"/></worksheet>',
        ),
      },
      'Worksheet contains duplicate legacyDrawing elements',
    ],
    [
      {
        'xl/worksheets/sheet1.xml': WORKSHEET.replace('rIdVml', 'missing'),
      },
      'Worksheet legacy comment drawing relationship is invalid',
    ],
    [
      {
        'xl/worksheets/sheet1.xml': WORKSHEET.replace('rIdVml', ''),
      },
      'Worksheet legacy drawing relationship reference is invalid',
    ],
  ] as const)(
    'rejects invalid comment contract %#',
    async (overrides, message) => {
      expect((await capture(overrides)).diagnostic.message).toBe(message);
    },
  );

  it('rejects external comment relationships without fetching', async () => {
    const external = WORKSHEET_RELS.replace(
      `Id="rIdComments" Type="${COMMENTS_RELATIONSHIP}" Target="../comments1.xml"`,
      `Id="rIdComments" Type="${COMMENTS_RELATIONSHIP}" Target="https://example.invalid/comments.xml" TargetMode="External"`,
    );
    expect(
      (
        await capture({
          'xl/worksheets/_rels/sheet1.xml.rels': external,
        })
      ).diagnostic.message,
    ).toBe('Worksheet legacy comments relationship must be internal');
  });

  it('accepts empty and exact-boundary person metadata without threaded comments', async () => {
    const withoutThreaded = WORKSHEET_RELS.replace(
      /\s*<Relationship Id="rIdThreaded"[^>]+\/>/u,
      '',
    );
    const empty = await parseXlsx(
      await bytes({
        'xl/persons/person.xml': `<personList xmlns="${THREADED_NAMESPACE}"/>`,
        'xl/threadedComments/threadedComment1.xml': null,
        'xl/worksheets/_rels/sheet1.xml.rels': withoutThreaded,
        'xl/worksheets/sheet1.xml': WORKSHEET.replace(
          '<legacyDrawing r:id="rIdVml"/>',
          '',
        ),
      }),
      { errorMode: 'strict' },
    );
    expect(empty.workbook.commentPersons).toBeUndefined();
    const boundaryId = 'p'.repeat(255);
    const exact = await parseXlsx(
      await bytes({
        'xl/persons/person.xml': `<personList xmlns="${THREADED_NAMESPACE}"><person displayName=" One " id="${boundaryId}"/><person displayName="Two" id="two"/></personList>`,
        'xl/threadedComments/threadedComment1.xml': null,
        'xl/worksheets/_rels/sheet1.xml.rels': withoutThreaded,
        'xl/worksheets/sheet1.xml': WORKSHEET.replace(
          '<legacyDrawing r:id="rIdVml"/>',
          '',
        ),
      }),
      { errorMode: 'strict', limits: { maxComments: 2 } },
    );
    expect(exact.workbook.commentPersons).toMatchObject([
      { displayName: ' One ', id: boundaryId },
      { displayName: 'Two', id: 'two' },
    ]);
    expect(
      (
        await capture(
          {
            'xl/persons/person.xml': `<personList xmlns="${THREADED_NAMESPACE}"><person displayName="One" id="one"/><person displayName="Two" id="two"/><person displayName="Three" id="three"/></personList>`,
            'xl/threadedComments/threadedComment1.xml': null,
            'xl/worksheets/_rels/sheet1.xml.rels': withoutThreaded,
            'xl/worksheets/sheet1.xml': WORKSHEET.replace(
              '<legacyDrawing r:id="rIdVml"/>',
              '',
            ),
          },
          { errorMode: 'strict', limits: { maxComments: 2 } },
        )
      ).diagnostic,
    ).toMatchObject({
      actual: 3,
      code: 'resource-limit-exceeded',
      limit: 2,
      limitName: 'maxComments',
    });
  });

  it.each([
    [
      {
        'xl/persons/person.xml': null,
      },
      'Required XLSX part is missing: xl/persons/person.xml',
    ],
    [
      {
        'xl/persons/person.xml': `<wrong xmlns="${THREADED_NAMESPACE}"/>`,
      },
      'personList root is missing',
    ],
    [
      {
        'xl/persons/person.xml': '<personList xmlns="urn:wrong"/>',
      },
      'personList root has the wrong namespace',
    ],
    [
      {
        'xl/persons/person.xml': `<personList xmlns="${THREADED_NAMESPACE}"><person>bad</person></personList>`,
      },
      'Comment-person list is invalid',
    ],
    [
      {
        'xl/persons/person.xml': `<personList xmlns="${THREADED_NAMESPACE}"><person displayName="A" id=""/></personList>`,
      },
      'Comment person ID is invalid',
    ],
    [
      {
        'xl/persons/person.xml': `<personList xmlns="${THREADED_NAMESPACE}"><person displayName="A" id=" spaced "/></personList>`,
      },
      'Comment person ID is invalid',
    ],
    [
      {
        'xl/persons/person.xml': `<personList xmlns="${THREADED_NAMESPACE}"><person displayName="A" id="${'p'.repeat(256)}"/></personList>`,
      },
      'Comment person ID is invalid',
    ],
    [
      {
        'xl/persons/person.xml': `<personList xmlns="${THREADED_NAMESPACE}"><person displayName="" id="p"/></personList>`,
      },
      'Comment person display name is invalid',
    ],
    [
      {
        'xl/comments1.xml': `<wrong xmlns="${XLSX_SPREADSHEET_NS}"/>`,
      },
      'comments root is missing',
    ],
    [
      {
        'xl/comments1.xml': '<comments xmlns="urn:wrong"/>',
      },
      'comments root has the wrong namespace',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors/><commentList><comment ref="A1" authorId="0"><text><t>x</t></text></comment></commentList></comments>`,
      },
      'Legacy comments structure is invalid',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>A</author></authors><commentList/></comments>`,
      },
      'Legacy comments structure is invalid',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author><bad/></author></authors><commentList><comment ref="A1" authorId="0"><text><t>x</t></text></comment></commentList></comments>`,
      },
      'Legacy comment author is invalid',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="01"><text><t>x</t></text></comment></commentList></comments>`,
      },
      'Legacy comment author reference is invalid',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>A</author></authors><commentList><comment>bad</comment></commentList></comments>`,
      },
      'Legacy comments structure is invalid',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><r>bad</r></text></comment></commentList></comments>`,
      },
      'Comment text is invalid',
    ],
    [
      {
        'xl/comments1.xml': `<comments xmlns="${XLSX_SPREADSHEET_NS}"><authors><author>A</author></authors><commentList><comment ref="A1" authorId="0"><text><t><bad/></t></text></comment></commentList></comments>`,
      },
      'Comment text is invalid',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<wrong xmlns="${THREADED_NAMESPACE}"/>`,
      },
      'ThreadedComments root is missing',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml':
          '<ThreadedComments xmlns="urn:wrong"/>',
      },
      'ThreadedComments root has the wrong namespace',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"/>`,
      },
      'Threaded-comment collection is invalid',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"><threadedComment ref="C3" dT="2024-01-02T03:04:05Z" personId="person-1" id=""><text>x</text></threadedComment></ThreadedComments>`,
      },
      'Threaded comment ID is invalid',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"><threadedComment ref="C3" dT="2024-01-02T03:04:05Z" personId=" person-1 " id="thread"><text>x</text></threadedComment></ThreadedComments>`,
      },
      'Threaded comment person reference is invalid',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"><threadedComment ref="C3" personId="person-1" id="thread"><text>x</text></threadedComment></ThreadedComments>`,
      },
      'Threaded comment timestamp is invalid',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"><threadedComment ref="C3" dT="2024-01-02T03:04:05Z" personId="person-1" id="thread" parentId=" "><text>x</text></threadedComment></ThreadedComments>`,
      },
      'Threaded comment parent ID is invalid',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"><threadedComment ref="C3" dT="2024-01-02T03:04:05Z" personId="person-1" id="root"><text>x</text></threadedComment><threadedComment ref="C4" dT="2024-01-02T03:04:06Z" personId="person-2" id="reply" parentId="root"><text>y</text></threadedComment></ThreadedComments>`,
      },
      'Threaded comment parent reference is invalid',
    ],
    [
      {
        'xl/threadedComments/threadedComment1.xml': `<ThreadedComments xmlns="${THREADED_NAMESPACE}"><threadedComment ref="C3" dT="2024-01-02T03:04:05Z" personId="person-1" id="root"><text>x</text></threadedComment><threadedComment ref="C3" dT="2024-01-02T03:04:06Z" personId="person-2" id="reply" parentId="root"><text>y</text></threadedComment><threadedComment ref="C3" dT="2024-01-02T03:04:07Z" personId="person-1" id="nested" parentId="reply"><text>z</text></threadedComment></ThreadedComments>`,
      },
      'Threaded comment parent reference is invalid',
    ],
  ] as const)(
    'rejects normalized comment edge %#',
    async (overrides, message) => {
      expect((await capture(overrides)).diagnostic.message).toBe(message);
    },
  );

  it.each([
    [
      {
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELS.replace(
          '</Relationships>',
          `<Relationship Id="rIdPersons2" Type="${PERSON_RELATIONSHIP}" Target="persons/person.xml"/></Relationships>`,
        ),
      },
      'Workbook comment-person list has duplicate relationships',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELS.replace(
          `Target="persons/person.xml"`,
          `Target="https://example.invalid/person.xml" TargetMode="External"`,
        ),
      },
      'Workbook comment-person list relationship must be internal',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'application/vnd.ms-excel.person+xml',
          'application/xml',
        ),
      },
      'Workbook comment-person list target has the wrong content type',
    ],
    [
      {
        'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replace(
          '</Relationships>',
          `<Relationship Id="rIdComments2" Type="${COMMENTS_RELATIONSHIP}" Target="../comments1.xml"/></Relationships>`,
        ),
      },
      'Worksheet legacy comments has duplicate relationships',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
          'application/xml',
        ),
      },
      'Worksheet legacy comments target has the wrong content type',
    ],
    [
      {
        'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replace(
          '</Relationships>',
          `<Relationship Id="rIdThreaded2" Type="${THREADED_RELATIONSHIP}" Target="../threadedComments/threadedComment1.xml"/></Relationships>`,
        ),
      },
      'Worksheet threaded comments has duplicate relationships',
    ],
    [
      {
        'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replace(
          `Target="../threadedComments/threadedComment1.xml"`,
          `Target="https://example.invalid/thread.xml" TargetMode="External"`,
        ),
      },
      'Worksheet threaded comments relationship must be internal',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'application/vnd.ms-excel.threadedcomments+xml',
          'application/xml',
        ),
      },
      'Worksheet threaded comments target has the wrong content type',
    ],
    [
      {
        'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replace(
          `Type="${VML_RELATIONSHIP}"`,
          `Type="${XLSX_OFFICE_REL_TYPE}drawing"`,
        ),
      },
      'Worksheet legacy comment drawing relationship is invalid',
    ],
    [
      {
        'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replace(
          `Target="../drawings/vmlDrawing1.vml"`,
          `Target="https://example.invalid/comments.vml" TargetMode="External"`,
        ),
      },
      'Worksheet legacy comment drawing relationship is invalid',
    ],
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'application/vnd.openxmlformats-officedocument.vmlDrawing',
          'application/xml',
        ),
      },
      'Worksheet legacy comment drawing has the wrong content type',
    ],
  ] as const)(
    'rejects comment relationship contract %#',
    async (overrides, message) => {
      expect((await capture(overrides)).diagnostic.message).toBe(message);
    },
  );

  it('enforces comment, text, selection-work, and grid limits exactly', async () => {
    const exact = await parseXlsx(await bytes(), {
      errorMode: 'strict',
      limits: {
        maxComments: 4,
        maxReturnedCells: 5,
        maxScannedCells: 5,
        maxTextCharacters: 190,
      },
      selection: { ranges: { Sheet1: ['A1:C3'] } },
    });
    expect(exact.sheets).toHaveLength(1);
    const commentOnlyParts: XlsxBlackBoxOverrides = {
      'xl/_rels/workbook.xml.rels': WORKBOOK_RELS.replace(
        /\s*<Relationship Id="rIdPersons"[^>]+\/>/u,
        '',
      ),
      'xl/drawings/vmlDrawing1.vml': null,
      'xl/persons/person.xml': null,
      'xl/threadedComments/threadedComment1.xml': null,
      'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replace(
        /\s*<Relationship Id="rIdVml"[^>]+\/>/u,
        '',
      ).replace(/\s*<Relationship Id="rIdThreaded"[^>]+\/>/u, ''),
      'xl/worksheets/sheet1.xml': WORKSHEET.replace(
        '<legacyDrawing r:id="rIdVml"/>',
        '',
      ),
    };
    const exactGrid = await parseXlsx(
      await bytes({
        ...commentOnlyParts,
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELS.replace(
          /\s*<Relationship Id="rIdPersons"[^>]+\/>/u,
          '',
        ),
        'xl/comments1.xml': LEGACY_COMMENTS.replace(
          /\s*<comment ref="B2"[\s\S]*?<\/comment>/u,
          '',
        ),
      }),
      {
        errorMode: 'strict',
        limits: { maxColumnsPerWorksheet: 1, maxRowsPerWorksheet: 1 },
      },
    );
    expect(exactGrid.sheets).toHaveLength(1);
    for (const [options, expected] of [
      [
        { limits: { maxComments: 3 } },
        { actual: 4, limit: 3, limitName: 'maxComments' },
      ],
      [
        {
          limits: { maxReturnedCells: 4, maxScannedCells: 4 },
          selection: { ranges: { Sheet1: ['A1:C3'] } },
        },
        { actual: 5, limit: 4, limitName: 'maxScannedCells' },
      ],
      [
        { limits: { maxTextCharacters: 189 } },
        { actual: 190, limit: 189, limitName: 'maxTextCharacters' },
      ],
    ] as const) {
      expect(
        (await capture({}, { errorMode: 'strict', ...options })).diagnostic,
      ).toMatchObject({ code: 'resource-limit-exceeded', ...expected });
    }
    for (const [override, limits, expected] of [
      [
        {
          'xl/comments1.xml': LEGACY_COMMENTS.replace('ref="A1"', 'ref="A2"'),
        },
        { maxRowsPerWorksheet: 1 },
        { actual: 2, limit: 1, limitName: 'maxRowsPerWorksheet' },
      ],
      [
        {
          'xl/comments1.xml': LEGACY_COMMENTS.replace('ref="A1"', 'ref="B1"'),
        },
        { maxColumnsPerWorksheet: 1 },
        { actual: 2, limit: 1, limitName: 'maxColumnsPerWorksheet' },
      ],
    ] as const) {
      expect(
        (
          await capture(
            { ...commentOnlyParts, ...override },
            { errorMode: 'strict', limits },
          )
        ).diagnostic,
      ).toMatchObject({ code: 'resource-limit-exceeded', ...expected });
    }
  });
});

describe('XLSX normalized comment metadata', () => {
  it.each([
    '0001-01-01T00:00:00Z',
    '2000-02-29T23:59:59.12Z',
    '2004-02-29T12:30:45-13:59',
    '2024-01-31T01:02:03+14:00',
    '2024-12-01T01:02:03Z',
  ])('accepts exact threaded timestamp boundary %s', (value) => {
    expect(isValidXlsxCommentTimestamp(value)).toBe(true);
  });

  it.each([
    'x2024-01-02T03:04:05Z',
    '2024-01-02T03:04:05Zx',
    '2024-01-02T03:04:05.xZ',
    '0000-01-01T00:00:00Z',
    '2024-00-01T00:00:00Z',
    '2024-13-01T00:00:00Z',
    '2024-01-00T00:00:00Z',
    '2024-04-31T00:00:00Z',
    '1900-02-29T00:00:00Z',
    '2001-02-29T00:00:00Z',
    '2024-01-01T24:00:00Z',
    '2024-01-01T00:60:00Z',
    '2024-01-01T00:00:60Z',
    '2024-01-01T00:00:00+15:00',
    '2024-01-01T00:00:00+14:01',
    '2024-01-01T00:00:00+13:60',
  ])('rejects threaded timestamp boundary %s', (value) => {
    expect(isValidXlsxCommentTimestamp(value)).toBe(false);
  });

  it('normalizes VML visibility sources and rejects lookalike CSS', () => {
    const client = (
      row: string,
      column: string,
      extra: Record<string, unknown> = {},
    ) => ({
      attrs: { ObjectType: 'Note' },
      'x:Column': column,
      'x:Row': row,
      ...extra,
    });
    const tree = {
      xml: {
        'v:shape': [
          {
            attrs: { style: ' visibility : visible ;' },
            'x:ClientData': client('0', '0'),
          },
          {
            attrs: { style: 'visibility:hidden' },
            'x:ClientData': client('1', '0', { 'x:Visible': { attrs: {} } }),
          },
          ...[
            'xvisibility:visible',
            'visibilityx:visible',
            'visibility:xvisible',
            'visibility:visiblex',
          ].map((style, index) => ({
            attrs: { style },
            'x:ClientData': client(String(index + 2), '0'),
          })),
          {
            attrs: { style: 'visibility:visible' },
            'x:ClientData': { attrs: { ObjectType: 'Button' } },
          },
        ],
        'v:rect': { 'x:ClientData': client('9', '9') },
      },
    } as unknown as XmlLookupValue;
    expect([
      ...parseXlsxCommentVmlVisibility(
        tree,
        'xl/drawings/comments.vml',
        defaultXlsxResourceLimits(),
      ),
    ]).toEqual([
      ['A1', true],
      ['A2', true],
      ['A3', false],
      ['A4', false],
      ['A5', false],
      ['A6', false],
    ]);
  });

  it('enforces VML duplicate, count, and grid boundaries', () => {
    const shape = (row: string, column: string) => ({
      'x:ClientData': {
        attrs: { ObjectType: 'Note' },
        'x:Column': column,
        'x:Row': row,
      },
    });
    const parse = (shapes: unknown[], limits = defaultXlsxResourceLimits()) =>
      parseXlsxCommentVmlVisibility(
        { xml: { 'v:shape': shapes } } as unknown as XmlLookupValue,
        'xl/drawings/comments.vml',
        limits,
      );
    expect(() => parse([shape('0', '0'), shape('0', '0')])).toThrow(
      'Comment VML contains duplicate anchors',
    );
    expect(() =>
      parse([shape('0', '0'), shape('1', '0')], {
        ...defaultXlsxResourceLimits(),
        maxComments: 1,
      }),
    ).toThrow(/maxComments/u);
    expect(() =>
      parse([shape('1', '0')], {
        ...defaultXlsxResourceLimits(),
        maxRowsPerWorksheet: 1,
      }),
    ).toThrow(/maxRowsPerWorksheet/u);
    expect(() =>
      parse([shape('0', '1')], {
        ...defaultXlsxResourceLimits(),
        maxColumnsPerWorksheet: 1,
      }),
    ).toThrow(/maxColumnsPerWorksheet/u);
    for (const row of ['x1', '1x', '1a', '4294967296']) {
      expect(() => parse([shape(row, '0')])).toThrow(
        'Comment VML row is invalid',
      );
    }
    expect(() => parse([shape('4294967295', '0')])).toThrow(
      /maxRowsPerWorksheet/u,
    );
    expect([
      ...parse([shape('0', '0')], {
        ...defaultXlsxResourceLimits(),
        maxColumnsPerWorksheet: 1,
        maxComments: 1,
        maxRowsPerWorksheet: 1,
      }),
    ]).toEqual([['A1', false]]);
  });
});
