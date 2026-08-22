import type { XmlLookupValue } from '../../../common/xml/tree';
import { decodeXmlEntities } from '../../../common/text/html';
import { XlsxParseError } from '../errors';
import type {
  XlsxComment,
  XlsxCommentPerson,
  XlsxLegacyComment,
  XlsxThreadedComment,
} from '../types';
import { parseXlsxCellReference, xlsxColumnName } from './cell-reference';
import { getXlsxRelationshipPartName } from './package-identity';
import { XlsxPartReader } from './part-reader';
import { parseXlsxRelationships, type XlsxRelationship } from './relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import {
  type XlsxResolvedSheetSelection,
  xlsxSelectionIncludesCell,
} from './selection';
import {
  consumeXlsxWorksheetBudget,
  type XlsxWorksheetBudget,
} from './worksheet';
import {
  type XlsxWorkbookDiscovery,
  XLSX_SPREADSHEET_NAMESPACES,
} from './workbook-discovery';

type XmlRecord = Record<string, unknown>;
type XmlAttributes = Record<string, string>;
type IncludedSelection = Exclude<
  XlsxResolvedSheetSelection,
  { kind: 'not-selected' }
>;

const OFFICE_RELATIONSHIPS = {
  strict: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
  transitional:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
} as const;
const THREADED_COMMENTS_RELATIONSHIP =
  'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
const PERSON_RELATIONSHIP =
  'http://schemas.microsoft.com/office/2017/10/relationships/person';
const THREADED_COMMENTS_NAMESPACE =
  'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
const THREADED_COMMENTS_CONTENT_TYPE =
  'application/vnd.ms-excel.threadedcomments+xml';
const PERSON_CONTENT_TYPE = 'application/vnd.ms-excel.person+xml';
const VML_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.vmlDrawing';
const UNSIGNED_INTEGER_PATTERN = /^(?:0|[1-9]\d*)$/u;
const ISO_TIMESTAMP_PATTERN =
  /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(Z|([+-])(\d{2}):(\d{2}))$/u;

export interface XlsxCommentBudget {
  comments: number;
}

export interface XlsxCommentPersonTable {
  byId: ReadonlyMap<string, XlsxCommentPerson>;
  values: XlsxCommentPerson[];
}

function fail(
  code: 'invalid-document-structure' | 'invalid-document-value',
  message: string,
  part: string,
): never {
  throw new XlsxParseError({ code, message, part, severity: 'error' });
}

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function records(value: unknown): XmlRecord[] | undefined {
  if (value === undefined) return [];
  const values = Array.isArray(value) ? value : [value];
  const output: XmlRecord[] = [];
  for (const item of values) {
    const parsed = record(item);
    if (!parsed) return undefined;
    output.push(parsed);
  }
  return output;
}

function items(value: unknown): unknown[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}

function scalarText(value: unknown): string | undefined {
  if (typeof value === 'string') return decodeXmlEntities(value);
  const node = record(value);
  return typeof node?.value === 'string'
    ? decodeXmlEntities(node.value)
    : undefined;
}

function attributes(node: XmlRecord): XmlAttributes {
  return Object.fromEntries(
    Object.entries(record(node.attrs) ?? {}).map(([name, value]) => [
      name,
      decodeXmlEntities(value as string),
    ]),
  );
}

function localName(name: string): string {
  return name.slice(name.lastIndexOf(':') + 1);
}

function child(node: XmlRecord, prefix: string, name: string): unknown {
  return node[prefix ? `${prefix}:${name}` : name];
}

function childByLocal(node: XmlRecord, name: string): unknown {
  const entry = Object.entries(node).find(([key]) => localName(key) === name);
  return entry?.[1];
}

function root(
  value: XmlLookupValue,
  expectedName: string,
  expectedNamespace: string,
  part: string,
): { node: XmlRecord; prefix: string } {
  const document = value as unknown as XmlRecord;
  const entry = Object.entries(document).find(
    ([name]) => localName(name) === expectedName,
  );
  const node = record(entry?.[1]);
  if (!entry || !node) {
    fail('invalid-document-structure', `${expectedName} root is missing`, part);
  }
  const pieces = entry[0].split(':');
  const prefix = pieces.length === 1 ? '' : pieces[0]!;
  const namespace = attributes(node)[prefix ? `xmlns:${prefix}` : 'xmlns'];
  if (namespace !== expectedNamespace) {
    fail(
      'invalid-document-structure',
      `${expectedName} root has the wrong namespace`,
      part,
    );
  }
  return { node, prefix };
}

function token(
  value: string | undefined,
  message: string,
  part: string,
): string {
  if (
    value === undefined ||
    value.length === 0 ||
    value.length > 255 ||
    value !== value.trim()
  ) {
    fail('invalid-document-value', message, part);
  }
  return value;
}

function requiredText(
  value: string | undefined,
  message: string,
  part: string,
): string {
  if (value === undefined || value.length === 0) {
    fail('invalid-document-value', message, part);
  }
  return value;
}

function optionalText(value: string | undefined): string | undefined {
  return value;
}

function unsignedInteger(
  value: string | undefined,
  message: string,
  part: string,
): number {
  if (value === undefined || !UNSIGNED_INTEGER_PATTERN.test(value)) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (parsed > 0xffff_ffff) {
    fail('invalid-document-value', message, part);
  }
  return parsed;
}

function consumeText(
  amount: number,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): void {
  consumeXlsxWorksheetBudget(
    budget,
    'textCharacters',
    amount,
    'maxTextCharacters',
    limits,
    part,
  );
}

function consumeComment(
  budget: XlsxCommentBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): void {
  const actual = budget.comments + 1;
  if (!Number.isSafeInteger(actual) || actual > limits.maxComments) {
    throw new XlsxResourceLimitError(
      'maxComments',
      actual,
      limits.maxComments,
      part,
    );
  }
  budget.comments = actual;
}

export function isValidXlsxCommentTimestamp(value: string): boolean {
  const match = ISO_TIMESTAMP_PATTERN.exec(value);
  if (!match) return false;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);
  const offsetHour = match[8] === undefined ? 0 : Number(match[9]);
  const offsetMinute = match[8] === undefined ? 0 : Number(match[10]);
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const monthDays = [
    31,
    leap ? 29 : 28,
    31,
    30,
    31,
    30,
    31,
    31,
    30,
    31,
    30,
    31,
  ];
  return (
    year >= 1 &&
    day >= 1 &&
    day <= monthDays[month - 1]! &&
    hour <= 23 &&
    minute <= 59 &&
    second <= 59 &&
    offsetHour <= 14 &&
    offsetMinute <= 59 &&
    (offsetHour !== 14 || offsetMinute === 0)
  );
}

function relationship(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  types: ReadonlySet<string>,
  expectedContentType: string,
  discovery: XlsxWorkbookDiscovery,
  ownerPart: string,
  label: string,
): Extract<XlsxRelationship, { mode: 'internal' }> | undefined {
  const matches = [...relationships.values()].filter((candidate) =>
    types.has(candidate.type),
  );
  if (matches.length > 1) {
    fail(
      'invalid-document-structure',
      `${label} has duplicate relationships`,
      getXlsxRelationshipPartName(ownerPart),
    );
  }
  const match = matches[0];
  if (match === undefined) return undefined;
  if (match.mode !== 'internal') {
    fail(
      'invalid-document-structure',
      `${label} relationship must be internal`,
      getXlsxRelationshipPartName(ownerPart),
    );
  }
  if (
    discovery.contentTypes.contentTypeFor(match.target) !== expectedContentType
  ) {
    fail(
      'invalid-document-structure',
      `${label} target has the wrong content type`,
      match.target,
    );
  }
  return match;
}

function relationTypes(...types: string[]): ReadonlySet<string> {
  return new Set(types);
}

export async function loadXlsxCommentPersons(
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  worksheetBudget: XlsxWorksheetBudget,
): Promise<XlsxCommentPersonTable> {
  const relationshipPart = getXlsxRelationshipPartName(discovery.part);
  const relationshipsXml = (await reader.readXml(relationshipPart))!;
  const relationships = parseXlsxRelationships(
    relationshipsXml,
    discovery.part,
    limits.maxRelationships,
  );
  const personRelationship = relationship(
    relationships,
    relationTypes(PERSON_RELATIONSHIP),
    PERSON_CONTENT_TYPE,
    discovery,
    discovery.part,
    'Workbook comment-person list',
  );
  if (!personRelationship) return { byId: new Map(), values: [] };
  const xml = await reader.readXml(personRelationship.target, {
    required: true,
  });
  const { node, prefix } = root(
    xml,
    'personList',
    THREADED_COMMENTS_NAMESPACE,
    personRelationship.target,
  );
  const nodes = records(child(node, prefix, 'person'));
  if (!nodes) {
    fail(
      'invalid-document-structure',
      'Comment-person list is invalid',
      personRelationship.target,
    );
  }
  if (nodes.length > limits.maxComments) {
    throw new XlsxResourceLimitError(
      'maxComments',
      nodes.length,
      limits.maxComments,
      personRelationship.target,
    );
  }
  const byId = new Map<string, XlsxCommentPerson>();
  const values: XlsxCommentPerson[] = [];
  for (const personNode of nodes) {
    const attrs = attributes(personNode);
    const id = token(
      attrs.id,
      'Comment person ID is invalid',
      personRelationship.target,
    );
    const displayName = requiredText(
      attrs.displayName,
      'Comment person display name is invalid',
      personRelationship.target,
    );
    if (byId.has(id)) {
      fail(
        'invalid-document-value',
        'Comment-person list contains duplicate IDs',
        personRelationship.target,
      );
    }
    const providerId = optionalText(attrs.providerId);
    const userId = optionalText(attrs.userId);
    consumeText(
      id.length +
        displayName.length +
        (providerId?.length ?? 0) +
        (userId?.length ?? 0),
      worksheetBudget,
      limits,
      personRelationship.target,
    );
    const person: XlsxCommentPerson = {
      displayName,
      id,
      ...(providerId === undefined ? {} : { providerId }),
      ...(userId === undefined ? {} : { userId }),
    };
    byId.set(id, person);
    values.push(person);
  }
  return { byId, values };
}

function commentReference(
  value: unknown,
  part: string,
  limits: ResolvedXlsxResourceLimits,
): { column: number; reference: string; row: number } {
  const parsed = parseXlsxCellReference(value);
  if (!parsed || parsed.absoluteColumn || parsed.absoluteRow) {
    fail('invalid-document-value', 'Comment cell reference is invalid', part);
  }
  if (parsed.row > limits.maxRowsPerWorksheet) {
    throw new XlsxResourceLimitError(
      'maxRowsPerWorksheet',
      parsed.row,
      limits.maxRowsPerWorksheet,
      part,
    );
  }
  if (parsed.column > limits.maxColumnsPerWorksheet) {
    throw new XlsxResourceLimitError(
      'maxColumnsPerWorksheet',
      parsed.column,
      limits.maxColumnsPerWorksheet,
      part,
    );
  }
  return {
    column: parsed.column,
    reference: parsed.address,
    row: parsed.row,
  };
}

function commentText(node: XmlRecord, part: string): string {
  let output = '';
  let found = false;
  const visit = (value: XmlRecord): void => {
    const children = { ...value };
    delete children.attrs;
    delete children.value;
    for (const [name, childValue] of Object.entries(children)) {
      for (const childItem of items(childValue)) {
        if (localName(name) === 't') {
          const text = scalarText(childItem);
          if (text === undefined) {
            fail('invalid-document-structure', 'Comment text is invalid', part);
          }
          found = true;
          output += text;
        } else {
          const childNode = record(childItem);
          if (!childNode) {
            fail('invalid-document-structure', 'Comment text is invalid', part);
          }
          visit(childNode);
        }
      }
    }
  };
  visit(node);
  if (!found)
    fail('invalid-document-structure', 'Comment text is missing', part);
  return output;
}

function selectionRelation(
  selection: IncludedSelection,
  row: number,
  column: number,
  worksheetBudget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxComment['selectionRelation'] | null {
  if (selection.kind === 'full-sheet') return 'full-sheet';
  consumeXlsxWorksheetBudget(
    worksheetBudget,
    'scannedCells',
    1,
    'maxScannedCells',
    limits,
    part,
  );
  return xlsxSelectionIncludesCell(selection, row, column)
    ? 'intersects-selection'
    : null;
}

export function parseXlsxCommentVmlVisibility(
  value: XmlLookupValue,
  part: string,
  limits: ResolvedXlsxResourceLimits,
): ReadonlyMap<string, boolean> {
  const output = new Map<string, boolean>();
  let noteShapes = 0;
  const visit = (node: XmlRecord): void => {
    for (const [name, childValue] of Object.entries(node)) {
      for (const childItem of items(childValue)) {
        if (typeof childItem === 'string') continue;
        const childNode = childItem as XmlRecord;
        if (localName(name) === 'shape') {
          const client = record(childByLocal(childNode, 'ClientData'));
          if (client && attributes(client).ObjectType === 'Note') {
            noteShapes += 1;
            if (noteShapes > limits.maxComments) {
              throw new XlsxResourceLimitError(
                'maxComments',
                noteShapes,
                limits.maxComments,
                part,
              );
            }
            const row = unsignedInteger(
              scalarText(childByLocal(client, 'Row')),
              'Comment VML row is invalid',
              part,
            );
            const column = unsignedInteger(
              scalarText(childByLocal(client, 'Column')),
              'Comment VML column is invalid',
              part,
            );
            if (row >= limits.maxRowsPerWorksheet) {
              throw new XlsxResourceLimitError(
                'maxRowsPerWorksheet',
                row + 1,
                limits.maxRowsPerWorksheet,
                part,
              );
            }
            if (column >= limits.maxColumnsPerWorksheet) {
              throw new XlsxResourceLimitError(
                'maxColumnsPerWorksheet',
                column + 1,
                limits.maxColumnsPerWorksheet,
                part,
              );
            }
            const reference = `${xlsxColumnName(column + 1)!}${row + 1}`;
            if (output.has(reference)) {
              fail(
                'invalid-document-value',
                'Comment VML contains duplicate anchors',
                part,
              );
            }
            const style = attributes(childNode).style;
            const visibleElement =
              childByLocal(client, 'Visible') !== undefined;
            const visibleStyle =
              typeof style === 'string' &&
              /(?:^|;)\s*visibility\s*:\s*visible\s*(?:;|$)/iu.test(style);
            output.set(reference, visibleElement || visibleStyle);
          }
        }
        visit(childNode);
      }
    }
  };
  visit(value);
  return output;
}

async function legacyVisibility(
  relationshipIds: readonly string[],
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  worksheetPart: string,
): Promise<ReadonlyMap<string, boolean>> {
  if (relationshipIds.length === 0) return new Map();
  const relation = relationships.get(relationshipIds[0]!);
  const vmlType = `${OFFICE_RELATIONSHIPS[discovery.dialect]}/vmlDrawing`;
  if (!relation || relation.mode !== 'internal' || relation.type !== vmlType) {
    fail(
      'invalid-document-structure',
      'Worksheet legacy comment drawing relationship is invalid',
      worksheetPart,
    );
  }
  if (
    discovery.contentTypes.contentTypeFor(relation.target) !== VML_CONTENT_TYPE
  ) {
    fail(
      'invalid-document-structure',
      'Worksheet legacy comment drawing has the wrong content type',
      relation.target,
    );
  }
  const xml = await reader.readXml(relation.target, { required: true });
  return parseXlsxCommentVmlVisibility(xml, relation.target, limits);
}

async function loadLegacyComments(
  relationshipValue:
    Extract<XlsxRelationship, { mode: 'internal' }> | undefined,
  visibility: ReadonlyMap<string, boolean>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  commentBudget: XlsxCommentBudget,
  worksheetBudget: XlsxWorksheetBudget,
  selection: IncludedSelection,
): Promise<XlsxLegacyComment[]> {
  if (!relationshipValue) return [];
  const xml = await reader.readXml(relationshipValue.target, {
    required: true,
  });
  const { node, prefix } = root(
    xml,
    'comments',
    XLSX_SPREADSHEET_NAMESPACES[discovery.dialect],
    relationshipValue.target,
  );
  const authorsNode = record(child(node, prefix, 'authors'));
  const commentListNode = record(child(node, prefix, 'commentList'));
  const authorNodes = authorsNode
    ? items(child(authorsNode, prefix, 'author'))
    : undefined;
  const commentNodes = commentListNode
    ? records(child(commentListNode, prefix, 'comment'))
    : undefined;
  if (
    !authorNodes ||
    authorNodes.length === 0 ||
    !commentNodes ||
    commentNodes.length === 0
  ) {
    fail(
      'invalid-document-structure',
      'Legacy comments structure is invalid',
      relationshipValue.target,
    );
  }
  const authors = authorNodes.map((authorNode) => {
    const author = scalarText(authorNode);
    if (author === undefined) {
      fail(
        'invalid-document-structure',
        'Legacy comment author is invalid',
        relationshipValue.target,
      );
    }
    return author;
  });
  const references = new Set<string>();
  const output: XlsxLegacyComment[] = [];
  for (const commentNode of commentNodes) {
    consumeComment(commentBudget, limits, relationshipValue.target);
    const attrs = attributes(commentNode);
    const reference = commentReference(
      attrs.ref,
      relationshipValue.target,
      limits,
    );
    if (references.has(reference.reference)) {
      fail(
        'invalid-document-value',
        'Legacy comments contain duplicate cell references',
        relationshipValue.target,
      );
    }
    references.add(reference.reference);
    const authorId = unsignedInteger(
      attrs.authorId,
      'Legacy comment author reference is invalid',
      relationshipValue.target,
    );
    if (authorId >= authors.length) {
      fail(
        'invalid-document-value',
        'Legacy comment author reference is invalid',
        relationshipValue.target,
      );
    }
    const textNode = record(child(commentNode, prefix, 'text'));
    if (!textNode) {
      fail(
        'invalid-document-structure',
        'Legacy comment text is missing',
        relationshipValue.target,
      );
    }
    const text = commentText(textNode, relationshipValue.target);
    consumeText(
      authors[authorId]!.length + text.length,
      worksheetBudget,
      limits,
      relationshipValue.target,
    );
    const relation = selectionRelation(
      selection,
      reference.row,
      reference.column,
      worksheetBudget,
      limits,
      relationshipValue.target,
    );
    if (relation !== null) {
      output.push({
        author: authors[authorId]!,
        kind: 'note',
        reference: reference.reference,
        selectionRelation: relation,
        text,
        visible: visibility.get(reference.reference) ?? false,
      });
    }
  }
  return output;
}

async function loadThreadedComments(
  relationshipValue:
    Extract<XlsxRelationship, { mode: 'internal' }> | undefined,
  persons: XlsxCommentPersonTable,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  commentBudget: XlsxCommentBudget,
  worksheetBudget: XlsxWorksheetBudget,
  selection: IncludedSelection,
): Promise<XlsxThreadedComment[]> {
  if (!relationshipValue) return [];
  const xml = await reader.readXml(relationshipValue.target, {
    required: true,
  });
  const { node, prefix } = root(
    xml,
    'ThreadedComments',
    THREADED_COMMENTS_NAMESPACE,
    relationshipValue.target,
  );
  const commentNodes = records(child(node, prefix, 'threadedComment'));
  if (!commentNodes || commentNodes.length === 0) {
    fail(
      'invalid-document-structure',
      'Threaded-comment collection is invalid',
      relationshipValue.target,
    );
  }
  const ids = new Set<string>();
  const parsed: Array<{
    comment: XlsxThreadedComment;
    column: number;
    parentId?: string;
    relation: XlsxComment['selectionRelation'] | null;
    row: number;
  }> = [];
  for (const commentNode of commentNodes) {
    consumeComment(commentBudget, limits, relationshipValue.target);
    const attrs = attributes(commentNode);
    const id = token(
      attrs.id,
      'Threaded comment ID is invalid',
      relationshipValue.target,
    );
    if (ids.has(id)) {
      fail(
        'invalid-document-value',
        'Threaded comments contain duplicate IDs',
        relationshipValue.target,
      );
    }
    ids.add(id);
    const personId = token(
      attrs.personId,
      'Threaded comment person reference is invalid',
      relationshipValue.target,
    );
    if (!persons.byId.has(personId)) {
      fail(
        'invalid-document-value',
        'Threaded comment person is missing',
        relationshipValue.target,
      );
    }
    const timestamp = token(
      attrs.dT,
      'Threaded comment timestamp is invalid',
      relationshipValue.target,
    );
    if (!isValidXlsxCommentTimestamp(timestamp)) {
      fail(
        'invalid-document-value',
        'Threaded comment timestamp is invalid',
        relationshipValue.target,
      );
    }
    const rawParentId = optionalText(attrs.parentId);
    const parentId =
      rawParentId === undefined
        ? undefined
        : token(
            rawParentId,
            'Threaded comment parent ID is invalid',
            relationshipValue.target,
          );
    const reference = commentReference(
      attrs.ref,
      relationshipValue.target,
      limits,
    );
    const text = scalarText(child(commentNode, prefix, 'text'));
    if (text === undefined) {
      fail(
        'invalid-document-structure',
        'Threaded comment text is missing or invalid',
        relationshipValue.target,
      );
    }
    consumeText(
      id.length +
        personId.length +
        timestamp.length +
        (parentId?.length ?? 0) +
        text.length,
      worksheetBudget,
      limits,
      relationshipValue.target,
    );
    const relation = selectionRelation(
      selection,
      reference.row,
      reference.column,
      worksheetBudget,
      limits,
      relationshipValue.target,
    );
    parsed.push({
      column: reference.column,
      comment: {
        id,
        kind: 'threaded',
        ...(parentId === undefined ? {} : { parentId }),
        personId,
        reference: reference.reference,
        selectionRelation: relation ?? 'full-sheet',
        text,
        timestamp,
      },
      ...(parentId === undefined ? {} : { parentId }),
      relation,
      row: reference.row,
    });
  }
  const byId = new Map(parsed.map((item) => [item.comment.id, item]));
  for (const item of parsed) {
    if (item.parentId === undefined) continue;
    const parent = byId.get(item.parentId);
    if (
      !parent ||
      parent.row !== item.row ||
      parent.column !== item.column ||
      parent.comment.parentId !== undefined
    ) {
      fail(
        'invalid-document-value',
        'Threaded comment parent reference is invalid',
        relationshipValue.target,
      );
    }
  }
  return parsed.flatMap((item) => {
    return item.relation === null
      ? []
      : [{ ...item.comment, selectionRelation: item.relation }];
  });
}

export async function loadXlsxWorksheetComments(
  worksheetPart: string,
  legacyDrawingRelationshipIds: readonly string[],
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  commentBudget: XlsxCommentBudget,
  worksheetBudget: XlsxWorksheetBudget,
  selection: IncludedSelection,
  persons: XlsxCommentPersonTable,
): Promise<XlsxComment[]> {
  const relationshipBase = OFFICE_RELATIONSHIPS[discovery.dialect];
  const legacyRelationship = relationship(
    relationships,
    relationTypes(`${relationshipBase}/comments`),
    COMMENTS_CONTENT_TYPE,
    discovery,
    worksheetPart,
    'Worksheet legacy comments',
  );
  const threadedRelationship = relationship(
    relationships,
    relationTypes(THREADED_COMMENTS_RELATIONSHIP),
    THREADED_COMMENTS_CONTENT_TYPE,
    discovery,
    worksheetPart,
    'Worksheet threaded comments',
  );
  const visibility = await legacyVisibility(
    legacyDrawingRelationshipIds,
    relationships,
    discovery,
    reader,
    limits,
    worksheetPart,
  );
  return [
    ...(await loadLegacyComments(
      legacyRelationship,
      visibility,
      discovery,
      reader,
      limits,
      commentBudget,
      worksheetBudget,
      selection,
    )),
    ...(await loadThreadedComments(
      threadedRelationship,
      persons,
      reader,
      limits,
      commentBudget,
      worksheetBudget,
      selection,
    )),
  ];
}
