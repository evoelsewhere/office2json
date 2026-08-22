import { encodeBase64 } from '../../../common/binary/base64';
import { decodeXmlEntities } from '../../../common/text/html';
import type { XmlLookupValue } from '../../../common/xml/tree';
import { getXmlNodeOrder } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type {
  XlsxChart,
  XlsxDrawing,
  XlsxDrawingColor,
  XlsxDrawingConnector,
  XlsxDrawingExtent,
  XlsxDrawingFill,
  XlsxDrawingGroup,
  XlsxDrawingLine,
  XlsxDrawingMarker,
  XlsxDrawingObject,
  XlsxDrawingObjectTransform,
  XlsxDrawingShape,
  XlsxEmbeddedImage,
  XlsxImageCrop,
  XlsxImageMode,
} from '../types';
import { loadXlsxChart } from './chart';
import { getXlsxRelationshipPartName } from './package-identity';
import { XlsxPartReader } from './part-reader';
import { parseXlsxRelationships, type XlsxRelationship } from './relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import type { XlsxResolvedSheetSelection } from './selection';
import {
  consumeXlsxWorksheetBudget,
  type XlsxWorksheetBudget,
} from './worksheet';
import type { XlsxWorkbookDiscovery } from './workbook-discovery';

type XmlRecord = Record<string, unknown>;
type IncludedSelection = Exclude<
  XlsxResolvedSheetSelection,
  { kind: 'not-selected' }
>;

const DRAWING_NAMESPACES = {
  strict: 'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing',
  transitional:
    'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
} as const;
const OFFICE_RELATIONSHIPS = {
  strict: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
  transitional:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
} as const;
const DRAWING_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawing+xml';
const CHART_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const EMU_PER_POINT = 12_700;
const DRAWING_OBJECT_NAMES = [
  'pic',
  'sp',
  'cxnSp',
  'grpSp',
  'graphicFrame',
] as const;

const SAFE_IMAGE_CONTENT_TYPES = new Set([
  'image/bmp',
  'image/gif',
  'image/jpeg',
  'image/png',
  'image/tiff',
  'image/webp',
  'image/x-emf',
  'image/x-wmf',
]);

export function isSafeXlsxImageContentType(contentType: string): boolean {
  return SAFE_IMAGE_CONTENT_TYPES.has(contentType);
}

export interface XlsxDrawingBudget {
  charts: number;
  drawings: number;
}

interface LoadedImagePayload {
  base64?: string;
  blobUrl?: string;
  byteLength: number;
}

export interface XlsxMediaCheckpoint {
  cacheKeys: Set<string>;
  objectUrls: Set<string>;
  returnedBytes: number;
}

export interface XlsxMediaReader {
  hasPart(part: string): boolean;
  readBytes(
    part: string,
    limitName: 'maxMediaBytes',
  ): Promise<Uint8Array | null>;
}

export class XlsxMediaSession {
  private readonly cache = new Map<string, LoadedImagePayload>();
  private readonly objectUrls = new Set<string>();
  private returnedBytes = 0;

  constructor(
    private readonly mode: XlsxImageMode,
    private readonly limits: ResolvedXlsxResourceLimits,
  ) {}

  checkpoint(): XlsxMediaCheckpoint {
    return {
      cacheKeys: new Set(this.cache.keys()),
      objectUrls: new Set(this.objectUrls),
      returnedBytes: this.returnedBytes,
    };
  }

  rollback(checkpoint: XlsxMediaCheckpoint): void {
    for (const url of this.objectUrls) {
      if (!checkpoint.objectUrls.has(url)) {
        URL.revokeObjectURL(url);
        this.objectUrls.delete(url);
      }
    }
    for (const key of this.cache.keys()) {
      if (!checkpoint.cacheKeys.has(key)) this.cache.delete(key);
    }
    this.returnedBytes = checkpoint.returnedBytes;
  }

  async image(
    part: string,
    contentType: string,
    reader: XlsxMediaReader,
    includePayload = true,
  ): Promise<LoadedImagePayload | undefined> {
    if (!reader.hasPart(part)) {
      throw new XlsxParseError({
        code: 'missing-required-part',
        message: `Required XLSX image part is missing: ${part}`,
        part,
        severity: 'error',
      });
    }
    if (this.mode === 'none' || !includePayload) return undefined;
    const cached = this.cache.get(part);
    const payload = cached ?? (await this.load(part, contentType, reader));
    this.cache.set(part, payload);
    const actual = this.returnedBytes + payload.byteLength;
    if (!Number.isSafeInteger(actual) || actual > this.limits.maxMediaBytes) {
      throw new XlsxResourceLimitError(
        'maxMediaBytes',
        actual,
        this.limits.maxMediaBytes,
        part,
      );
    }
    this.returnedBytes = actual;
    return payload;
  }

  revokeAll(): void {
    for (const url of this.objectUrls) URL.revokeObjectURL(url);
    this.objectUrls.clear();
  }

  private async load(
    part: string,
    contentType: string,
    reader: XlsxMediaReader,
  ): Promise<LoadedImagePayload> {
    const bytes = await reader.readBytes(part, 'maxMediaBytes');
    if (!bytes) {
      throw new XlsxParseError({
        code: 'missing-required-part',
        message: `Required XLSX image part is missing: ${part}`,
        part,
        severity: 'error',
      });
    }
    const output: LoadedImagePayload = { byteLength: bytes.byteLength };
    if (this.mode === 'base64' || this.mode === 'both') {
      output.base64 = `data:${contentType};base64,${encodeBase64(bytes)}`;
    }
    if (this.mode === 'blob' || this.mode === 'both') {
      const copy = bytes.slice().buffer;
      const url = URL.createObjectURL(new Blob([copy], { type: contentType }));
      this.objectUrls.add(url);
      output.blobUrl = url;
    }
    return output;
  }
}

function fail(
  code:
    | 'invalid-document-structure'
    | 'invalid-document-value'
    | 'security-rejected-content',
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
  const values = Array.isArray(value) ? value : [value];
  const output: XmlRecord[] = [];
  for (const item of values) {
    const parsed = record(item);
    if (!parsed) return undefined;
    output.push(parsed);
  }
  return output;
}

function attributes(node: XmlRecord): Record<string, string> {
  return (record(node.attrs) ?? {}) as Record<string, string>;
}

function localName(name: string): string {
  return name.slice(name.lastIndexOf(':') + 1);
}

function childByLocal(node: XmlRecord, name: string): unknown {
  return Object.entries(node).find(([key]) => localName(key) === name)?.[1];
}

function scalarText(value: unknown): string | undefined {
  if (typeof value === 'string') return decodeXmlEntities(value);
  const node = record(value);
  return typeof node?.value === 'string'
    ? decodeXmlEntities(node.value)
    : undefined;
}

function root(
  value: XmlLookupValue,
  expectedNamespace: string,
  part: string,
): XmlRecord {
  const entry = Object.entries(value).find(
    ([name]) => localName(name) === 'wsDr',
  );
  const node = record(entry?.[1]);
  if (!entry || !node) {
    fail(
      'invalid-document-structure',
      'Worksheet drawing root is missing',
      part,
    );
  }
  const pieces = entry[0].split(':');
  const prefix = pieces.length === 1 ? '' : pieces[0]!;
  const namespace = attributes(node)[prefix ? `xmlns:${prefix}` : 'xmlns'];
  if (namespace !== expectedNamespace) {
    fail(
      'invalid-document-structure',
      'Worksheet drawing root has the wrong namespace',
      part,
    );
  }
  return node;
}

export function parseXlsxDrawingInteger(
  value: unknown,
  signed: boolean,
  message: string,
  part: string,
): number {
  const parsed = Number(value);
  const canonical = String(parsed) === value || (signed && value === '-0');
  if (!canonical || !Number.isSafeInteger(parsed) || (!signed && parsed < 0)) {
    fail('invalid-document-value', message, part);
  }
  return parsed;
}

const integer = parseXlsxDrawingInteger;

function emuPoints(
  value: unknown,
  signed: boolean,
  message: string,
  part: string,
): number {
  return integer(value, signed, message, part) / EMU_PER_POINT;
}

function booleanAttribute(
  value: string | undefined,
  message: string,
  part: string,
): boolean {
  if (value === undefined || value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  fail('invalid-document-value', message, part);
}

function marker(
  value: unknown,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxDrawingMarker {
  const node = record(value);
  if (!node) {
    fail(
      'invalid-document-structure',
      'Drawing anchor marker is invalid',
      part,
    );
  }
  const column = integer(
    scalarText(childByLocal(node, 'col')),
    false,
    'Drawing anchor column is invalid',
    part,
  );
  const row = integer(
    scalarText(childByLocal(node, 'row')),
    false,
    'Drawing anchor row is invalid',
    part,
  );
  if (column >= limits.maxColumnsPerWorksheet) {
    throw new XlsxResourceLimitError(
      'maxColumnsPerWorksheet',
      column + 1,
      limits.maxColumnsPerWorksheet,
      part,
    );
  }
  if (row >= limits.maxRowsPerWorksheet) {
    throw new XlsxResourceLimitError(
      'maxRowsPerWorksheet',
      row + 1,
      limits.maxRowsPerWorksheet,
      part,
    );
  }
  return {
    column: column + 1,
    columnOffset: emuPoints(
      scalarText(childByLocal(node, 'colOff')),
      false,
      'Drawing anchor column offset is invalid',
      part,
    ),
    row: row + 1,
    rowOffset: emuPoints(
      scalarText(childByLocal(node, 'rowOff')),
      false,
      'Drawing anchor row offset is invalid',
      part,
    ),
  };
}

function extent(value: unknown, part: string): XlsxDrawingExtent {
  const node = record(value);
  if (!node)
    fail('invalid-document-structure', 'Drawing extent is invalid', part);
  const attrs = attributes(node);
  return {
    height: emuPoints(attrs.cy, false, 'Drawing height is invalid', part),
    width: emuPoints(attrs.cx, false, 'Drawing width is invalid', part),
  };
}

function crop(value: unknown, part: string): XlsxImageCrop {
  const attrs = record(value) ? attributes(value as XmlRecord) : {};
  const side = (name: string) => {
    const raw = attrs[name];
    if (raw === undefined) return 0;
    const result =
      integer(raw, true, `Image crop ${name} is invalid`, part) / 1000;
    if (result < -100 || result > 100) {
      fail('invalid-document-value', `Image crop ${name} is invalid`, part);
    }
    return result;
  };
  const output = {
    bottom: side('b'),
    left: side('l'),
    right: side('r'),
    top: side('t'),
  };
  if (output.left + output.right >= 100 || output.top + output.bottom >= 100) {
    fail(
      'invalid-document-value',
      'Image crop removes the complete image',
      part,
    );
  }
  return output;
}

function pictureTransform(
  pic: XmlRecord,
  part: string,
): {
  extent: XlsxDrawingExtent;
  transform: XlsxEmbeddedImage['transform'];
} {
  const properties = record(childByLocal(pic, 'spPr'));
  const xfrm = properties
    ? record(childByLocal(properties, 'xfrm'))
    : undefined;
  if (!xfrm) {
    fail('invalid-document-structure', 'Image transform is missing', part);
  }
  const attrs = attributes(xfrm);
  const rotation =
    attrs.rot === undefined
      ? 0
      : integer(attrs.rot, true, 'Image rotation is invalid', part) / 60_000;
  return {
    extent: extent(childByLocal(xfrm, 'ext'), part),
    transform: {
      flipHorizontal: booleanAttribute(
        attrs.flipH,
        'Image horizontal-flip flag is invalid',
        part,
      ),
      flipVertical: booleanAttribute(
        attrs.flipV,
        'Image vertical-flip flag is invalid',
        part,
      ),
      rotation,
    },
  };
}

function anchorSelection(
  kind: XlsxDrawing['kind'],
  from: XlsxDrawingMarker | undefined,
  to: XlsxDrawingMarker | undefined,
  selection: IncludedSelection,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxDrawing['selectionRelation'] | null {
  if (selection.kind === 'full-sheet') return 'full-sheet';
  if (kind === 'absolute') return 'worksheet-global';
  const start = from!;
  const end = to ?? from!;
  for (const range of selection.ranges) {
    consumeXlsxWorksheetBudget(
      budget,
      'scannedCells',
      1,
      'maxScannedCells',
      limits,
      part,
    );
    if (
      start.row <= range.end.row &&
      end.row >= range.start.row &&
      start.column <= range.end.column &&
      end.column >= range.start.column
    ) {
      return 'intersects-selection';
    }
  }
  return null;
}

function consumeDrawing(
  budget: XlsxDrawingBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): void {
  const actual = budget.drawings + 1;
  if (!Number.isSafeInteger(actual) || actual > limits.maxDrawings) {
    throw new XlsxResourceLimitError(
      'maxDrawings',
      actual,
      limits.maxDrawings,
      part,
    );
  }
  budget.drawings = actual;
}

function consumeChart(
  budget: XlsxDrawingBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): void {
  const actual = budget.charts + 1;
  if (!Number.isSafeInteger(actual) || actual > limits.maxCharts) {
    throw new XlsxResourceLimitError(
      'maxCharts',
      actual,
      limits.maxCharts,
      part,
    );
  }
  budget.charts = actual;
}

async function picture(
  pic: XmlRecord,
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  media: XlsxMediaSession,
  part: string,
  includePayload: boolean,
): Promise<{ extent: XlsxDrawingExtent; image: XlsxEmbeddedImage }> {
  const nonVisual = record(childByLocal(pic, 'nvPicPr'));
  const properties = nonVisual
    ? record(childByLocal(nonVisual, 'cNvPr'))
    : undefined;
  if (!properties) {
    fail('invalid-document-structure', 'Image properties are missing', part);
  }
  const attrs = attributes(properties);
  const id = integer(attrs.id, false, 'Image ID is invalid', part);
  if (id === 0) fail('invalid-document-value', 'Image ID is invalid', part);
  if (typeof attrs.name !== 'string' || attrs.name.length === 0) {
    fail('invalid-document-value', 'Image name is invalid', part);
  }
  const hidden = booleanAttribute(
    attrs.hidden,
    'Image hidden flag is invalid',
    part,
  );
  const fill = record(childByLocal(pic, 'blipFill'));
  const blip = fill ? record(childByLocal(fill, 'blip')) : undefined;
  if (!fill || !blip) {
    fail('invalid-document-structure', 'Image fill is missing', part);
  }
  const blipAttrs = attributes(blip);
  const relationshipNamespace = OFFICE_RELATIONSHIPS[discovery.dialect];
  if (blipAttrs['r:link'] !== undefined) {
    fail(
      'security-rejected-content',
      'Externally linked drawing images are not loaded',
      part,
    );
  }
  const relationshipId = blipAttrs['r:embed'];
  if (!relationshipId) {
    fail(
      'invalid-document-value',
      'Image relationship reference is invalid',
      part,
    );
  }
  const relation = relationships.get(relationshipId);
  const imageType = `${relationshipNamespace}/image`;
  if (
    !relation ||
    relation.mode !== 'internal' ||
    relation.type !== imageType
  ) {
    fail('invalid-document-structure', 'Image relationship is invalid', part);
  }
  const contentType = discovery.contentTypes.contentTypeFor(relation.target);
  if (!contentType || !isSafeXlsxImageContentType(contentType)) {
    fail(
      'security-rejected-content',
      'Drawing image content type is not safely supported',
      relation.target,
    );
  }
  const loaded = await media.image(
    relation.target,
    contentType,
    reader,
    includePayload,
  );
  const transformed = pictureTransform(pic, part);
  const description = attrs.descr;
  return {
    extent: transformed.extent,
    image: {
      ...(loaded?.base64 === undefined ? {} : { base64: loaded.base64 }),
      ...(loaded?.blobUrl === undefined ? {} : { blobUrl: loaded.blobUrl }),
      ...(loaded === undefined ? {} : { byteLength: loaded.byteLength }),
      contentType,
      crop: crop(childByLocal(fill, 'srcRect'), part),
      ...(description === undefined ? {} : { description }),
      hidden,
      id,
      kind: 'image',
      name: attrs.name,
      part: relation.target,
      transform: transformed.transform,
    },
  };
}

function objectProperties(
  node: XmlRecord,
  containerName: string,
  part: string,
): {
  description?: string;
  hidden: boolean;
  id: number;
  name: string;
} {
  const container = record(childByLocal(node, containerName));
  const properties = container
    ? record(childByLocal(container, 'cNvPr'))
    : undefined;
  if (!properties) {
    fail(
      'invalid-document-structure',
      'Drawing object properties are missing',
      part,
    );
  }
  const attrs = attributes(properties);
  const id = integer(attrs.id, false, 'Drawing object ID is invalid', part);
  if (id === 0)
    fail('invalid-document-value', 'Drawing object ID is invalid', part);
  if (!attrs.name)
    fail('invalid-document-value', 'Drawing object name is invalid', part);
  return {
    ...(attrs.descr === undefined ? {} : { description: attrs.descr }),
    hidden: booleanAttribute(
      attrs.hidden,
      'Drawing object hidden flag is invalid',
      part,
    ),
    id,
    name: attrs.name,
  };
}

function drawingColor(node: XmlRecord, part: string): XlsxDrawingColor {
  const colorChildren = { ...node };
  delete colorChildren.attrs;
  delete colorChildren.value;
  const entries = Object.entries(colorChildren);
  if (entries.length !== 1) {
    fail('invalid-document-structure', 'Drawing color is invalid', part);
  }
  const [name, colorValue] = entries[0]!;
  const attrs = attributes(record(colorValue) ?? {});
  if (
    localName(name) === 'srgbClr' &&
    typeof attrs.val === 'string' &&
    /^[0-9A-Fa-f]{6}$/u.test(attrs.val)
  ) {
    return { kind: 'rgb', value: attrs.val.toUpperCase() };
  }
  if (localName(name) === 'schemeClr' && attrs.val) {
    return { kind: 'scheme', value: attrs.val };
  }
  if (localName(name) === 'sysClr' && attrs.val) {
    if (
      attrs.lastClr !== undefined &&
      !/^[0-9A-Fa-f]{6}$/u.test(attrs.lastClr)
    ) {
      fail('invalid-document-value', 'Drawing color is invalid', part);
    }
    return {
      kind: 'system',
      ...(attrs.lastClr === undefined
        ? {}
        : { lastColor: attrs.lastClr.toUpperCase() }),
      value: attrs.val,
    };
  }
  fail('invalid-document-value', 'Drawing color is invalid', part);
}

function drawingFill(
  node: XmlRecord,
  part: string,
): XlsxDrawingFill | undefined {
  if (childByLocal(node, 'noFill') !== undefined) return { kind: 'none' };
  const solidValue = childByLocal(node, 'solidFill');
  const solid = record(solidValue);
  if (solidValue !== undefined && !solid) {
    fail('invalid-document-structure', 'Drawing color is invalid', part);
  }
  return solid
    ? { color: drawingColor(solid, part), kind: 'solid' }
    : undefined;
}

function drawingLine(
  node: XmlRecord,
  part: string,
): XlsxDrawingLine | undefined {
  const line = record(childByLocal(node, 'ln'));
  if (!line) return undefined;
  const attrs = attributes(line);
  const presetDash = record(childByLocal(line, 'prstDash'));
  const dash = presetDash ? attributes(presetDash).val : undefined;
  const fill = drawingFill(line, part);
  return {
    ...(dash === undefined ? {} : { dash }),
    ...(fill === undefined ? {} : { fill }),
    ...(attrs.w === undefined
      ? {}
      : {
          width: emuPoints(
            attrs.w,
            false,
            'Drawing line width is invalid',
            part,
          ),
        }),
  };
}

function objectTransform(
  value: unknown,
  part: string,
): XlsxDrawingObjectTransform {
  const xfrm = record(value);
  if (!xfrm)
    fail(
      'invalid-document-structure',
      'Drawing object transform is missing',
      part,
    );
  const off = record(childByLocal(xfrm, 'off'));
  const size = record(childByLocal(xfrm, 'ext'));
  if (!off || !size) {
    fail(
      'invalid-document-structure',
      'Drawing object transform is incomplete',
      part,
    );
  }
  const attrs = attributes(xfrm);
  const offAttrs = attributes(off);
  const extentValue = extent(size, part);
  return {
    flipHorizontal: booleanAttribute(
      attrs.flipH,
      'Drawing object horizontal-flip flag is invalid',
      part,
    ),
    flipVertical: booleanAttribute(
      attrs.flipV,
      'Drawing object vertical-flip flag is invalid',
      part,
    ),
    height: extentValue.height,
    rotation:
      attrs.rot === undefined
        ? 0
        : integer(attrs.rot, true, 'Drawing object rotation is invalid', part) /
          60_000,
    width: extentValue.width,
    x: emuPoints(
      offAttrs.x,
      true,
      'Drawing object X position is invalid',
      part,
    ),
    y: emuPoints(
      offAttrs.y,
      true,
      'Drawing object Y position is invalid',
      part,
    ),
  };
}

function shapeGeometry(
  node: XmlRecord,
  part: string,
): XlsxDrawingShape['geometry'] {
  const preset = record(childByLocal(node, 'prstGeom'));
  if (preset) {
    const value = attributes(preset).prst;
    if (!value)
      fail(
        'invalid-document-value',
        'Drawing preset geometry is invalid',
        part,
      );
    return { kind: 'preset', preset: value };
  }
  if (childByLocal(node, 'custGeom') !== undefined) return { kind: 'custom' };
  fail('invalid-document-structure', 'Drawing geometry is missing', part);
}

function drawingText(
  node: XmlRecord,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): string | undefined {
  const body = record(childByLocal(node, 'txBody'));
  if (!body) return undefined;
  const paragraphs = records(childByLocal(body, 'p'));
  if (!paragraphs)
    fail('invalid-document-structure', 'Drawing text is invalid', part);
  const text = paragraphs
    .map((paragraph) => {
      const values: string[] = [];
      const visit = (value: XmlRecord): void => {
        const childrenByName = { ...value };
        delete childrenByName.attrs;
        delete childrenByName.value;
        const children = Object.entries(childrenByName)
          .flatMap(([name, childValue]) => {
            const childValues: unknown[] = Array.isArray(childValue)
              ? [...(childValue as unknown[])]
              : [childValue];
            return childValues.map((child) => ({
              child,
              name,
              order: getXmlNodeOrder(child),
            }));
          })
          .sort(
            (left, right) =>
              (left.order ?? Number.MAX_SAFE_INTEGER) -
              (right.order ?? Number.MAX_SAFE_INTEGER),
          );
        for (const child of children) {
          const name = localName(child.name);
          if (name === 't') {
            const scalar = scalarText(child.child);
            if (scalar === undefined) {
              fail(
                'invalid-document-structure',
                'Drawing text is invalid',
                part,
              );
            }
            values.push(scalar);
          } else if (name === 'br') {
            values.push('\n');
          } else {
            const childNode = record(child.child);
            if (childNode) visit(childNode);
          }
        }
      };
      visit(paragraph);
      return values.join('');
    })
    .join('\n');
  consumeXlsxWorksheetBudget(
    budget,
    'textCharacters',
    text.length,
    'maxTextCharacters',
    limits,
    part,
  );
  return text;
}

function connection(
  value: unknown,
  part: string,
): { shapeId: number; site: number } | undefined {
  const node = record(value);
  if (!node) return undefined;
  const attrs = attributes(node);
  return {
    shapeId: integer(
      attrs.id,
      false,
      'Drawing connector shape reference is invalid',
      part,
    ),
    site: integer(attrs.idx, false, 'Drawing connector site is invalid', part),
  };
}

function shapeObject(
  node: XmlRecord,
  connector: boolean,
  budget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): XlsxDrawingShape | XlsxDrawingConnector {
  const props = objectProperties(
    node,
    connector ? 'nvCxnSpPr' : 'nvSpPr',
    part,
  );
  const shapeProperties = record(childByLocal(node, 'spPr'));
  if (!shapeProperties) {
    fail(
      'invalid-document-structure',
      'Drawing shape properties are missing',
      part,
    );
  }
  const fill = drawingFill(shapeProperties, part);
  const line = drawingLine(shapeProperties, part);
  const text = drawingText(node, budget, limits, part);
  const base = {
    ...props,
    ...(fill === undefined ? {} : { fill }),
    geometry: shapeGeometry(shapeProperties, part),
    ...(line === undefined ? {} : { line }),
    ...(text === undefined ? {} : { text }),
    transform: objectTransform(childByLocal(shapeProperties, 'xfrm'), part),
  };
  if (!connector) return { ...base, kind: 'shape' };
  const nonVisual = record(childByLocal(node, 'nvCxnSpPr'))!;
  const connectionProperties = record(childByLocal(nonVisual, 'cNvCxnSpPr'));
  const endConnection = connection(
    childByLocal(connectionProperties ?? {}, 'endCxn'),
    part,
  );
  const startConnection = connection(
    childByLocal(connectionProperties ?? {}, 'stCxn'),
    part,
  );
  return {
    ...base,
    ...(endConnection === undefined ? {} : { endConnection }),
    kind: 'connector',
    ...(startConnection === undefined ? {} : { startConnection }),
  };
}

async function chartObject(
  frame: XmlRecord,
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  drawingBudget: XlsxDrawingBudget,
  worksheetBudget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  part: string,
): Promise<XlsxChart> {
  const props = objectProperties(frame, 'nvGraphicFramePr', part);
  const transform = objectTransform(childByLocal(frame, 'xfrm'), part);
  const graphic = record(childByLocal(frame, 'graphic'));
  const graphicData = graphic
    ? record(childByLocal(graphic, 'graphicData'))
    : undefined;
  if (!graphicData) {
    fail('invalid-document-structure', 'Chart graphic data is missing', part);
  }
  const chartNamespace =
    discovery.dialect === 'strict'
      ? 'http://purl.oclc.org/ooxml/drawingml/chart'
      : 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  if (attributes(graphicData).uri !== chartNamespace) {
    fail('invalid-document-value', 'Chart graphic data URI is invalid', part);
  }
  const chart = record(childByLocal(graphicData, 'chart'));
  if (!chart) {
    fail('invalid-document-structure', 'Chart reference is missing', part);
  }
  const relationshipId = attributes(chart)['r:id'];
  if (!relationshipId) {
    fail(
      'invalid-document-value',
      'Chart relationship reference is invalid',
      part,
    );
  }
  const relationshipNamespace = OFFICE_RELATIONSHIPS[discovery.dialect];
  const relation = relationships.get(relationshipId);
  if (relation?.mode === 'external') {
    fail(
      'security-rejected-content',
      'Externally linked charts are not loaded',
      part,
    );
  }
  if (!relation || relation.type !== `${relationshipNamespace}/chart`) {
    fail('invalid-document-structure', 'Chart relationship is invalid', part);
  }
  if (
    discovery.contentTypes.contentTypeFor(relation.target) !==
    CHART_CONTENT_TYPE
  ) {
    fail(
      'invalid-document-structure',
      'Chart target has the wrong content type',
      relation.target,
    );
  }
  consumeChart(drawingBudget, limits, relation.target);
  const loaded = await loadXlsxChart(
    relation.target,
    discovery,
    reader,
    worksheetBudget,
    limits,
  );
  return {
    ...loaded,
    ...props,
    kind: 'chart',
    part: relation.target,
    transform,
  };
}

function registerObjectId(id: number, ids: Set<number>, part: string): void {
  if (ids.has(id)) {
    fail(
      'invalid-document-value',
      'Worksheet drawing contains duplicate object IDs',
      part,
    );
  }
  ids.add(id);
}

async function drawingObject(
  node: XmlRecord,
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  media: XlsxMediaSession,
  includePayload: boolean,
  drawingBudget: XlsxDrawingBudget,
  worksheetBudget: XlsxWorksheetBudget,
  limits: ResolvedXlsxResourceLimits,
  ids: Set<number>,
  part: string,
): Promise<
  { extent: XlsxDrawingExtent; object: XlsxDrawingObject } | undefined
> {
  const candidates = DRAWING_OBJECT_NAMES.flatMap((name) => {
    const value = childByLocal(node, name);
    if (value === undefined) return [];
    const nodes = records(value);
    if (!nodes) {
      fail(
        'invalid-document-structure',
        'Drawing anchor objects are invalid',
        part,
      );
    }
    return nodes.map((candidate) => ({ name, node: candidate }));
  });
  if (candidates.length === 0) return undefined;
  if (candidates.length !== 1) {
    fail(
      'invalid-document-structure',
      'Drawing anchor has multiple objects',
      part,
    );
  }
  const candidate = candidates[0]!;
  if (candidate.name === 'pic') {
    const parsed = await picture(
      candidate.node,
      relationships,
      discovery,
      reader,
      media,
      part,
      includePayload,
    );
    registerObjectId(parsed.image.id, ids, part);
    return { extent: parsed.extent, object: parsed.image };
  }
  if (candidate.name === 'sp' || candidate.name === 'cxnSp') {
    const parsed = shapeObject(
      candidate.node,
      candidate.name === 'cxnSp',
      worksheetBudget,
      limits,
      part,
    );
    registerObjectId(parsed.id, ids, part);
    return {
      extent: {
        height: parsed.transform.height,
        width: parsed.transform.width,
      },
      object: parsed,
    };
  }
  if (candidate.name === 'graphicFrame') {
    const parsed = await chartObject(
      candidate.node,
      relationships,
      discovery,
      reader,
      drawingBudget,
      worksheetBudget,
      limits,
      part,
    );
    registerObjectId(parsed.id, ids, part);
    return {
      extent: {
        height: parsed.transform.height,
        width: parsed.transform.width,
      },
      object: parsed,
    };
  }
  const props = objectProperties(candidate.node, 'nvGrpSpPr', part);
  registerObjectId(props.id, ids, part);
  const groupProperties = record(childByLocal(candidate.node, 'grpSpPr'));
  const xfrm = groupProperties
    ? record(childByLocal(groupProperties, 'xfrm'))
    : undefined;
  if (!xfrm) {
    fail(
      'invalid-document-structure',
      'Drawing group transform is missing',
      part,
    );
  }
  const baseTransform = objectTransform(xfrm, part);
  const childOffset = record(childByLocal(xfrm, 'chOff'));
  const childExtent = record(childByLocal(xfrm, 'chExt'));
  if (!childOffset || !childExtent) {
    fail(
      'invalid-document-structure',
      'Drawing group child transform is missing',
      part,
    );
  }
  const childOffsetAttrs = attributes(childOffset);
  const childSize = extent(childExtent, part);
  const childNodes = Object.entries(candidate.node)
    .flatMap(([name, value]) => {
      if (
        !DRAWING_OBJECT_NAMES.includes(
          localName(name) as (typeof DRAWING_OBJECT_NAMES)[number],
        )
      )
        return [];
      const nodes = records(value);
      if (!nodes) {
        fail(
          'invalid-document-structure',
          'Drawing group children are invalid',
          part,
        );
      }
      return nodes.map((child) => ({
        child,
        name: localName(name),
        order: getXmlNodeOrder(child),
      }));
    })
    .sort(
      (left, right) =>
        (left.order ?? Number.MAX_SAFE_INTEGER) -
        (right.order ?? Number.MAX_SAFE_INTEGER),
    );
  const children: XlsxDrawingObject[] = [];
  for (const child of childNodes) {
    consumeDrawing(drawingBudget, limits, part);
    const parsed = await drawingObject(
      { [child.name]: child.child },
      relationships,
      discovery,
      reader,
      media,
      includePayload,
      drawingBudget,
      worksheetBudget,
      limits,
      ids,
      part,
    );
    if (parsed) children.push(parsed.object);
  }
  const group: XlsxDrawingGroup = {
    children,
    ...props,
    kind: 'group',
    transform: {
      ...baseTransform,
      childHeight: childSize.height,
      childWidth: childSize.width,
      childX: emuPoints(
        childOffsetAttrs.x,
        true,
        'Drawing group child X position is invalid',
        part,
      ),
      childY: emuPoints(
        childOffsetAttrs.y,
        true,
        'Drawing group child Y position is invalid',
        part,
      ),
    },
  };
  return {
    extent: { height: group.transform.height, width: group.transform.width },
    object: group,
  };
}

function anchorKind(name: string): XlsxDrawing['kind'] | undefined {
  if (name === 'absoluteAnchor') return 'absolute';
  if (name === 'oneCellAnchor') return 'one-cell';
  if (name === 'twoCellAnchor') return 'two-cell';
  return undefined;
}

function editAs(
  value: string | undefined,
  part: string,
): XlsxDrawing['editAs'] | undefined {
  if (value === undefined) return undefined;
  if (value === 'absolute') return 'absolute';
  if (value === 'oneCell') return 'one-cell';
  if (value === 'twoCell') return 'two-cell';
  fail('invalid-document-value', 'Drawing anchor edit mode is invalid', part);
}

export async function loadXlsxDrawings(
  drawingRelationshipIds: readonly string[],
  worksheetRelationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  drawingBudget: XlsxDrawingBudget,
  worksheetBudget: XlsxWorksheetBudget,
  selection: IncludedSelection,
  media: XlsxMediaSession,
  worksheetPart: string,
): Promise<XlsxDrawing[]> {
  if (drawingRelationshipIds.length === 0) return [];
  const relation = worksheetRelationships.get(drawingRelationshipIds[0]!);
  const relationshipNamespace = OFFICE_RELATIONSHIPS[discovery.dialect];
  if (
    !relation ||
    relation.mode !== 'internal' ||
    relation.type !== `${relationshipNamespace}/drawing`
  ) {
    fail(
      'invalid-document-structure',
      'Worksheet drawing relationship is invalid',
      worksheetPart,
    );
  }
  if (
    discovery.contentTypes.contentTypeFor(relation.target) !==
    DRAWING_CONTENT_TYPE
  ) {
    fail(
      'invalid-document-structure',
      'Worksheet drawing target has the wrong content type',
      relation.target,
    );
  }
  const xml = await reader.readXml(relation.target, { required: true });
  const drawingRoot = root(
    xml,
    DRAWING_NAMESPACES[discovery.dialect],
    relation.target,
  );
  const relationshipPart = getXlsxRelationshipPartName(relation.target);
  const relationshipsXml = await reader.readXml(relationshipPart);
  const relationships =
    relationshipsXml === null
      ? new Map<string, XlsxRelationship>()
      : parseXlsxRelationships(
          relationshipsXml,
          relation.target,
          limits.maxRelationships,
        );
  const anchors = Object.entries(drawingRoot)
    .flatMap(([name, value]) => {
      const kind = anchorKind(localName(name));
      if (!kind) return [];
      const nodes = records(value);
      if (!nodes) {
        fail(
          'invalid-document-structure',
          'Worksheet drawing anchors are invalid',
          relation.target,
        );
      }
      return nodes.map((node) => ({ kind, node }));
    })
    .sort(
      (left, right) =>
        (getXmlNodeOrder(left.node) ?? Number.MAX_SAFE_INTEGER) -
        (getXmlNodeOrder(right.node) ?? Number.MAX_SAFE_INTEGER),
    );
  const output: XlsxDrawing[] = [];
  const objectIds = new Set<number>();
  for (const anchor of anchors) {
    consumeDrawing(drawingBudget, limits, relation.target);
    const from =
      anchor.kind === 'absolute'
        ? undefined
        : marker(childByLocal(anchor.node, 'from'), limits, relation.target);
    const to =
      anchor.kind === 'two-cell'
        ? marker(childByLocal(anchor.node, 'to'), limits, relation.target)
        : undefined;
    if (
      from &&
      to &&
      (to.row < from.row ||
        to.column < from.column ||
        (to.row === from.row && to.rowOffset < from.rowOffset) ||
        (to.column === from.column && to.columnOffset < from.columnOffset))
    ) {
      fail(
        'invalid-document-value',
        'Two-cell drawing anchor is reversed',
        relation.target,
      );
    }
    const selected = anchorSelection(
      anchor.kind,
      from,
      to,
      selection,
      worksheetBudget,
      limits,
      relation.target,
    );
    const parsedObject = await drawingObject(
      anchor.node,
      relationships,
      discovery,
      reader,
      media,
      selected !== null,
      drawingBudget,
      worksheetBudget,
      limits,
      objectIds,
      relation.target,
    );
    if (!parsedObject) continue;
    if (selected === null) continue;
    const outerExtent =
      anchor.kind === 'two-cell'
        ? parsedObject.extent
        : extent(childByLocal(anchor.node, 'ext'), relation.target);
    const positionNode = record(childByLocal(anchor.node, 'pos'));
    if (anchor.kind === 'absolute' && !positionNode) {
      fail(
        'invalid-document-structure',
        'Absolute drawing position is missing',
        relation.target,
      );
    }
    if (anchor.kind !== 'absolute' && positionNode) {
      fail(
        'invalid-document-structure',
        'Cell drawing anchor contains an absolute position',
        relation.target,
      );
    }
    const positionAttrs = positionNode ? attributes(positionNode) : undefined;
    const rawEditAs = attributes(anchor.node).editAs;
    if (anchor.kind !== 'two-cell' && rawEditAs !== undefined) {
      fail(
        'invalid-document-value',
        'Drawing anchor edit mode is invalid',
        relation.target,
      );
    }
    const anchorEditAs = editAs(rawEditAs, relation.target);
    output.push({
      ...(anchorEditAs === undefined ? {} : { editAs: anchorEditAs }),
      extent: outerExtent,
      ...(from === undefined ? {} : { from }),
      kind: anchor.kind,
      object: parsedObject.object,
      ...(positionAttrs === undefined
        ? {}
        : {
            position: {
              x: emuPoints(
                positionAttrs.x,
                true,
                'Drawing X position is invalid',
                relation.target,
              ),
              y: emuPoints(
                positionAttrs.y,
                true,
                'Drawing Y position is invalid',
                relation.target,
              ),
            },
          }),
      selectionRelation: selected,
      ...(to === undefined ? {} : { to }),
    });
  }
  return output;
}
