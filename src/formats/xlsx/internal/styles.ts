import type { XmlLookupValue } from '../../../common/xml/tree';
import { decodeXmlEntities } from '../../../common/text/html';
import { XlsxParseError } from '../errors';
import type { XlsxNamedStyle, XlsxStyle } from '../types';
import { xlsxBuiltinNumberFormatCode } from './number-format';
import { getXlsxRelationshipPartName } from './package-identity';
import { XlsxPartReader } from './part-reader';
import { parseXlsxRelationships } from './relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import { parseXlsxStyleBorder } from './style-border';
import { parseXlsxDifferentialStyle } from './style-differential';
import {
  EMPTY_XLSX_FEATURE_PROPERTY_BAGS,
  loadXlsxFeaturePropertyBags,
  xlsxFeaturePropertyBagNamespace,
  type XlsxFeaturePropertyBagRegistry,
} from './feature-property-bag';
import { parseXlsxStyleFont } from './style-font';
import { parseXlsxStyleFill } from './style-fill';
import { parseXlsxXfFormatting } from './style-formatting';
import {
  type XlsxWorkbookDiscovery,
  XLSX_SPREADSHEET_NAMESPACES,
} from './workbook-discovery';

type XmlRecord = Record<string, unknown>;

type StyleCategory = keyof Pick<
  XlsxStyle,
  | 'alignment'
  | 'border'
  | 'checkbox'
  | 'fill'
  | 'font'
  | 'numberFormat'
  | 'protection'
>;

interface XlsxDirectXfStyle {
  present: Readonly<Record<StyleCategory, boolean>>;
  style: XlsxStyle;
}

type XlsxApplyFlags = Readonly<Record<StyleCategory, boolean | undefined>>;

export interface XlsxCellXf {
  normalizedStyle: number;
  numberFormat?: string;
}

export interface XlsxStyleTable {
  cellXfs: readonly XlsxCellXf[];
  differentialStyles: readonly XlsxStyle[];
  namedStyles: readonly XlsxNamedStyle[];
  part: string | null;
  recordCount: number;
  styles: readonly XlsxStyle[];
}

export const EMPTY_XLSX_STYLE_TABLE: XlsxStyleTable = Object.freeze({
  cellXfs: Object.freeze([]),
  differentialStyles: Object.freeze([]),
  namedStyles: Object.freeze([]),
  part: null,
  recordCount: 0,
  styles: Object.freeze([]),
});

function structureFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-structure',
    message,
    part,
    severity: 'error',
  });
}

function valueFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function records(value: unknown): XmlRecord[] | undefined {
  if (value === undefined) return [];
  const items: unknown[] = Array.isArray(value) ? value : [value];
  const output: XmlRecord[] = [];
  for (const item of items) {
    const parsed = record(item);
    if (!parsed) return undefined;
    output.push(parsed);
  }
  return output;
}

function attributes(value: XmlRecord): XmlRecord {
  return record(value.attrs) ?? {};
}

function rootEntry(
  value: XmlLookupValue,
  dialect: XlsxWorkbookDiscovery['dialect'],
  part: string,
): { node: XmlRecord; prefix: string } {
  const document = value as unknown as XmlRecord;
  const entries = Object.entries(document);
  if (entries.length !== 1) {
    structureFailure('Styles root is missing', part);
  }
  const [qualifiedName, sourceNode] = entries[0]!;
  const node = record(sourceNode);
  const [first, second] = qualifiedName.split(':') as [string, string?];
  const prefix = second === undefined ? '' : first;
  if (!node || (second ?? first) !== 'styleSheet') {
    structureFailure('Styles root is missing', part);
  }
  const namespace = attributes(node)[prefix ? `xmlns:${prefix}` : 'xmlns'];
  if (namespace !== XLSX_SPREADSHEET_NAMESPACES[dialect]) {
    structureFailure('Styles root has the wrong namespace', part);
  }
  return { node, prefix };
}

function child(node: XmlRecord, prefix: string, localName: string): unknown {
  return node[prefix ? `${prefix}:${localName}` : localName];
}

function localName(name: string): string {
  return name.slice(name.lastIndexOf(':') + 1);
}

function sourcePrefix(name: string): string {
  const parts = name.split(':', 2);
  if (parts.length === 1) return '';
  return parts[0]!.startsWith('ns_') ? parts[0]!.slice(3) : parts[0]!;
}

function checkboxStyle(
  xf: XmlRecord,
  prefix: string,
  part: string,
  featureBags: XlsxFeaturePropertyBagRegistry,
): boolean {
  const rawList = child(xf, prefix, 'extLst');
  if (rawList === undefined) return false;
  const list = record(rawList);
  if (!list) structureFailure('Styles XF extension list is invalid', part);
  const extensions = records(child(list, prefix, 'ext'));
  if (!extensions)
    structureFailure('Styles XF extension list is invalid', part);
  const matches = extensions.filter(
    (extension) =>
      attributes(extension).uri === '{C7286773-470A-42A8-94C5-96B5CB345126}',
  );
  if (matches.length === 0) return false;
  if (matches.length !== 1) {
    structureFailure('Styles checkbox extension is duplicated', part);
  }
  if (featureBags.part === null) {
    throw new XlsxParseError({
      code: 'missing-required-part',
      message: 'Styles checkbox feature property bag is missing',
      part,
      severity: 'error',
    });
  }
  const values: XmlRecord[] = [];
  for (const [qualifiedName, value] of Object.entries(matches[0]!)) {
    if (localName(qualifiedName) !== 'xfComplement') continue;
    const nodes = Array.isArray(value) ? value : [value];
    for (const node of nodes) {
      const parsed = record(node);
      if (!parsed)
        structureFailure('Styles checkbox extension is invalid', part);
      const source = sourcePrefix(qualifiedName);
      const declaration = source ? `xmlns:${source}` : 'xmlns';
      if (
        (attributes(parsed)[declaration] ??
          attributes(matches[0]!)[declaration]) !==
        xlsxFeaturePropertyBagNamespace()
      ) {
        structureFailure(
          'Styles checkbox extension has the wrong namespace',
          part,
        );
      }
      values.push(parsed);
    }
  }
  if (values.length !== 1) {
    structureFailure('Styles checkbox extension is invalid', part);
  }
  const index = unsignedInteger(
    attributes(values[0]!).i,
    'Styles checkbox feature reference is invalid',
    part,
  );
  if (!featureBags.checkboxComplements.has(index)) {
    valueFailure('Styles checkbox feature reference is invalid', part);
  }
  return true;
}

function unsignedInteger(
  value: unknown,
  message: string,
  part: string,
): number {
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    valueFailure(message, part);
  }
  const parsed = Number(value);
  if (parsed > 0xffff_ffff) {
    valueFailure(message, part);
  }
  return parsed;
}

function optionalBoolean(
  value: unknown,
  message: string,
  part: string,
): boolean | undefined {
  if (value === undefined) return undefined;
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  valueFailure(message, part);
}

function collection(
  root: XmlRecord,
  prefix: string,
  collectionName: string,
  itemName: string,
  part: string,
  required: boolean,
): XmlRecord[] {
  const source = record(child(root, prefix, collectionName));
  if (!source) {
    if (!required) return [];
    structureFailure(`Styles ${collectionName} collection is missing`, part);
  }
  const items = records(child(source, prefix, itemName));
  if (!items) {
    structureFailure(`Styles ${collectionName} collection is invalid`, part);
  }
  const count = unsignedInteger(
    attributes(source).count,
    `Styles ${collectionName} count is invalid`,
    part,
  );
  if (count !== items.length) {
    structureFailure(`Styles ${collectionName} count does not match`, part);
  }
  if (required && items.length === 0) {
    structureFailure(`Styles ${collectionName} collection is empty`, part);
  }
  return items;
}

function customNumberFormats(
  root: XmlRecord,
  prefix: string,
  part: string,
): Map<number, string> {
  const values = new Map<number, string>();
  for (const item of collection(
    root,
    prefix,
    'numFmts',
    'numFmt',
    part,
    false,
  )) {
    const attrs = attributes(item);
    const id = unsignedInteger(
      attrs.numFmtId,
      'Styles number-format ID is invalid',
      part,
    );
    if (id < 164) {
      valueFailure('Styles custom number-format ID is reserved', part);
    }
    if (typeof attrs.formatCode !== 'string' || attrs.formatCode.length === 0) {
      valueFailure('Styles number-format code is invalid', part);
    }
    if (values.has(id)) {
      structureFailure('Styles contain a duplicate number-format ID', part);
    }
    values.set(id, decodeXmlEntities(attrs.formatCode));
  }
  return values;
}

function referencedIndex(
  value: unknown,
  count: number,
  message: string,
  part: string,
): number {
  const index = value === undefined ? 0 : unsignedInteger(value, message, part);
  if (index >= count) valueFailure(message, part);
  return index;
}

function numberFormat(
  id: number,
  custom: ReadonlyMap<number, string>,
  part: string,
): string | undefined {
  if (id >= 164) {
    const code = custom.get(id);
    if (code === undefined) {
      valueFailure('Styles XF references a missing custom number format', part);
    }
    return code;
  }
  const code = xlsxBuiltinNumberFormatCode(id);
  if (code === undefined) {
    throw new XlsxParseError({
      code: 'unsupported-feature',
      message: 'Styles XF uses a locale-dependent built-in number format',
      part,
      severity: 'error',
    });
  }
  return code === 'General' ? undefined : code;
}

function semanticFill(fill: XlsxStyle['fill']): XlsxStyle['fill'] {
  if (
    fill?.kind === 'pattern' &&
    fill.pattern === 'none' &&
    fill.foregroundColor === undefined &&
    fill.backgroundColor === undefined
  ) {
    return undefined;
  }
  return fill;
}

function directXfStyle(
  xf: XmlRecord,
  prefix: string,
  part: string,
  custom: ReadonlyMap<number, string>,
  fonts: readonly NonNullable<XlsxStyle['font']>[],
  fills: readonly NonNullable<XlsxStyle['fill']>[],
  borders: readonly NonNullable<XlsxStyle['border']>[],
  featureBags: XlsxFeaturePropertyBagRegistry,
): XlsxDirectXfStyle {
  const attrs = attributes(xf);
  const numFmtId =
    attrs.numFmtId === undefined
      ? 0
      : unsignedInteger(
          attrs.numFmtId,
          'Styles XF number-format ID is invalid',
          part,
        );
  const fontId = referencedIndex(
    attrs.fontId,
    fonts.length,
    'Styles XF font reference is invalid',
    part,
  );
  const fillId = referencedIndex(
    attrs.fillId,
    fills.length,
    'Styles XF fill reference is invalid',
    part,
  );
  const borderId = referencedIndex(
    attrs.borderId,
    borders.length,
    'Styles XF border reference is invalid',
    part,
  );
  const formatting = parseXlsxXfFormatting(xf, prefix, part);
  const checkbox = checkboxStyle(xf, prefix, part, featureBags);
  const border = borders[borderId]!;
  const fill = semanticFill(fills[fillId]);
  const font = fonts[fontId]!;
  const formatCode = numberFormat(numFmtId, custom, part);
  return {
    present: {
      alignment: child(xf, prefix, 'alignment') !== undefined,
      border: attrs.borderId !== undefined,
      checkbox,
      fill: attrs.fillId !== undefined,
      font: attrs.fontId !== undefined,
      numberFormat: attrs.numFmtId !== undefined,
      protection: child(xf, prefix, 'protection') !== undefined,
    },
    style: {
      ...formatting,
      ...(checkbox ? { checkbox: true } : {}),
      ...(Object.keys(border).length === 0 ? {} : { border }),
      ...(fill === undefined ? {} : { fill }),
      ...(Object.keys(font).length === 0 ? {} : { font }),
      ...(formatCode === undefined ? {} : { numberFormat: formatCode }),
    },
  };
}

function applyFlags(attrs: XmlRecord, part: string): XlsxApplyFlags {
  return {
    alignment: optionalBoolean(
      attrs.applyAlignment,
      'Styles XF applyAlignment flag is invalid',
      part,
    ),
    border: optionalBoolean(
      attrs.applyBorder,
      'Styles XF applyBorder flag is invalid',
      part,
    ),
    checkbox: undefined,
    fill: optionalBoolean(
      attrs.applyFill,
      'Styles XF applyFill flag is invalid',
      part,
    ),
    font: optionalBoolean(
      attrs.applyFont,
      'Styles XF applyFont flag is invalid',
      part,
    ),
    numberFormat: optionalBoolean(
      attrs.applyNumberFormat,
      'Styles XF applyNumberFormat flag is invalid',
      part,
    ),
    protection: optionalBoolean(
      attrs.applyProtection,
      'Styles XF applyProtection flag is invalid',
      part,
    ),
  };
}

function appliedValue<T>(
  base: T | undefined,
  direct: T | undefined,
  apply: boolean | undefined,
  present: boolean,
): T | undefined {
  if (apply === false) return base;
  if (apply === true || present) return direct;
  return base;
}

function resolvedCellStyle(
  base: XlsxStyle,
  direct: XlsxDirectXfStyle,
  flags: XlsxApplyFlags,
): XlsxStyle {
  const style: XlsxStyle = {};
  for (const category of [
    'alignment',
    'border',
    'checkbox',
    'fill',
    'font',
    'numberFormat',
    'protection',
  ] as const) {
    const value = appliedValue(
      base[category],
      direct.style[category],
      flags[category],
      direct.present[category],
    );
    if (value !== undefined) Object.assign(style, { [category]: value });
  }
  return style;
}

function parseNamedStyles(
  values: readonly XmlRecord[],
  baseStyles: readonly XlsxStyle[],
  part: string,
): XlsxNamedStyle[] {
  const names = new Set<string>();
  const output: XlsxNamedStyle[] = [];
  for (const value of values) {
    const attrs = attributes(value);
    if (typeof attrs.name !== 'string' || attrs.name.length === 0) {
      valueFailure('Named style name is invalid', part);
    }
    const normalizedName = attrs.name.toLowerCase();
    if (names.has(normalizedName)) {
      structureFailure('Styles contain a duplicate named-style name', part);
    }
    names.add(normalizedName);
    if (attrs.xfId === undefined) {
      valueFailure('Named style base-style reference is invalid', part);
    }
    const baseIndex = referencedIndex(
      attrs.xfId,
      baseStyles.length,
      'Named style base-style reference is invalid',
      part,
    );
    const builtinId =
      attrs.builtinId === undefined
        ? undefined
        : unsignedInteger(
            attrs.builtinId,
            'Named style builtin ID is invalid',
            part,
          );
    const customBuiltin = optionalBoolean(
      attrs.customBuiltin,
      'Named style customBuiltin flag is invalid',
      part,
    );
    const hidden = optionalBoolean(
      attrs.hidden,
      'Named style hidden flag is invalid',
      part,
    );
    const outlineLevel =
      attrs.iLevel === undefined
        ? undefined
        : unsignedInteger(
            attrs.iLevel,
            'Named style outline level is invalid',
            part,
          );
    output.push(
      Object.freeze({
        ...(builtinId === undefined ? {} : { builtinId }),
        ...(customBuiltin ? { customBuiltin: true } : {}),
        ...(hidden ? { hidden: true } : {}),
        name: attrs.name,
        ...(outlineLevel === undefined ? {} : { outlineLevel }),
        style: baseStyles[baseIndex]!,
      }),
    );
  }
  return output;
}

export function parseXlsxStylePart(
  value: XmlLookupValue,
  dialect: XlsxWorkbookDiscovery['dialect'],
  part: string,
  limits: ResolvedXlsxResourceLimits,
  featureBags: XlsxFeaturePropertyBagRegistry = EMPTY_XLSX_FEATURE_PROPERTY_BAGS,
): XlsxStyleTable {
  const { node: root, prefix } = rootEntry(value, dialect, part);
  const custom = customNumberFormats(root, prefix, part);
  const fonts = collection(root, prefix, 'fonts', 'font', part, true).map(
    (font) => parseXlsxStyleFont(font, prefix, part),
  );
  const fills = collection(root, prefix, 'fills', 'fill', part, true).map(
    (fill) => parseXlsxStyleFill(fill, prefix, part),
  );
  const borders = collection(root, prefix, 'borders', 'border', part, true).map(
    (border) => parseXlsxStyleBorder(border, prefix, part),
  );
  const baseXfs = collection(root, prefix, 'cellStyleXfs', 'xf', part, true);
  const xfs = collection(root, prefix, 'cellXfs', 'xf', part, true);
  const differentialNodes = collection(
    root,
    prefix,
    'dxfs',
    'dxf',
    part,
    false,
  );
  const namedNodes = collection(
    root,
    prefix,
    'cellStyles',
    'cellStyle',
    part,
    false,
  );
  const totalStyles =
    custom.size +
    fonts.length +
    fills.length +
    borders.length +
    baseXfs.length +
    xfs.length +
    differentialNodes.length +
    namedNodes.length;
  const aggregateStyles = totalStyles + featureBags.records;
  if (aggregateStyles > limits.maxStyles) {
    throw new XlsxResourceLimitError(
      'maxStyles',
      aggregateStyles,
      limits.maxStyles,
      part,
    );
  }

  const styles: XlsxStyle[] = [];
  const cellXfs: XlsxCellXf[] = [];
  const normalizedStyles = new Map<string, number>();
  const baseStyles = baseXfs.map((xf) => {
    applyFlags(attributes(xf), part);
    return Object.freeze(
      directXfStyle(
        xf,
        prefix,
        part,
        custom,
        fonts,
        fills,
        borders,
        featureBags,
      ).style,
    );
  });
  const differentialStyles = differentialNodes.map((dxf) =>
    parseXlsxDifferentialStyle(dxf, prefix, part),
  );
  const namedStyles = parseNamedStyles(namedNodes, baseStyles, part);
  for (const xf of xfs) {
    const attrs = attributes(xf);
    const baseIndex = referencedIndex(
      attrs.xfId,
      baseStyles.length,
      'Styles XF base-style reference is invalid',
      part,
    );
    const direct = directXfStyle(
      xf,
      prefix,
      part,
      custom,
      fonts,
      fills,
      borders,
      featureBags,
    );
    const style = resolvedCellStyle(
      baseStyles[baseIndex]!,
      direct,
      applyFlags(attrs, part),
    );
    const styleKey = JSON.stringify(style);
    let normalizedStyle = normalizedStyles.get(styleKey);
    if (normalizedStyle === undefined) {
      normalizedStyle = styles.length;
      normalizedStyles.set(styleKey, normalizedStyle);
      styles.push(Object.freeze(style));
    }
    cellXfs.push(
      Object.freeze({
        normalizedStyle,
        ...(style.numberFormat === undefined
          ? {}
          : { numberFormat: style.numberFormat }),
      }),
    );
  }
  return Object.freeze({
    cellXfs: Object.freeze(cellXfs),
    differentialStyles: Object.freeze(differentialStyles),
    namedStyles: Object.freeze(namedStyles),
    part,
    recordCount: aggregateStyles,
    styles: Object.freeze(styles),
  });
}

function stylesContentType(): string {
  return 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
}

function stylesRelationshipType(
  dialect: XlsxWorkbookDiscovery['dialect'],
): string {
  return dialect === 'strict'
    ? 'http://purl.oclc.org/ooxml/officeDocument/relationships/styles'
    : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
}

export async function loadXlsxStyles(
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
): Promise<XlsxStyleTable> {
  const relationshipPart = getXlsxRelationshipPartName(discovery.part);
  const relationshipXml = await reader.readXml(relationshipPart, {
    required: true,
  });
  const relationships = parseXlsxRelationships(
    relationshipXml,
    discovery.part,
    limits.maxRelationships,
  );
  const relationshipType = stylesRelationshipType(discovery.dialect);
  const candidates = [...relationships.values()].filter(
    (relationship) => relationship.type === relationshipType,
  );
  const featureBags = await loadXlsxFeaturePropertyBags(
    relationships,
    discovery,
    reader,
    limits,
  );
  if (candidates.length === 0) return EMPTY_XLSX_STYLE_TABLE;
  if (candidates.length !== 1) {
    structureFailure(
      'Workbook contains multiple styles relationships',
      relationshipPart,
    );
  }
  const relationship = candidates[0]!;
  if (relationship.mode !== 'internal') {
    throw new XlsxParseError({
      code: 'invalid-relationship-target',
      message: 'Workbook styles relationship must be internal',
      part: relationshipPart,
      relationshipType: relationship.type,
      severity: 'error',
    });
  }
  if (
    discovery.contentTypes.contentTypeFor(relationship.target) !==
    stylesContentType()
  ) {
    structureFailure(
      'Workbook styles target has the wrong content type',
      relationship.target,
    );
  }
  const value = await reader.readXml(relationship.target, { required: true });
  return parseXlsxStylePart(
    value,
    discovery.dialect,
    relationship.target,
    limits,
    featureBags,
  );
}
