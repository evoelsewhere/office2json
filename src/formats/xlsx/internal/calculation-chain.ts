import type { XmlLookupValue } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type { XlsxCalculationChainEntry } from '../types';
import { parseXlsxCellReference } from './cell-reference';
import type { XlsxPartReader } from './part-reader';
import type { XlsxRelationship } from './relationships';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import {
  type XlsxWorkbookDiscovery,
  XLSX_SPREADSHEET_NAMESPACES,
} from './workbook-discovery';

type XmlRecord = Record<string, unknown>;

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
  for (const value of values) {
    const valueRecord = record(value);
    if (!valueRecord) return undefined;
    output.push(valueRecord);
  }
  return output;
}

function attributes(node: XmlRecord): Record<string, string> {
  return (record(node.attrs) ?? {}) as Record<string, string>;
}

function localName(name: string): string {
  return name.slice(name.lastIndexOf(':') + 1);
}

function sourcePrefix(name: string): string {
  const pieces = name.split(':', 2);
  if (pieces.length === 1) return '';
  return pieces[0]!.startsWith('ns_') ? pieces[0]!.slice(3) : pieces[0]!;
}

function namespaceFor(
  qualifiedName: string,
  node: XmlRecord,
  inherited: Readonly<Record<string, string>>,
): string | undefined {
  const prefix = sourcePrefix(qualifiedName);
  const declaration = prefix ? `xmlns:${prefix}` : 'xmlns';
  return attributes(node)[declaration] ?? inherited[declaration];
}

function unsignedInteger(
  value: unknown,
  message: string,
  part: string,
): number {
  if (typeof value !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    fail('invalid-document-value', message, part);
  }
  const parsed = Number(value);
  if (parsed > 0xffff_ffff) fail('invalid-document-value', message, part);
  return parsed;
}

function boolean(value: unknown, message: string, part: string): boolean {
  if (value === undefined || value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  fail('invalid-document-value', message, part);
}

export function parseXlsxCalculationChainPart(
  value: XmlLookupValue,
  dialect: XlsxWorkbookDiscovery['dialect'],
  part: string,
  sheetIdIndexes: ReadonlyMap<number, number>,
  limits: ResolvedXlsxResourceLimits,
): XlsxCalculationChainEntry[] {
  const namespace = XLSX_SPREADSHEET_NAMESPACES[dialect];
  const roots = Object.entries(value).filter(
    ([name]) => localName(name) === 'calcChain',
  );
  if (roots.length !== 1) {
    fail(
      'invalid-document-structure',
      'Calculation chain root is missing or duplicated',
      part,
    );
  }
  const [qualifiedName, rawRoot] = roots[0]!;
  const root = record(rawRoot);
  if (
    !root ||
    namespaceFor(qualifiedName, root, attributes(root)) !== namespace
  ) {
    fail(
      'invalid-document-structure',
      'Calculation chain root has the wrong namespace',
      part,
    );
  }
  const nodes: XmlRecord[] = [];
  for (const [childName, rawNodes] of Object.entries(root)) {
    if (localName(childName) !== 'c') continue;
    const childNodes = records(rawNodes);
    if (!childNodes) {
      fail(
        'invalid-document-structure',
        'Calculation chain entry collection is invalid',
        part,
      );
    }
    for (const node of childNodes) {
      if (namespaceFor(childName, node, attributes(root)) !== namespace) {
        fail(
          'invalid-document-structure',
          'Calculation chain entry has the wrong namespace',
          part,
        );
      }
      nodes.push(node);
    }
  }
  if (nodes.length > limits.maxCalculationChainEntries) {
    throw new XlsxResourceLimitError(
      'maxCalculationChainEntries',
      nodes.length,
      limits.maxCalculationChainEntries,
      part,
    );
  }
  const seen = new Set<string>();
  const output: XlsxCalculationChainEntry[] = [];
  let sheetIndex: number | undefined;
  for (const node of nodes) {
    const attrs = attributes(node);
    if (attrs.i !== undefined) {
      const sheetId = unsignedInteger(
        attrs.i,
        'Calculation chain sheet reference is invalid',
        part,
      );
      sheetIndex = sheetIdIndexes.get(sheetId);
      if (sheetIndex === undefined) {
        fail(
          'invalid-document-value',
          'Calculation chain sheet reference is invalid',
          part,
        );
      }
    } else if (sheetIndex === undefined) {
      fail(
        'invalid-document-value',
        'Calculation chain sheet reference is invalid',
        part,
      );
    }
    const cell = parseXlsxCellReference(attrs.r);
    if (!cell) {
      fail(
        'invalid-document-value',
        'Calculation chain cell reference is invalid',
        part,
      );
    }
    const key = `${sheetIndex}:${cell.address}`;
    if (seen.has(key)) {
      fail(
        'invalid-document-structure',
        'Calculation chain contains a duplicate cell',
        part,
      );
    }
    seen.add(key);
    output.push({
      address: cell.address,
      arrayFormula: boolean(
        attrs.a,
        'Calculation chain array flag is invalid',
        part,
      ),
      childChain: boolean(
        attrs.s,
        'Calculation chain child flag is invalid',
        part,
      ),
      newDependencyLevel: boolean(
        attrs.l,
        'Calculation chain dependency-level flag is invalid',
        part,
      ),
      newThread: boolean(
        attrs.t,
        'Calculation chain thread flag is invalid',
        part,
      ),
      sheetIndex,
    });
  }
  return output;
}

export async function loadXlsxCalculationChain(
  relationships: ReadonlyMap<string, XlsxRelationship>,
  discovery: XlsxWorkbookDiscovery,
  reader: XlsxPartReader,
  sheetIdIndexes: ReadonlyMap<number, number>,
  limits: ResolvedXlsxResourceLimits,
): Promise<XlsxCalculationChainEntry[]> {
  const base =
    discovery.dialect === 'strict'
      ? 'http://purl.oclc.org/ooxml/officeDocument/relationships'
      : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const matches = [...relationships.values()].filter(
    (relationship) => relationship.type === `${base}/calcChain`,
  );
  if (matches.length === 0) return [];
  if (matches.length !== 1) {
    fail(
      'invalid-document-structure',
      'Calculation chain relationship is duplicated',
      discovery.part,
    );
  }
  const relationship = matches[0]!;
  if (relationship.mode !== 'internal') {
    fail(
      'invalid-document-structure',
      'Calculation chain relationship must be internal',
      discovery.part,
    );
  }
  if (
    discovery.contentTypes.contentTypeFor(relationship.target) !==
    'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'
  ) {
    fail(
      'invalid-document-structure',
      'Calculation chain target has the wrong content type',
      relationship.target,
    );
  }
  return parseXlsxCalculationChainPart(
    await reader.readXml(relationship.target, { required: true }),
    discovery.dialect,
    relationship.target,
    sheetIdIndexes,
    limits,
  );
}
