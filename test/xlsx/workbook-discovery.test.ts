import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { XlsxPartReader } from '../../src/formats/xlsx/internal/part-reader';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import {
  discoverXlsxWorkbook,
  XLSX_SPREADSHEET_NAMESPACES,
} from '../../src/formats/xlsx/internal/workbook-discovery';
import type { XlsxDiagnostic } from '../../src/formats/xlsx/types';
import {
  createIndependentXlsx,
  independentWorkbook,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
} from '../black-box/xlsx-package';

const WORKBOOK_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const STRICT_OFFICE_REL_TYPE =
  'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument';

function contentTypes(part: string, type = WORKBOOK_CONTENT_TYPE): string {
  return `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/${part}" ContentType="${type}"/>
  </Types>`;
}

function packageRelationships(entries: string): string {
  return `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">${entries}</Relationships>`;
}

function officeRelationship(
  target: string,
  options: { id?: string; mode?: string; type?: string } = {},
): string {
  return `<Relationship Id="${options.id ?? 'rIdWorkbook'}" Type="${
    options.type ?? `${XLSX_OFFICE_REL_TYPE}officeDocument`
  }" Target="${target}"${
    options.mode === undefined ? '' : ` TargetMode="${options.mode}"`
  }/>`;
}

async function discover(
  overrides: Record<string, string | Uint8Array | null> = {},
): Promise<{
  diagnostics: XlsxDiagnostic[];
  result: Awaited<ReturnType<typeof discoverXlsxWorkbook>>;
}> {
  const zip = await JSZip.loadAsync(await createIndependentXlsx(overrides));
  const diagnostics: XlsxDiagnostic[] = [];
  const limits = defaultXlsxResourceLimits();
  const result = await discoverXlsxWorkbook(
    new XlsxPartReader(zip, diagnostics, limits),
    limits,
  );
  return { diagnostics, result };
}

async function captureDiscoveryError(
  overrides: Record<string, string | Uint8Array | null>,
): Promise<XlsxParseError> {
  try {
    await discover(overrides);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected workbook discovery to fail');
}

describe('XLSX workbook discovery', () => {
  it('discovers the conventional Transitional workbook', async () => {
    const { diagnostics, result } = await discover();

    expect(result).toMatchObject({
      dialect: 'transitional',
      part: 'xl/workbook.xml',
      root: { workbook: {} },
    });
    expect(diagnostics).toEqual([]);
  });

  it('discovers a relocated prefixed Strict workbook', async () => {
    const part = 'custom/books/main.xml';
    const strictWorkbook = independentWorkbook(
      '<sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/>',
    )
      .replaceAll(
        'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        XLSX_SPREADSHEET_NAMESPACES.strict,
      )
      .replace('<workbook ', '<s:workbook ')
      .replace('</workbook>', '</s:workbook>')
      .replace('xmlns="', 'xmlns:s="');
    const { result } = await discover({
      '[Content_Types].xml': contentTypes(part),
      '_rels/.rels': packageRelationships(
        officeRelationship(`/${part}`, { type: STRICT_OFFICE_REL_TYPE }),
      ),
      'xl/workbook.xml': null,
      [part]: strictWorkbook,
    });

    expect(result.dialect).toBe('strict');
    expect(result.part).toBe(part);
    expect(result.root).toMatchObject({ 's:workbook': {} });
  });

  it.each([
    [
      { '[Content_Types].xml': null },
      'Required XLSX part is missing: [Content_Types].xml',
      'missing-required-part',
      '[Content_Types].xml',
    ],
    [
      { '_rels/.rels': null },
      'Required XLSX part is missing: _rels/.rels',
      'missing-required-part',
      '_rels/.rels',
    ],
    [
      {},
      'Package must contain exactly one office-document relationship',
      'invalid-document-structure',
      '_rels/.rels',
    ],
    [
      {
        '_rels/.rels': packageRelationships(
          `${officeRelationship('xl/workbook.xml')}${officeRelationship(
            'xl/other.xml',
            { id: 'rIdOther' },
          )}`,
        ),
      },
      'Package must contain exactly one office-document relationship',
      'invalid-document-structure',
      '_rels/.rels',
    ],
    [
      {
        '_rels/.rels': packageRelationships(
          officeRelationship('https://example.com/book.xlsx', {
            mode: 'External',
          }),
        ),
      },
      'Package office-document relationship must be internal',
      'invalid-relationship-target',
      '_rels/.rels',
    ],
    ...[
      'application/vnd.ms-excel.addin.macroEnabled.main+xml',
      'application/vnd.ms-excel.sheet.binary.macroEnabled.main',
      'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
      'application/vnd.ms-excel.template.macroEnabled.main+xml',
    ].map(
      (type) =>
        [
          { '[Content_Types].xml': contentTypes('xl/workbook.xml', type) },
          'Macro-enabled or binary spreadsheet main parts are not accepted',
          'security-rejected-content',
          'xl/workbook.xml',
        ] as const,
    ),
    [
      {
        '[Content_Types].xml': contentTypes(
          'xl/workbook.xml',
          'application/xml',
        ),
      },
      'Office-document relationship does not target an XLSX workbook main part',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      { 'xl/workbook.xml': null },
      'Required XLSX part is missing: xl/workbook.xml',
      'missing-required-part',
      'xl/workbook.xml',
    ],
    [
      { 'xl/workbook.xml': '<presentation xmlns="urn:wrong"/>' },
      'Workbook root is missing or has an unsupported namespace',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
    [
      {
        'xl/workbook.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NAMESPACES.transitional}"/>`,
      },
      'Workbook root is missing or has an unsupported namespace',
      'invalid-document-structure',
      'xl/workbook.xml',
    ],
  ] as const)(
    'rejects invalid workbook discovery %#',
    async (overrides, message, code, part) => {
      const effectiveOverrides =
        Object.keys(overrides).length === 0
          ? {
              '_rels/.rels': packageRelationships(
                `<Relationship Id="rIdOther" Type="${XLSX_OFFICE_REL_NS}/metadata" Target="metadata.xml"/>`,
              ),
            }
          : overrides;
      const error = await captureDiscoveryError(effectiveOverrides);

      expect(error.diagnostic).toEqual({
        code,
        message,
        part,
        severity: 'error',
      });
    },
  );
});
