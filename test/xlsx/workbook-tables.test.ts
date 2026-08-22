import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { XlsxPartReader } from '../../src/formats/xlsx/internal/part-reader';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
} from '../../src/formats/xlsx/internal/resource-limits';
import { discoverXlsxWorkbook } from '../../src/formats/xlsx/internal/workbook-discovery';
import { loadXlsxSharedStrings } from '../../src/formats/xlsx/internal/workbook-tables';
import {
  createIndependentXlsx,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
  type XlsxBlackBoxOverrides,
} from '../black-box/xlsx-package';

const STRICT_OFFICE_REL_NS =
  'http://purl.oclc.org/ooxml/officeDocument/relationships';
const STRICT_SPREADSHEET_NS = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const SHARED_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';

function relationships(entries: string): string {
  return `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">${entries}</Relationships>`;
}

function relationship(
  id: string,
  type: string,
  target: string,
  mode?: 'External',
): string {
  return `<Relationship Id="${id}" Type="${type}" Target="${target}"${
    mode === undefined ? '' : ` TargetMode="${mode}"`
  }/>`;
}

function contentTypes(sharedOverride: string): string {
  return `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    ${sharedOverride}
  </Types>`;
}

async function load(
  overrides: XlsxBlackBoxOverrides = {},
  limitOverrides: Partial<ResolvedXlsxResourceLimits> = {},
) {
  const zip = await JSZip.loadAsync(await createIndependentXlsx(overrides));
  const limits = { ...defaultXlsxResourceLimits(), ...limitOverrides };
  const reader = new XlsxPartReader(zip, [], limits);
  const discovery = await discoverXlsxWorkbook(reader, limits);
  return loadXlsxSharedStrings(discovery, reader, limits);
}

async function captureError(
  overrides: XlsxBlackBoxOverrides,
): Promise<XlsxParseError> {
  try {
    await load(overrides);
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected workbook table loading to fail');
}

describe('XLSX workbook tables', () => {
  it('loads the conventional shared-string relationship', async () => {
    await expect(load()).resolves.toEqual({
      part: 'xl/sharedStrings.xml',
      values: [{ text: 'Black box' }],
    });
  });

  it('returns an immutable empty table when the relationship is absent', async () => {
    const result = await load({
      'xl/_rels/workbook.xml.rels': relationships(
        relationship(
          'sheet',
          `${XLSX_OFFICE_REL_TYPE}worksheet`,
          'worksheets/sheet1.xml',
        ),
      ),
    });

    expect(result).toEqual({ part: null, values: [] });
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.values)).toBe(true);
  });

  it('resolves a relocated shared-string part from its workbook owner', async () => {
    const result = await load({
      '[Content_Types].xml': contentTypes(
        `<Override PartName="/data/strings.xml" ContentType="${SHARED_CONTENT_TYPE}"/>`,
      ),
      'data/strings.xml': `<sst xmlns="${XLSX_SPREADSHEET_NS}"><si><t>Relocated</t></si></sst>`,
      'xl/_rels/workbook.xml.rels': relationships(
        relationship(
          'strings',
          `${XLSX_OFFICE_REL_TYPE}sharedStrings`,
          '../data/strings.xml',
        ),
      ),
      'xl/sharedStrings.xml': null,
    });

    expect(result).toEqual({
      part: 'data/strings.xml',
      values: [{ text: 'Relocated' }],
    });
  });

  it('uses the Strict relationship and part namespaces together', async () => {
    const strictRelationshipType = `${STRICT_OFFICE_REL_NS}/sharedStrings`;
    const result = await load({
      'xl/_rels/workbook.xml.rels': relationships(
        relationship('strings', strictRelationshipType, 'sharedStrings.xml'),
      ),
      'xl/sharedStrings.xml': `<s:sst xmlns:s="${STRICT_SPREADSHEET_NS}"><s:si><s:t>Strict</s:t></s:si></s:sst>`,
      'xl/workbook.xml': `<s:workbook xmlns:s="${STRICT_SPREADSHEET_NS}" xmlns:r="${STRICT_OFFICE_REL_NS}"/>`,
    });

    expect(result.values).toEqual([{ text: 'Strict' }]);
  });

  it('ignores a shared-string relationship from the other dialect', async () => {
    const result = await load({
      'xl/_rels/workbook.xml.rels': relationships(
        relationship(
          'strings',
          `${STRICT_OFFICE_REL_NS}/sharedStrings`,
          'sharedStrings.xml',
        ),
      ),
    });

    expect(result).toEqual({ part: null, values: [] });
  });

  it.each([
    [
      {
        'xl/_rels/workbook.xml.rels': relationships(`
          ${relationship('one', `${XLSX_OFFICE_REL_TYPE}sharedStrings`, 'sharedStrings.xml')}
          ${relationship('two', `${XLSX_OFFICE_REL_TYPE}sharedStrings`, 'other.xml')}`),
      },
      'invalid-document-structure',
      'Workbook contains multiple shared-string relationships',
      'xl/_rels/workbook.xml.rels',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': relationships(
          relationship(
            'strings',
            `${XLSX_OFFICE_REL_TYPE}sharedStrings`,
            'https://example.com/strings.xml',
            'External',
          ),
        ),
      },
      'invalid-relationship-target',
      'Workbook shared-string relationship must be internal',
      'xl/_rels/workbook.xml.rels',
    ],
    [
      {
        '[Content_Types].xml': contentTypes(''),
      },
      'invalid-document-structure',
      'Workbook shared-string target has the wrong content type',
      'xl/sharedStrings.xml',
    ],
  ] as const)(
    'rejects invalid shared-string relationship graph %#',
    async (overrides, code, message, part) => {
      const error = await captureError(overrides);
      expect(error.diagnostic).toMatchObject({
        code,
        message,
        part,
        severity: 'error',
      });
    },
  );

  it('includes relationship type metadata for an external target', async () => {
    const type = `${XLSX_OFFICE_REL_NS}/sharedStrings`;
    const error = await captureError({
      'xl/_rels/workbook.xml.rels': relationships(
        relationship(
          'strings',
          type,
          'https://example.com/strings.xml',
          'External',
        ),
      ),
    });

    expect(error.diagnostic.relationshipType).toBe(type);
  });

  it('requires the target when a valid relationship exists', async () => {
    const error = await captureError({ 'xl/sharedStrings.xml': null });
    expect(error.diagnostic).toMatchObject({
      code: 'missing-required-part',
      part: 'xl/sharedStrings.xml',
    });
  });

  it('requires the workbook relationship part', async () => {
    const error = await captureError({
      'xl/_rels/workbook.xml.rels': null,
    });
    expect(error.diagnostic).toMatchObject({
      code: 'missing-required-part',
      part: 'xl/_rels/workbook.xml.rels',
    });
  });
});
