import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  parseXlsxWithDiagnostics,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
  XlsxParseError,
} from '../../src/formats/xlsx';
import {
  parseXlsxExternalBoolean,
  parseXlsxExternalUnsignedInteger,
  redactXlsxExternalTarget,
  xlsxConnectionCredentialsOmitted,
  xlsxConnectionSourceDataOmitted,
} from '../../src/formats/xlsx/internal/external-metadata';
import {
  createIndependentXlsx,
  type XlsxBlackBoxOverrides,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const EXTERNAL_LINK_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}externalLink`;
const EXTERNAL_PATH_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}externalLinkPath`;
const CONNECTIONS_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}connections`;
const QUERY_TABLE_RELATIONSHIP = `${XLSX_OFFICE_REL_TYPE}queryTable`;

const CONTENT_TYPES = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
  <Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>
  <Override PartName="/xl/queryTables/queryTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>
</Types>`;
const WORKBOOK = `<workbook xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="sheet"/></sheets></workbook>`;
const WORKBOOK_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="sheet" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="external" Type="${EXTERNAL_LINK_RELATIONSHIP}" Target="externalLinks/externalLink1.xml"/><Relationship Id="connections" Type="${CONNECTIONS_RELATIONSHIP}" Target="connections.xml"/></Relationships>`;
const EXTERNAL_LINK = `<externalLink xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><externalBook r:id="path"><sheetNames><sheetName val="Data"/></sheetNames><definedNames><definedName name="RemoteValue" sheetId="0" refersTo="'[1]Data'!$A$1"/></definedNames></externalBook></externalLink>`;
const EXTERNAL_LINK_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="path" Type="${EXTERNAL_PATH_RELATIONSHIP}" Target="https://user:password@example.com/private/book.xlsx?token=secret#fragment" TargetMode="External"/></Relationships>`;
const CONNECTIONS = `<connections xmlns="${XLSX_SPREADSHEET_NS}" count="1"><connection id="1" name="Warehouse" description="Sales data" type="5" refreshedVersion="8" interval="30" background="1" deleted="0" keepAlive="true" refreshOnLoad="1" saveData="0" credentials="stored" singleSignOnId="secret-sso"><dbPr connection="Server=secret;Password=hunter2" command="SELECT * FROM private_table"/></connection></connections>`;
const WORKSHEET = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData/></worksheet>`;
const WORKSHEET_RELS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="query" Type="${QUERY_TABLE_RELATIONSHIP}" Target="../queryTables/queryTable1.xml"/></Relationships>`;
const QUERY_TABLE = `<queryTable xmlns="${XLSX_SPREADSHEET_NS}" name="Sales Query" connectionId="1" adjustColumnWidth="0" applyAlignmentFormats="1" applyBorderFormats="0" applyFontFormats="true" applyNumberFormats="false" applyPatternFormats="1" applyWidthHeightFormats="0" backgroundRefresh="false" disableEdit="1" preserveFormatting="0" refreshOnLoad="true"/>`;

function parts(changes: XlsxBlackBoxOverrides = {}): XlsxBlackBoxOverrides {
  return {
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
    'xl/connections.xml': CONNECTIONS,
    'xl/externalLinks/externalLink1.xml': EXTERNAL_LINK,
    'xl/externalLinks/_rels/externalLink1.xml.rels': EXTERNAL_LINK_RELS,
    'xl/queryTables/queryTable1.xml': QUERY_TABLE,
    'xl/sharedStrings.xml': null,
    'xl/styles.xml': null,
    'xl/workbook.xml': WORKBOOK,
    'xl/worksheets/sheet1.xml': WORKSHEET,
    'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS,
    ...changes,
  };
}

async function bytes(changes: XlsxBlackBoxOverrides = {}): Promise<Uint8Array> {
  return createIndependentXlsx(parts(changes));
}

async function capture(
  changes: XlsxBlackBoxOverrides,
): Promise<XlsxParseError> {
  try {
    await parseXlsx(await bytes(changes), { errorMode: 'strict' });
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    return error as XlsxParseError;
  }
  throw new Error('Expected external metadata parsing to fail');
}

describe('XLSX external metadata', () => {
  it('returns redacted external links, connections, and query tables', async () => {
    const document = await parseXlsx(await bytes(), { errorMode: 'strict' });
    expect(document.externalLinks).toStrictEqual([
      {
        definedNames: [
          {
            formula: "'[1]Data'!$A$1",
            name: 'RemoteValue',
            sheetId: 0,
          },
        ],
        index: 0,
        sheetNames: ['Data'],
        target: {
          host: 'example.com',
          kind: 'https',
          redacted: true,
        },
      },
    ]);
    expect(document.connections).toStrictEqual([
      {
        background: true,
        credentialsOmitted: true,
        deleted: false,
        description: 'Sales data',
        id: 1,
        keepAlive: true,
        name: 'Warehouse',
        refreshedVersion: 8,
        refreshInterval: 30,
        refreshOnLoad: true,
        saveData: false,
        sourceDataOmitted: true,
        type: 5,
      },
    ]);
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.queryTables : undefined,
    ).toStrictEqual([
      {
        adjustColumnWidth: false,
        applyAlignmentFormats: true,
        applyBorderFormats: false,
        applyFontFormats: true,
        applyNumberFormats: false,
        applyPatternFormats: true,
        applyWidthHeightFormats: false,
        backgroundRefresh: false,
        connectionId: 1,
        disableEdit: true,
        name: 'Sales Query',
        preserveFormatting: false,
        refreshOnLoad: true,
      },
    ]);
    expect(JSON.stringify(document)).not.toMatch(
      /password|hunter2|private_table|secret-sso|token=secret/u,
    );
  });

  it('normalizes lexical helpers and redacted target classes', () => {
    expect(
      parseXlsxExternalUnsignedInteger('4294967295', undefined, 'id', 'part'),
    ).toBe(0xffff_ffff);
    expect(parseXlsxExternalUnsignedInteger(undefined, 2, 'id', 'part')).toBe(
      2,
    );
    for (const value of ['-1', '01', '1.0', '1 ', '4294967296']) {
      expect(() =>
        parseXlsxExternalUnsignedInteger(value, undefined, 'id', 'part'),
      ).toThrow('id');
    }
    expect(parseXlsxExternalBoolean('0', true, 'flag', 'part')).toBe(false);
    expect(parseXlsxExternalBoolean('false', true, 'flag', 'part')).toBe(false);
    expect(parseXlsxExternalBoolean('1', false, 'flag', 'part')).toBe(true);
    expect(parseXlsxExternalBoolean('true', false, 'flag', 'part')).toBe(true);
    expect(parseXlsxExternalBoolean(undefined, true, 'flag', 'part')).toBe(
      true,
    );
    expect(() =>
      parseXlsxExternalBoolean('bad', false, 'flag', 'part'),
    ).toThrow('flag');
    expect(
      redactXlsxExternalTarget('../private/book.xlsx', 'part'),
    ).toStrictEqual({ kind: 'relative', redacted: true });
    expect(
      redactXlsxExternalTarget('file:///private/book.xlsx', 'part'),
    ).toStrictEqual({ kind: 'file', redacted: true });
    expect(
      redactXlsxExternalTarget('http://EXAMPLE.com/private', 'part'),
    ).toStrictEqual({ host: 'example.com', kind: 'http', redacted: true });
    expect(redactXlsxExternalTarget('book.xlsx', 'part')).toStrictEqual({
      kind: 'relative',
      redacted: true,
    });
    expect(() => redactXlsxExternalTarget('bad :// target', 'part')).toThrow(
      'External workbook target is invalid',
    );
  });

  it('detects every redacted connection secret marker at any depth', () => {
    for (const name of ['credentials', 'singleSignOnId', 'ssoId']) {
      expect(xlsxConnectionCredentialsOmitted({ [name]: 'secret' })).toBe(true);
    }
    expect(xlsxConnectionCredentialsOmitted({ name: 'safe' })).toBe(false);
    for (const name of [
      'command',
      'connection',
      'credentials',
      'odcFile',
      'serverCredentialsMethod',
      'sourceFile',
      'ssoId',
      'url',
    ]) {
      expect(
        xlsxConnectionSourceDataOmitted({
          child: { attrs: { [name]: 'secret' } },
        }),
      ).toBe(true);
    }
    expect(xlsxConnectionSourceDataOmitted({ attrs: { name: 'safe' } })).toBe(
      false,
    );
    expect(
      xlsxConnectionSourceDataOmitted({
        safe: { attrs: { name: 'safe' } },
        secret: { attrs: { connection: 'secret' } },
      }),
    ).toBe(true);
    expect(
      xlsxConnectionSourceDataOmitted({
        children: [
          { attrs: { name: 'safe' } },
          { attrs: { command: 'secret' } },
        ],
      }),
    ).toBe(true);
    expect(xlsxConnectionSourceDataOmitted('bad')).toBe(false);
  });

  it('preserves safe metadata through portable exact R0', async () => {
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

  it('preserves defaults without inventing source details', async () => {
    const document = await parseXlsx(
      await bytes({
        'xl/connections.xml': `<connections xmlns="${XLSX_SPREADSHEET_NS}"><connection id="1"/></connections>`,
        'xl/externalLinks/externalLink1.xml': `<externalLink xmlns="${XLSX_SPREADSHEET_NS}"/>`,
        'xl/externalLinks/_rels/externalLink1.xml.rels': null,
        'xl/queryTables/queryTable1.xml': `<queryTable xmlns="${XLSX_SPREADSHEET_NS}" name="Defaults" connectionId="1"/>`,
      }),
      { errorMode: 'strict' },
    );
    expect(document.externalLinks?.[0]).not.toHaveProperty('target');
    expect(document.connections?.[0]).toStrictEqual({
      background: false,
      credentialsOmitted: false,
      deleted: false,
      id: 1,
      keepAlive: false,
      refreshOnLoad: false,
      saveData: false,
      sourceDataOmitted: false,
    });
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.queryTables?.[0] : undefined,
    ).toStrictEqual({
      adjustColumnWidth: true,
      applyAlignmentFormats: false,
      applyBorderFormats: false,
      applyFontFormats: false,
      applyNumberFormats: false,
      applyPatternFormats: false,
      applyWidthHeightFormats: false,
      backgroundRefresh: true,
      connectionId: 1,
      disableEdit: false,
      name: 'Defaults',
      preserveFormatting: true,
      refreshOnLoad: false,
    });
  });

  it('accepts an empty connection collection when no query table depends on it', async () => {
    const document = await parseXlsx(
      await bytes({
        'xl/connections.xml': `<connections xmlns="${XLSX_SPREADSHEET_NS}"/>`,
        'xl/queryTables/queryTable1.xml': null,
        'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"/>`,
      }),
      { errorMode: 'strict' },
    );
    expect(document).not.toHaveProperty('connections');
    const sheet = document.sheets[0]!;
    expect(sheet).not.toHaveProperty('queryTables');
  });

  it('parses a complete Strict external-metadata graph', async () => {
    const strictSheet = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictRelationship =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const source = await createIndependentXlsx(
      parts({
        '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="root" Type="${strictRelationship}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELS.replaceAll(
          XLSX_OFFICE_REL_TYPE,
          `${strictRelationship}/`,
        ),
        'xl/connections.xml': CONNECTIONS.replaceAll(
          XLSX_SPREADSHEET_NS,
          strictSheet,
        ),
        'xl/externalLinks/externalLink1.xml': EXTERNAL_LINK.replaceAll(
          XLSX_SPREADSHEET_NS,
          strictSheet,
        ).replaceAll(XLSX_OFFICE_REL_NS, strictRelationship),
        'xl/externalLinks/_rels/externalLink1.xml.rels':
          EXTERNAL_LINK_RELS.replaceAll(
            XLSX_OFFICE_REL_TYPE,
            `${strictRelationship}/`,
          ),
        'xl/queryTables/queryTable1.xml': QUERY_TABLE.replaceAll(
          XLSX_SPREADSHEET_NS,
          strictSheet,
        ),
        'xl/workbook.xml': WORKBOOK.replaceAll(
          XLSX_SPREADSHEET_NS,
          strictSheet,
        ).replaceAll(XLSX_OFFICE_REL_NS, strictRelationship),
        'xl/worksheets/_rels/sheet1.xml.rels': WORKSHEET_RELS.replaceAll(
          XLSX_OFFICE_REL_TYPE,
          `${strictRelationship}/`,
        ),
        'xl/worksheets/sheet1.xml': WORKSHEET.replaceAll(
          XLSX_SPREADSHEET_NS,
          strictSheet,
        ),
      }),
    );
    const document = await parseXlsx(source, { errorMode: 'strict' });
    expect(document.externalLinks).toHaveLength(1);
    expect(document.connections).toHaveLength(1);
    const sheet = document.sheets[0]!;
    expect(
      sheet.kind === 'worksheet' ? sheet.queryTables : undefined,
    ).toHaveLength(1);
  });

  it('parses reserved and arbitrary-prefix external metadata roots', async () => {
    const prefix = (
      xml: string,
      names: readonly string[],
      sourcePrefix: 'p' | 'x',
    ): string => {
      let output = xml.replace(
        `xmlns="${XLSX_SPREADSHEET_NS}"`,
        `xmlns:${sourcePrefix}="${XLSX_SPREADSHEET_NS}"`,
      );
      for (const name of names) {
        output = output.replaceAll(`<${name}`, `<${sourcePrefix}:${name}`);
        output = output.replaceAll(`</${name}>`, `</${sourcePrefix}:${name}>`);
      }
      return output;
    };
    for (const sourcePrefix of ['p', 'x'] as const) {
      const document = await parseXlsx(
        await bytes({
          'xl/connections.xml': prefix(
            CONNECTIONS,
            ['connections', 'connection', 'dbPr'],
            sourcePrefix,
          ),
          'xl/externalLinks/externalLink1.xml': prefix(
            EXTERNAL_LINK,
            [
              'externalLink',
              'externalBook',
              'sheetNames',
              'sheetName',
              'definedNames',
              'definedName',
            ],
            sourcePrefix,
          ),
          'xl/queryTables/queryTable1.xml': prefix(
            QUERY_TABLE,
            ['queryTable'],
            sourcePrefix,
          ),
        }),
        { errorMode: 'strict' },
      );
      expect(document.connections).toHaveLength(1);
      expect(document.externalLinks).toHaveLength(1);
    }
  });

  it('treats DDE as optional rejected active content by mode', async () => {
    const dde = `<externalLink xmlns="${XLSX_SPREADSHEET_NS}"><ddeLink ddeService="Excel" ddeTopic="secret"/></externalLink>`;
    const tolerant = await parseXlsxWithDiagnostics(
      await bytes({ 'xl/externalLinks/externalLink1.xml': dde }),
    );
    expect(tolerant.document).not.toHaveProperty('externalLinks');
    expect(tolerant.document).not.toHaveProperty('connections');
    const sheet = tolerant.document.sheets[0]!;
    expect(sheet).not.toHaveProperty('queryTables');
    expect(tolerant.diagnostics).toStrictEqual([
      {
        code: 'security-rejected-content',
        message: 'External DDE or OLE link is not allowed',
        part: 'xl/externalLinks/externalLink1.xml',
        severity: 'warning',
      },
      {
        code: 'invalid-document-value',
        message: 'Query table connection reference is invalid',
        part: 'xl/queryTables/queryTable1.xml',
        severity: 'warning',
      },
    ]);
    expect(
      (await capture({ 'xl/externalLinks/externalLink1.xml': dde })).diagnostic
        .code,
    ).toBe('security-rejected-content');
  });

  it('enforces connection count limits at exact boundaries', async () => {
    await expect(
      parseXlsx(await bytes(), {
        errorMode: 'strict',
        limits: { maxTables: 1 },
      }),
    ).resolves.toBeDefined();
    const twoConnections = CONNECTIONS.replace(
      'count="1"',
      'count="2"',
    ).replace('</connections>', '<connection id="2"/></connections>');
    try {
      await parseXlsx(await bytes({ 'xl/connections.xml': twoConnections }), {
        errorMode: 'strict',
        limits: { maxTables: 1 },
      });
    } catch (error) {
      expect((error as XlsxParseError).diagnostic).toMatchObject({
        actual: 2,
        limit: 1,
        limitName: 'maxTables',
      });
      return;
    }
    throw new Error('Expected connection count limit failure');
  });

  it.each([
    ['background', '1', 'Connection background flag is invalid'],
    ['deleted', '0', 'Connection deleted flag is invalid'],
    ['keepAlive', 'true', 'Connection keep-alive flag is invalid'],
    ['refreshOnLoad', '1', 'Connection refresh-on-load flag is invalid'],
    ['saveData', '0', 'Connection save-data flag is invalid'],
  ] as const)(
    'rejects invalid connection flag %s',
    async (name, value, message) => {
      const xml = CONNECTIONS.replace(`${name}="${value}"`, `${name}="bad"`);
      expect(
        (await capture({ 'xl/connections.xml': xml })).diagnostic.message,
      ).toBe(message);
    },
  );

  it.each([
    [
      'adjustColumnWidth',
      '0',
      'Query table adjust-column-width flag is invalid',
    ],
    [
      'applyAlignmentFormats',
      '1',
      'Query table alignment-format flag is invalid',
    ],
    ['applyBorderFormats', '0', 'Query table border-format flag is invalid'],
    ['applyFontFormats', 'true', 'Query table font-format flag is invalid'],
    [
      'applyNumberFormats',
      'false',
      'Query table number-format flag is invalid',
    ],
    ['applyPatternFormats', '1', 'Query table pattern-format flag is invalid'],
    [
      'applyWidthHeightFormats',
      '0',
      'Query table width-height-format flag is invalid',
    ],
    [
      'backgroundRefresh',
      'false',
      'Query table background-refresh flag is invalid',
    ],
    ['disableEdit', '1', 'Query table disable-edit flag is invalid'],
    [
      'preserveFormatting',
      '0',
      'Query table preserve-formatting flag is invalid',
    ],
    ['refreshOnLoad', 'true', 'Query table refresh-on-load flag is invalid'],
  ] as const)(
    'rejects invalid query-table flag %s',
    async (name, value, message) => {
      const xml = QUERY_TABLE.replace(`${name}="${value}"`, `${name}="bad"`);
      expect(
        (await capture({ 'xl/queryTables/queryTable1.xml': xml })).diagnostic
          .message,
      ).toBe(message);
    },
  );

  it('rejects case-insensitive duplicate query-table names across owned parts', async () => {
    const contentTypes = CONTENT_TYPES.replace(
      '</Types>',
      '<Override PartName="/xl/queryTables/queryTable2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/></Types>',
    );
    const relationships = WORKSHEET_RELS.replace(
      '</Relationships>',
      `<Relationship Id="query2" Type="${QUERY_TABLE_RELATIONSHIP}" Target="../queryTables/queryTable2.xml"/></Relationships>`,
    );
    expect(
      (
        await capture({
          '[Content_Types].xml': contentTypes,
          'xl/queryTables/queryTable1.xml': QUERY_TABLE.replace(
            'name="Sales Query"',
            'name="Straße"',
          ),
          'xl/queryTables/queryTable2.xml': QUERY_TABLE.replace(
            'name="Sales Query"',
            'name="STRASSE"',
          ),
          'xl/worksheets/_rels/sheet1.xml.rels': relationships,
        })
      ).diagnostic.message,
    ).toBe('Query table names are duplicated');
  });

  it.each([
    [
      {
        '[Content_Types].xml': CONTENT_TYPES.replace(
          'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml',
          'application/xml',
        ),
      },
      'External metadata target has the wrong content type',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELS.replace(
          'Target="connections.xml"',
          'Target="https://example.invalid/connections.xml" TargetMode="External"',
        ),
      },
      'External metadata relationship must be internal',
    ],
    [
      {
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELS.replace(
          '</Relationships>',
          `<Relationship Id="connections2" Type="${CONNECTIONS_RELATIONSHIP}" Target="connections.xml"/></Relationships>`,
        ),
      },
      'External metadata relationship target is duplicated',
    ],
    [
      {
        'xl/externalLinks/_rels/externalLink1.xml.rels':
          EXTERNAL_LINK_RELS.replace(
            'https://user:password@example.com/private/book.xlsx?token=secret#fragment',
            'javascript:alert(1)',
          ),
      },
      'External workbook protocol is not allowed',
    ],
    [
      {
        'xl/externalLinks/externalLink1.xml': EXTERNAL_LINK.replace(
          'r:id="path"',
          'r:id=""',
        ),
      },
      'External workbook relationship reference is invalid',
    ],
    [
      { 'xl/externalLinks/_rels/externalLink1.xml.rels': null },
      'External workbook relationships are missing',
    ],
    [
      {
        'xl/externalLinks/_rels/externalLink1.xml.rels':
          EXTERNAL_LINK_RELS.replace(
            EXTERNAL_PATH_RELATIONSHIP,
            `${XLSX_OFFICE_REL_TYPE}image`,
          ),
      },
      'External workbook path relationship is invalid',
    ],
    [
      {
        'xl/externalLinks/_rels/externalLink1.xml.rels':
          EXTERNAL_LINK_RELS.replace(
            'Target="https://user:password@example.com/private/book.xlsx?token=secret#fragment" TargetMode="External"',
            'Target="../book.xlsx"',
          ),
      },
      'External workbook path relationship is invalid',
    ],
    [
      { 'xl/connections.xml': `<wrong xmlns="${XLSX_SPREADSHEET_NS}"/>` },
      'XLSX connections root is missing',
    ],
    [
      {
        'xl/externalLinks/externalLink1.xml': EXTERNAL_LINK.replace(
          XLSX_SPREADSHEET_NS,
          'urn:wrong',
        ),
      },
      'XLSX externalLink root has the wrong namespace',
    ],
    [
      {
        'xl/queryTables/queryTable1.xml': QUERY_TABLE.replace(
          XLSX_SPREADSHEET_NS,
          'urn:wrong',
        ),
      },
      'XLSX queryTable root has the wrong namespace',
    ],
    [
      {
        'xl/connections.xml': `<connections xmlns="${XLSX_SPREADSHEET_NS}"><connection>bad</connection></connections>`,
      },
      'Connection collection is invalid',
    ],
    [
      {
        'xl/connections.xml': CONNECTIONS.replace(
          '</connections>',
          `<x:connection xmlns:x="${XLSX_SPREADSHEET_NS}" id="2"/></connections>`,
        ),
      },
      'Connection collection is duplicated',
    ],
    [
      {
        'xl/connections.xml': CONNECTIONS.replace('count="1"', 'count="bad"'),
      },
      'Connection count is invalid',
    ],
    [
      {
        'xl/connections.xml': CONNECTIONS.replace('count="1"', 'count="2"'),
      },
      'Connection count does not match',
    ],
    [
      {
        'xl/connections.xml': CONNECTIONS.replace('id="1"', 'id="bad"'),
      },
      'Connection ID is invalid',
    ],
    [
      {
        'xl/connections.xml': CONNECTIONS.replace(
          'count="1"',
          'count="2"',
        ).replace('</connections>', '<connection id="1"/></connections>'),
      },
      'Connection IDs are duplicated',
    ],
    [
      {
        'xl/connections.xml': CONNECTIONS.replace(
          'refreshedVersion="8"',
          'refreshedVersion="bad"',
        ),
      },
      'Connection refreshed version is invalid',
    ],
    [
      {
        'xl/connections.xml': CONNECTIONS.replace(
          'interval="30"',
          'interval="bad"',
        ),
      },
      'Connection refresh interval is invalid',
    ],
    [
      {
        'xl/connections.xml': CONNECTIONS.replace('type="5"', 'type="bad"'),
      },
      'Connection type is invalid',
    ],
    [
      {
        'xl/externalLinks/externalLink1.xml': EXTERNAL_LINK.replace(
          'name="RemoteValue"',
          'name=""',
        ),
      },
      'External defined name is invalid',
    ],
    [
      {
        'xl/externalLinks/externalLink1.xml': EXTERNAL_LINK.replace(
          'refersTo="\'[1]Data\'!$A$1"',
          'refersTo=" bad "',
        ),
      },
      'External defined-name formula is invalid',
    ],
    [
      {
        'xl/externalLinks/externalLink1.xml': EXTERNAL_LINK.replace(
          'sheetId="0"',
          'sheetId="bad"',
        ),
      },
      'External defined-name sheet ID is invalid',
    ],
    [
      {
        'xl/externalLinks/externalLink1.xml': EXTERNAL_LINK.replace(
          'val="Data"',
          'val=""',
        ),
      },
      'External sheet name is invalid',
    ],
    [
      {
        'xl/externalLinks/externalLink1.xml': EXTERNAL_LINK.replace(
          '<externalBook r:id="path">',
          '<externalBook r:id="path"></externalBook><externalBook r:id="path">',
        ),
      },
      'External workbook owner is duplicated',
    ],
    [
      {
        'xl/externalLinks/externalLink1.xml': `<externalLink xmlns="${XLSX_SPREADSHEET_NS}"><oleLink progId="Excel.Sheet"/></externalLink>`,
      },
      'External DDE or OLE link is not allowed',
    ],
    [
      {
        'xl/queryTables/queryTable1.xml': QUERY_TABLE.replace(
          'connectionId="1"',
          'connectionId="2"',
        ),
      },
      'Query table connection reference is invalid',
    ],
    [
      {
        'xl/queryTables/queryTable1.xml': QUERY_TABLE.replace(
          'connectionId="1"',
          'connectionId="bad"',
        ),
      },
      'Query table connection ID is invalid',
    ],
    [
      {
        'xl/queryTables/queryTable1.xml': QUERY_TABLE.replace(
          'name="Sales Query"',
          'name=""',
        ),
      },
      'Query table name is invalid',
    ],
    [
      {
        'xl/queryTables/queryTable1.xml': QUERY_TABLE.replace(
          'refreshOnLoad="true"',
          'refreshOnLoad="bad"',
        ),
      },
      'Query table refresh-on-load flag is invalid',
    ],
  ] as const)(
    'rejects invalid external metadata %#',
    async (changes, message) => {
      expect((await capture(changes)).diagnostic.message).toBe(message);
    },
  );
});
