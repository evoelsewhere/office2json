import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common/xml/tree';
import {
  parseXlsx,
  readXlsxRoundTrip,
  writeXlsxRoundTrip,
} from '../../src/formats/xlsx';
import {
  parseXlsxFeaturePropertyBagPart,
  xlsxFeaturePropertyBagNamespace,
} from '../../src/formats/xlsx/internal/feature-property-bag';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import { parseXlsxStylePart } from '../../src/formats/xlsx/internal/styles';
import {
  createIndependentXlsx,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const FEATURE_REL =
  'http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag';
const CHECKBOX_URI = '{C7286773-470A-42A8-94C5-96B5CB345126}';
const FEATURE_NAMESPACE =
  'http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag';

const CONTENT_TYPES = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/featurePropertyBag/featurePropertyBag.xml" ContentType="application/vnd.ms-excel.featurepropertybag+xml"/>
</Types>`;

const RELATIONSHIPS = `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
  <Relationship Id="rIdSheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdStyles" Type="${XLSX_OFFICE_REL_TYPE}styles" Target="styles.xml"/>
  <Relationship Id="rIdSharedStrings" Type="${XLSX_OFFICE_REL_TYPE}sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="feature" Type="${FEATURE_REL}" Target="featurePropertyBag/featurePropertyBag.xml"/>
</Relationships>`;

const FEATURE_BAGS = `<FeaturePropertyBags xmlns="${FEATURE_NAMESPACE}">
  <bag type="Checkbox"/>
  <bag type="XFControls"><bagId k="CellControl">0</bagId></bag>
  <bag type="XFComplement"><bagId k="XFControls">1</bagId></bag>
  <bag type="XFComplements" extRef="XFComplementsMapperExtRef">
    <a k="MappedFeaturePropertyBags"><bagId>2</bagId></a>
  </bag>
</FeaturePropertyBags>`;

function checkboxExtension(index = '0', namespace = FEATURE_NAMESPACE): string {
  return `<extLst><ext uri="${CHECKBOX_URI}" xmlns:xfpb="${namespace}"><xfpb:xfComplement i="${index}"/></ext></extLst>`;
}

function styles(extension = checkboxExtension()): string {
  return `<styleSheet xmlns="${XLSX_SPREADSHEET_NS}">
    <fonts count="1"><font/></fonts>
    <fills count="1"><fill/></fills>
    <borders count="1"><border/></borders>
    <cellStyleXfs count="1"><xf/></cellStyleXfs>
    <cellXfs count="2"><xf/><xf>${extension}</xf></cellXfs>
  </styleSheet>`;
}

interface Overrides {
  contentTypes?: string;
  featureBags?: string | null;
  relationships?: string;
  styles?: string;
}

async function source(overrides: Overrides = {}): Promise<Uint8Array> {
  return createIndependentXlsx({
    '[Content_Types].xml': overrides.contentTypes ?? CONTENT_TYPES,
    'xl/_rels/workbook.xml.rels': overrides.relationships ?? RELATIONSHIPS,
    'xl/featurePropertyBag/featurePropertyBag.xml':
      overrides.featureBags === undefined
        ? FEATURE_BAGS
        : overrides.featureBags,
    'xl/styles.xml': overrides.styles ?? styles(),
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><sheetData>
      <row r="1"><c r="A1" s="1" t="b"><v>1</v></c></row>
      <row r="2"><c r="A2" s="0" t="b"><v>0</v></c></row>
    </sheetData></worksheet>`,
  });
}

describe('XLSX checkbox styles', () => {
  it('locks the published feature-property-bag namespace independently', () => {
    expect(xlsxFeaturePropertyBagNamespace()).toBe(FEATURE_NAMESPACE);
  });
  it('resolves the feature-property-bag graph into a portable checkbox style', async () => {
    const document = await parseXlsx(await source(), { errorMode: 'strict' });
    expect(document.styles).toStrictEqual([{}, { checkbox: true }]);
    const sheet = document.sheets[0];
    if (sheet?.kind !== 'worksheet') throw new Error('Expected worksheet');
    expect(sheet.rows[0]?.cells[0]).toMatchObject({
      content: { kind: 'value', value: { kind: 'boolean', value: true } },
      style: 1,
    });
    expect(sheet.rows[1]?.cells[0]).toMatchObject({
      content: { kind: 'value', value: { kind: 'boolean', value: false } },
      style: 0,
    });
    expect(JSON.parse(JSON.stringify(document))).toStrictEqual(document);
  });

  it('preserves checkbox style parts byte-for-byte through standalone R0 JSON', async () => {
    const input = await source();
    const snapshot = await readXlsxRoundTrip(input);
    const output = await writeXlsxRoundTrip(
      JSON.parse(JSON.stringify(snapshot)) as typeof snapshot,
    );
    expect(output.data).toStrictEqual(input);
    expect(output.report.level).toBe('R0');
  });

  it('ignores unrelated style extensions without requiring feature bags', async () => {
    const relationships = RELATIONSHIPS.replace(
      /\s*<Relationship Id="feature"[^>]+\/>/u,
      '',
    );
    const document = await parseXlsx(
      await source({
        featureBags: null,
        relationships,
        styles: styles(
          '<extLst><ext uri="urn:future"><future/></ext></extLst>',
        ),
      }),
      { errorMode: 'strict' },
    );
    expect(document.styles).toStrictEqual([{}]);
  });

  it('uses the authored mapping position when unrelated feature mappings precede checkbox', async () => {
    const featureBags = FEATURE_BAGS.replace(
      '<a k="MappedFeaturePropertyBags">',
      '<a k="Other"><bagId>2</bagId></a><a k="MappedFeaturePropertyBags">',
    );
    const document = await parseXlsx(
      await source({ featureBags, styles: styles(checkboxExtension('1')) }),
      { errorMode: 'strict' },
    );
    expect(document.styles).toStrictEqual([{}, { checkbox: true }]);
  });

  it('accepts a checkbox complement using a default feature namespace', async () => {
    const extension = `<extLst><ext uri="${CHECKBOX_URI}"><xfComplement xmlns="${FEATURE_NAMESPACE}" i="0"/></ext></extLst>`;
    const document = await parseXlsx(
      await source({ styles: styles(extension) }),
      {
        errorMode: 'strict',
      },
    );
    expect(document.styles).toStrictEqual([{}, { checkbox: true }]);
  });

  it.each([
    [
      {
        relationships: RELATIONSHIPS.replace(
          /\s*<Relationship Id="feature"[^>]+\/>/u,
          '',
        ),
        featureBags: null,
      },
      'Styles checkbox feature property bag is missing',
    ],
    [
      {
        relationships: RELATIONSHIPS.replace(
          '</Relationships>',
          `<Relationship Id="feature2" Type="${FEATURE_REL}" Target="featurePropertyBag/featurePropertyBag.xml"/></Relationships>`,
        ),
      },
      'Feature property bag relationship is duplicated',
    ],
    [
      {
        relationships: RELATIONSHIPS.replace(
          'Target="featurePropertyBag/featurePropertyBag.xml"',
          'Target="https://example.test/bag" TargetMode="External"',
        ),
      },
      'Feature property bag relationship must be internal',
    ],
    [
      {
        contentTypes: CONTENT_TYPES.replace(
          'application/vnd.ms-excel.featurepropertybag+xml',
          'application/xml',
        ),
      },
      'Feature property bag target has the wrong content type',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(FEATURE_NAMESPACE, 'urn:wrong'),
      },
      'Feature property bag root has the wrong namespace',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(
          '<bag type="Checkbox"/>',
          '<bag>text</bag>',
        ),
      },
      'Feature property bag collection is invalid',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(
          '<bag type="Checkbox"/>',
          '<x:bag xmlns:x="urn:wrong" type="Checkbox"/>',
        ),
      },
      'Feature property bag element has the wrong namespace',
    ],
    [
      { featureBags: FEATURE_BAGS.replace('type="Checkbox"', '') },
      'Feature property bag type is invalid',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(
          '<bagId>2</bagId>',
          '<bagId>9</bagId>',
        ),
      },
      'Feature property bag reference is invalid',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(
          '<bagId>2</bagId>',
          '<x:bagId xmlns:x="urn:wrong">2</x:bagId>',
        ),
      },
      'Feature property bag element has the wrong namespace',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(
          '<bagId>2</bagId>',
          '<bagId><nested/></bagId>',
        ),
      },
      'Feature property bag reference is invalid',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(
          '<bagId>2</bagId>',
          '<bagId>1a</bagId>',
        ),
      },
      'Feature property bag reference is invalid',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(
          '<bagId>2</bagId>',
          '<bagId>02</bagId>',
        ),
      },
      'Feature property bag reference is invalid',
    ],
    [
      { featureBags: FEATURE_BAGS.replace('k="XFControls"', 'k="Wrong"') },
      'Feature property bag mapping is invalid',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(
          '<bagId k="XFControls">1</bagId>',
          '<bagId>1</bagId>',
        ),
      },
      'Feature property bag mapping is invalid',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(
          '<bagId>2</bagId>',
          '<bagId>0</bagId>',
        ),
      },
      'Styles checkbox feature reference is invalid',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(
          '<bagId k="XFControls">1</bagId>',
          '<bagId k="XFControls">0</bagId>',
        ),
      },
      'Styles checkbox feature reference is invalid',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(
          '<bagId k="CellControl">0</bagId>',
          '<bagId k="CellControl">1</bagId>',
        ),
      },
      'Styles checkbox feature reference is invalid',
    ],
    [
      {
        featureBags: FEATURE_BAGS.replace(
          '</FeaturePropertyBags>',
          '<bag type="XFComplements"/></FeaturePropertyBags>',
        ),
      },
      'Feature property bag complement collection is duplicated',
    ],
    [
      { styles: styles(checkboxExtension('1')) },
      'Styles checkbox feature reference is invalid',
    ],
    [
      { styles: styles(checkboxExtension('bad')) },
      'Styles checkbox feature reference is invalid',
    ],
    [
      { styles: styles(checkboxExtension('0', 'urn:wrong')) },
      'Styles checkbox extension has the wrong namespace',
    ],
    [
      {
        styles: styles(
          `<extLst><ext uri="${CHECKBOX_URI}"/><ext uri="${CHECKBOX_URI}"/></extLst>`,
        ),
      },
      'Styles checkbox extension is duplicated',
    ],
    [
      {
        styles: styles(
          `<extLst><ext uri="${CHECKBOX_URI}" xmlns:xfpb="${FEATURE_NAMESPACE}"><xfpb:wrong/></ext></extLst>`,
        ),
      },
      'Styles checkbox extension is invalid',
    ],
    [
      { styles: styles('<extLst>text</extLst>') },
      'Styles XF extension list is invalid',
    ],
    [
      { styles: styles('<extLst><ext>text</ext></extLst>') },
      'Styles XF extension list is invalid',
    ],
    [
      {
        styles: styles(
          `<extLst><ext uri="${CHECKBOX_URI}" xmlns:xfpb="${FEATURE_NAMESPACE}"><xfpb:xfComplement>text</xfpb:xfComplement></ext></extLst>`,
        ),
      },
      'Styles checkbox extension is invalid',
    ],
  ] as const)(
    'rejects malformed checkbox style package %#',
    async (overrides, message) => {
      await expect(
        parseXlsx(await source(overrides), { errorMode: 'strict' }),
      ).rejects.toThrow(message);
    },
  );

  it('enforces feature bags in the aggregate style limit', async () => {
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        limits: { maxStyles: 14 },
      }),
    ).resolves.toBeDefined();
    await expect(
      parseXlsx(await source(), {
        errorMode: 'strict',
        limits: { maxStyles: 13 },
      }),
    ).rejects.toMatchObject({
      diagnostic: { code: 'resource-limit-exceeded', limitName: 'maxStyles' },
    });
    const eightBags = {
      FeaturePropertyBags: {
        attrs: { xmlns: FEATURE_NAMESPACE },
        bag: Array.from({ length: 8 }, (_, index) => ({
          attrs: { type: `Type${index}` },
        })),
      },
    } as unknown as XmlLookupValue;
    expect(() =>
      parseXlsxFeaturePropertyBagPart(eightBags, 'feature.xml', {
        ...defaultXlsxResourceLimits(),
        maxStyles: 8,
      }),
    ).not.toThrow();
    expect(() =>
      parseXlsxFeaturePropertyBagPart(eightBags, 'feature.xml', {
        ...defaultXlsxResourceLimits(),
        maxStyles: 7,
      }),
    ).toThrow('maxStyles');
  });

  it('parses reserved-prefix normalized feature bags', () => {
    const registry = parseXlsxFeaturePropertyBagPart(
      {
        'ns_fpb:FeaturePropertyBags': {
          attrs: { 'xmlns:fpb': FEATURE_NAMESPACE },
          'ns_fpb:bag': [
            {
              attrs: {
                type: 'Checkbox',
                'xmlns:fpb': FEATURE_NAMESPACE,
              },
            },
            {
              attrs: {
                type: 'XFControls',
                'xmlns:fpb': FEATURE_NAMESPACE,
              },
              'ns_fpb:bagId': {
                attrs: {
                  k: 'CellControl',
                  'xmlns:fpb': FEATURE_NAMESPACE,
                },
                value: '0',
              },
            },
            {
              attrs: {
                type: 'XFComplement',
                'xmlns:fpb': FEATURE_NAMESPACE,
              },
              'ns_fpb:bagId': {
                attrs: {
                  k: 'XFControls',
                  'xmlns:fpb': FEATURE_NAMESPACE,
                },
                value: '1',
              },
            },
            {
              attrs: {
                type: 'XFComplements',
                'xmlns:fpb': FEATURE_NAMESPACE,
              },
              'ns_fpb:a': {
                attrs: {
                  k: 'MappedFeaturePropertyBags',
                  'xmlns:fpb': FEATURE_NAMESPACE,
                },
                'ns_fpb:bagId': {
                  attrs: { 'xmlns:fpb': FEATURE_NAMESPACE },
                  value: '2',
                },
              },
            },
          ],
        },
      } as unknown as XmlLookupValue,
      'feature.xml',
      defaultXlsxResourceLimits(),
    );
    expect([...registry.checkboxComplements]).toStrictEqual([0]);
    expect(registry.records).toBe(8);
  });

  it('filters unrelated roots and accepts regular namespace prefixes', () => {
    expect(
      parseXlsxFeaturePropertyBagPart(
        {
          unrelated: {},
          'fpb:FeaturePropertyBags': {
            attrs: { 'xmlns:fpb': FEATURE_NAMESPACE },
          },
        } as unknown as XmlLookupValue,
        'feature.xml',
        defaultXlsxResourceLimits(),
      ),
    ).toMatchObject({ records: 0 });
    expect(() =>
      parseXlsxFeaturePropertyBagPart(
        {} as XmlLookupValue,
        'feature.xml',
        defaultXlsxResourceLimits(),
      ),
    ).toThrow('Feature property bag root is missing or duplicated');
  });

  it('keeps unrelated mapping positions out of the checkbox registry', () => {
    const registry = parseXlsxFeaturePropertyBagPart(
      {
        FeaturePropertyBags: {
          attrs: { xmlns: FEATURE_NAMESPACE },
          bag: [
            { attrs: { type: 'Checkbox' } },
            {
              attrs: { type: 'XFControls' },
              bagId: { attrs: { k: 'CellControl' }, value: '0' },
            },
            {
              attrs: { type: 'XFComplement' },
              bagId: { attrs: { k: 'XFControls' }, value: '1' },
            },
            {
              attrs: { type: 'XFComplements' },
              a: [
                { attrs: { k: 'Other' }, bagId: '2' },
                {
                  attrs: { k: 'MappedFeaturePropertyBags' },
                  bagId: '2',
                },
              ],
            },
          ],
        },
      } as unknown as XmlLookupValue,
      'feature.xml',
      defaultXlsxResourceLimits(),
    );
    expect([...registry.checkboxComplements]).toStrictEqual([1]);
  });

  it.each([2, { nested: {}, value: '0' }] as const)(
    'rejects normalized scalar feature reference %#',
    (bagId) => {
      expect(() =>
        parseXlsxFeaturePropertyBagPart(
          {
            FeaturePropertyBags: {
              attrs: { xmlns: FEATURE_NAMESPACE },
              bag: [
                { attrs: { type: 'Checkbox' } },
                {
                  attrs: { type: 'XFComplements' },
                  a: {
                    attrs: { k: 'MappedFeaturePropertyBags' },
                    bagId,
                  },
                },
              ],
            },
          } as unknown as XmlLookupValue,
          'feature.xml',
          defaultXlsxResourceLimits(),
        ),
      ).toThrow('Feature property bag reference is invalid');
    },
  );

  it('resolves multi-digit bag references without numeric coercion', () => {
    const bags: Record<string, unknown>[] = [
      { attrs: { type: 'Checkbox' } },
      ...Array.from({ length: 9 }, (_, index) => ({
        attrs: { type: `Unused${index}` },
      })),
      {
        attrs: { type: 'XFControls' },
        bagId: { attrs: { k: 'CellControl' }, value: '0' },
      },
      {
        attrs: { type: 'XFComplement' },
        bagId: { attrs: { k: 'XFControls' }, value: '10' },
      },
      { attrs: { type: 'UnusedTarget' } },
      {
        attrs: { type: 'XFComplements' },
        a: {
          attrs: { k: 'MappedFeaturePropertyBags' },
          bagId: '11',
        },
      },
    ];
    const registry = parseXlsxFeaturePropertyBagPart(
      {
        FeaturePropertyBags: {
          attrs: { xmlns: FEATURE_NAMESPACE },
          bag: bags,
        },
      } as unknown as XmlLookupValue,
      'feature.xml',
      defaultXlsxResourceLimits(),
    );
    expect([...registry.checkboxComplements]).toStrictEqual([0]);
  });

  it('parses a normalized reserved-prefix checkbox style extension', () => {
    const result = parseXlsxStylePart(
      {
        styleSheet: {
          attrs: { xmlns: XLSX_SPREADSHEET_NS },
          fonts: { attrs: { count: '1' }, font: {} },
          fills: { attrs: { count: '1' }, fill: {} },
          borders: { attrs: { count: '1' }, border: {} },
          cellStyleXfs: { attrs: { count: '1' }, xf: {} },
          cellXfs: {
            attrs: { count: '1' },
            xf: {
              extLst: {
                ext: {
                  attrs: { uri: CHECKBOX_URI, 'xmlns:xfpb': FEATURE_NAMESPACE },
                  'ns_xfpb:xfComplement': {
                    attrs: { i: '0', 'xmlns:xfpb': FEATURE_NAMESPACE },
                  },
                },
              },
            },
          },
        },
      } as unknown as XmlLookupValue,
      'transitional',
      'styles.xml',
      defaultXlsxResourceLimits(),
      {
        checkboxComplements: new Set([0]),
        part: 'feature.xml',
        records: 0,
      },
    );
    expect(result.styles).toStrictEqual([{ checkbox: true }]);
  });
});
