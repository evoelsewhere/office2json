import { describe, expect, it } from 'vitest';

import {
  parseXlsx,
  parseXlsxWithDiagnostics,
  XlsxParseError,
} from '../../src/formats/xlsx';
import {
  activeXlsxContentDiagnostic,
  classifyXlsxActiveContent,
} from '../../src/formats/xlsx/internal/security';
import {
  createIndependentXlsx,
  XLSX_CONTENT_TYPES_NS,
} from '../black-box/xlsx-package';

function contentTypes(overrides: string): string {
  return `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    ${overrides}
  </Types>`;
}

describe('XLSX active content security metadata', () => {
  it.each([
    [
      'xl/vbaProject.bin',
      'application/vnd.ms-office.vbaProject',
      'vba-project',
    ],
    [
      'xl/activeX/activeX1.bin',
      'application/vnd.ms-office.activeX',
      'active-x',
    ],
    [
      'xl/embeddings/oleObject1.bin',
      'application/vnd.openxmlformats-officedocument.oleObject',
      'ole-object',
    ],
    [
      'xl/embeddings/package1.bin',
      'application/vnd.openxmlformats-officedocument.embeddedPackage',
      'embedded-package',
    ],
    ['xl/embeddings/payload.exe', 'application/octet-stream', 'executable'],
    [
      'xl/customUI/customUI.xml',
      'application/vnd.ms-office.customUI+xml',
      'custom-ui',
    ],
    [
      'xl/webextensions/webextension1.xml',
      'application/vnd.ms-office.webextension+xml',
      'web-extension',
    ],
  ] as const)('classifies active content %s', (name, type, expected) => {
    expect(classifyXlsxActiveContent(name, type)).toBe(expected);
  });

  it.each([
    ['xl/vbaProject.bin', 'application/octet-stream', 'vba-project'],
    ['xl/normal.bin', 'application/vnd.ms-office.vbaProject', 'vba-project'],
    [
      'xl/normal.bin',
      'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
      'vba-project',
    ],
    ['xl/activeX/control.bin', 'application/octet-stream', 'active-x'],
    ['xl/normal.bin', 'application/vnd.ms-office.activeX', 'active-x'],
    [
      'xl/embeddings/package.bin',
      'application/octet-stream',
      'embedded-package',
    ],
    [
      'xl/normal.bin',
      'application/vnd.openxmlformats-officedocument.embeddedPackage',
      'embedded-package',
    ],
    ['xl/customUI/ribbon.xml', 'application/xml', 'custom-ui'],
    ['xl/normal.xml', 'application/vnd.ms-office.customUI+xml', 'custom-ui'],
    ['xl/webextensions/item.xml', 'application/xml', 'web-extension'],
    ['xl/taskpanes/taskpane.xml', 'application/xml', 'web-extension'],
    [
      'xl/normal.xml',
      'application/vnd.ms-office.webextension+xml',
      'web-extension',
    ],
    [
      'xl/normal.xml',
      'application/vnd.ms-office.taskpane+xml',
      'web-extension',
    ],
  ] as const)(
    'classifies independent active signal %#',
    (name, type, expected) => {
      expect(classifyXlsxActiveContent(name, type)).toBe(expected);
    },
  );

  it('does not classify ordinary XML, images, or slicer-cache XML', () => {
    expect(
      classifyXlsxActiveContent('xl/worksheets/sheet1.xml', 'application/xml'),
    ).toBeUndefined();
    expect(
      classifyXlsxActiveContent('xl/media/image1.png', 'image/png'),
    ).toBeUndefined();
    expect(
      classifyXlsxActiveContent(
        'xl/slicerCaches/slicerCache1.bin',
        'application/vnd.ms-excel.slicerCache+xml',
      ),
    ).toBeUndefined();
    expect(
      classifyXlsxActiveContent('xl/media/payload.exe.safe', undefined),
    ).toBeUndefined();
    expect(
      classifyXlsxActiveContent('xl/media/ordinary.bin', undefined),
    ).toBeUndefined();
  });

  it('creates stable structured security diagnostics', () => {
    expect(
      activeXlsxContentDiagnostic(
        { kind: 'active-x', part: 'xl/activeX/activeX1.bin' },
        'warning',
      ),
    ).toStrictEqual({
      code: 'security-rejected-content',
      message: 'XLSX active-x content was not loaded',
      part: 'xl/activeX/activeX1.bin',
      severity: 'warning',
    });
  });

  it('omits optional active payloads without reading them in tolerant mode', async () => {
    const source = await createIndependentXlsx({
      '[Content_Types].xml': contentTypes(`
        <Override PartName="/xl/activeX/activeX1.bin" ContentType="application/vnd.ms-office.activeX"/>
        <Override PartName="/xl/embeddings/oleObject1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
      `),
      'xl/activeX/activeX1.bin': new Uint8Array([0xff, 0x00, 0xfe]),
      'xl/embeddings/oleObject1.bin': new Uint8Array([0x00, 0xff, 0x01]),
    });
    const result = await parseXlsxWithDiagnostics(source);
    expect(result.document.sheets).toHaveLength(1);
    expect(result.diagnostics).toStrictEqual([
      {
        code: 'security-rejected-content',
        message: 'XLSX active-x content was not loaded',
        part: 'xl/activeX/activeX1.bin',
        severity: 'warning',
      },
      {
        code: 'security-rejected-content',
        message: 'XLSX ole-object content was not loaded',
        part: 'xl/embeddings/oleObject1.bin',
        severity: 'warning',
      },
    ]);
  });

  it('rejects the first canonical active part in strict mode', async () => {
    const source = await createIndependentXlsx({
      '[Content_Types].xml': contentTypes(`
        <Override PartName="/xl/activeX/activeX1.bin" ContentType="application/vnd.ms-office.activeX"/>
        <Override PartName="/xl/embeddings/oleObject1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
      `),
      'xl/activeX/activeX1.bin': 'not parsed',
      'xl/embeddings/oleObject1.bin': 'not parsed',
    });
    try {
      await parseXlsx(source, { errorMode: 'strict' });
    } catch (error) {
      expect(error).toBeInstanceOf(XlsxParseError);
      expect((error as XlsxParseError).diagnostic).toStrictEqual({
        code: 'security-rejected-content',
        message: 'XLSX active-x content was not loaded',
        part: 'xl/activeX/activeX1.bin',
        severity: 'error',
      });
      return;
    }
    throw new Error('Expected strict active-content rejection');
  });

  it('sorts active findings independently of ZIP insertion order', async () => {
    const source = await createIndependentXlsx({
      '[Content_Types].xml': contentTypes(`
        <Override PartName="/xl/z/late.bin" ContentType="application/vnd.ms-office.activeX"/>
        <Override PartName="/xl/a/early.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
      `),
      'xl/z/late.bin': 'inserted first',
      'xl/a/early.bin': 'inserted second',
    });
    const result = await parseXlsxWithDiagnostics(source);
    expect(result.diagnostics.map((item) => item.part)).toStrictEqual([
      'xl/a/early.bin',
      'xl/z/late.bin',
    ]);
  });

  it('does not classify digital-signature payloads as executable content', async () => {
    const source = await createIndependentXlsx({
      '[Content_Types].xml': contentTypes(
        '<Override PartName="/_xmlsignatures/payload.exe" ContentType="application/octet-stream"/>',
      ),
      '_xmlsignatures/payload.exe': 'signature bytes',
    });
    const result = await parseXlsxWithDiagnostics(source);
    expect(result.diagnostics).toStrictEqual([]);
  });

  it('does not inspect relationship parts even when an override claims an active content type', async () => {
    const source = await createIndependentXlsx({
      '[Content_Types].xml': contentTypes(
        '<Override PartName="/xl/activeX/_rels/control.bin.rels" ContentType="application/vnd.ms-office.activeX"/>',
      ),
      'xl/activeX/_rels/control.bin.rels': 'relationship metadata only',
    });
    const result = await parseXlsxWithDiagnostics(source);
    expect(result.diagnostics).toStrictEqual([]);
  });
});
