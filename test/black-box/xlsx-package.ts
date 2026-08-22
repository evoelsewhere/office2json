import JSZip from 'jszip';

export type XlsxBlackBoxPart = string | Uint8Array | null;
export type XlsxBlackBoxOverrides = Record<string, XlsxBlackBoxPart>;

export interface XlsxBlackBoxPackageOptions {
  compression?: 'DEFLATE' | 'STORE';
}

export const XLSX_CONTENT_TYPES_NS =
  'http://schemas.openxmlformats.org/package/2006/content-types';
export const XLSX_PACKAGE_REL_NS =
  'http://schemas.openxmlformats.org/package/2006/relationships';
export const XLSX_OFFICE_REL_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const XLSX_SPREADSHEET_NS =
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

export const XLSX_OFFICE_REL_TYPE = `${XLSX_OFFICE_REL_NS}/`;

export function independentWorksheet(sheetData: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
    <worksheet xmlns="${XLSX_SPREADSHEET_NS}">
      <dimension ref="A1:C3"/>
      <sheetData>${sheetData}</sheetData>
    </worksheet>`;
}

export function independentWorkbook(sheetDeclarations: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
    <workbook xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}">
      <workbookPr date1904="0"/>
      <bookViews><workbookView activeTab="0"/></bookViews>
      <sheets>${sheetDeclarations}</sheets>
      <calcPr calcMode="auto" fullCalcOnLoad="0" forceFullCalc="0"/>
    </workbook>`;
}

const BASE_PARTS: Readonly<Record<string, string>> = {
  '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="${XLSX_CONTENT_TYPES_NS}">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    </Types>`,
  '_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
      <Relationship Id="rIdWorkbook" Type="${XLSX_OFFICE_REL_TYPE}officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`,
  'xl/workbook.xml': independentWorkbook(
    '<sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/>',
  ),
  'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
      <Relationship Id="rIdSheet1" Type="${XLSX_OFFICE_REL_TYPE}worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rIdStyles" Type="${XLSX_OFFICE_REL_TYPE}styles" Target="styles.xml"/>
      <Relationship Id="rIdSharedStrings" Type="${XLSX_OFFICE_REL_TYPE}sharedStrings" Target="sharedStrings.xml"/>
    </Relationships>`,
  'xl/worksheets/sheet1.xml': independentWorksheet(`
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="B2"><v>42</v></c></row>
    <row r="3"><c r="C3" t="b"><v>1</v></c></row>
  `),
  'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8"?>
    <sst xmlns="${XLSX_SPREADSHEET_NS}" count="1" uniqueCount="1">
      <si><t>Black box</t></si>
    </sst>`,
  'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8"?>
    <styleSheet xmlns="${XLSX_SPREADSHEET_NS}">
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
      <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
      <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
    </styleSheet>`,
};

export async function createIndependentXlsx(
  overrides: XlsxBlackBoxOverrides = {},
  options: XlsxBlackBoxPackageOptions = {},
): Promise<Uint8Array> {
  const zip = new JSZip();
  const parts: Record<string, XlsxBlackBoxPart> = {
    ...BASE_PARTS,
    ...overrides,
  };

  for (const [name, content] of Object.entries(parts)) {
    if (content !== null) zip.file(name, content);
  }

  return zip.generateAsync({
    compression: options.compression ?? 'DEFLATE',
    type: 'uint8array',
  });
}
