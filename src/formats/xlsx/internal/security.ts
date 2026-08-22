import type JSZip from 'jszip';

import type { XlsxDiagnostic } from '../types';
import type { XlsxWorkbookDiscovery } from './workbook-discovery';

export type XlsxActiveContentKind =
  | 'active-x'
  | 'custom-ui'
  | 'embedded-package'
  | 'executable'
  | 'ole-object'
  | 'vba-project'
  | 'web-extension';

export interface XlsxActiveContentFinding {
  kind: XlsxActiveContentKind;
  part: string;
}

function executableExtension(name: string): boolean {
  return /\.(?:bat|cmd|com|dll|exe|jar|js|msi|ps1|scr|vbs)$/iu.test(name);
}

export function classifyXlsxActiveContent(
  name: string,
  contentType: string | undefined,
): XlsxActiveContentKind | undefined {
  const foldedName = name.toLowerCase();
  const foldedType =
    contentType === undefined ? undefined : contentType.toLowerCase();
  if (foldedName.includes('vbaproject')) return 'vba-project';
  if (
    foldedType?.includes('vbaproject') === true ||
    foldedType?.includes('macroenabled') === true
  )
    return 'vba-project';
  if (foldedName.includes('/activex/')) return 'active-x';
  if (foldedType?.includes('activex') === true) return 'active-x';
  if (executableExtension(foldedName)) return 'executable';
  if (foldedType?.includes('oleobject') === true) return 'ole-object';
  if (foldedName.includes('/embeddings/')) return 'embedded-package';
  if (foldedType?.includes('embeddedpackage') === true)
    return 'embedded-package';
  if (foldedName.includes('/customui/')) return 'custom-ui';
  if (foldedType?.includes('customui') === true) return 'custom-ui';
  if (
    foldedName.includes('/webextensions/') ||
    foldedName.includes('/taskpanes/')
  )
    return 'web-extension';
  if (
    foldedType?.includes('webextension') === true ||
    foldedType?.includes('taskpane') === true
  )
    return 'web-extension';
  return undefined;
}

export function scanXlsxActiveContent(
  archive: JSZip,
  discovery: XlsxWorkbookDiscovery,
): XlsxActiveContentFinding[] {
  const output: XlsxActiveContentFinding[] = [];
  const names = Object.values(archive.files)
    .filter((entry) => !entry.dir)
    .map((entry) => entry.name)
    .sort();
  for (const name of names) {
    if (name.endsWith('.rels') || name.startsWith('_xmlsignatures/')) {
      continue;
    }
    const kind = classifyXlsxActiveContent(
      name,
      discovery.contentTypes.contentTypeFor(name),
    );
    if (kind) output.push({ kind, part: name });
  }
  return output;
}

export function activeXlsxContentDiagnostic(
  finding: XlsxActiveContentFinding,
  severity: 'error' | 'warning',
): XlsxDiagnostic {
  return {
    code: 'security-rejected-content',
    message: `XLSX ${finding.kind} content was not loaded`,
    part: finding.part,
    severity,
  };
}
