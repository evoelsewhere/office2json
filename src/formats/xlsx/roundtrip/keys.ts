import type { XlsxDocument } from '../types';
import { canonicalXlsxSha256 } from './digest';
import type { XlsxRoundTripDocument, XlsxRoundTripSheet } from './types';

export const XLSX_KEY_ALGORITHM_VERSION = 'xlsx-snapshot-key-v1' as const;

async function snapshotKey(kind: string, value: unknown): Promise<string> {
  return `xlsx:${kind}:${(await canonicalXlsxSha256(value)).slice(0, 32)}`;
}

export async function createXlsxRoundTripDocument(
  document: XlsxDocument,
  sourceSha256: string,
  supportProfileId: string,
): Promise<XlsxRoundTripDocument> {
  const key = await snapshotKey('workbook', {
    sourceSha256,
    supportProfileId,
  });
  const sheets: XlsxRoundTripSheet[] = await Promise.all(
    document.sheets.map(async (sheet) => ({
      ...sheet,
      key: await snapshotKey('sheet', {
        index: sheet.index,
        kind: sheet.kind,
        sourceSha256,
        supportProfileId,
      }),
    })),
  );
  return { ...document, key, sheets };
}
