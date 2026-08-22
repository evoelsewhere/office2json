import { canonicalXlsxJson } from './canonical-json';

function hexadecimal(bytes: Uint8Array): string {
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join(
    '',
  );
}

export async function sha256XlsxBytes(bytes: Uint8Array): Promise<string> {
  const owned = bytes.slice();
  const digest = await globalThis.crypto.subtle.digest('SHA-256', owned);
  return hexadecimal(new Uint8Array(digest));
}

export function sha256XlsxText(value: string): Promise<string> {
  return sha256XlsxBytes(new TextEncoder().encode(value));
}

export function canonicalXlsxSha256(value: unknown): Promise<string> {
  return sha256XlsxText(canonicalXlsxJson(value));
}
