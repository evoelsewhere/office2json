import { XlsxParseError } from '../errors';
import { canonicalizeXlsxPartName } from './package-identity';

export const XLSX_CONTENT_TYPES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/content-types';

type XmlRecord = Record<string, unknown>;

export interface XlsxContentTypeTable {
  readonly defaults: ReadonlyMap<string, string>;
  readonly overrides: ReadonlyMap<string, string>;
  contentTypeFor(partName: string): string | undefined;
}

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function records(value: unknown): XmlRecord[] | undefined {
  if (value === undefined) return [];
  const values: unknown[] = Array.isArray(value) ? value : [value];
  const output: XmlRecord[] = [];
  for (const item of values) {
    const itemRecord = record(item);
    if (!itemRecord) return undefined;
    output.push(itemRecord);
  }
  return output;
}

function attributes(value: XmlRecord): XmlRecord {
  return record(value.attrs) ?? {};
}

function contentTypeFailure(message: string, cause?: unknown): never {
  throw new XlsxParseError(
    {
      code: 'invalid-document-structure',
      message,
      part: '[Content_Types].xml',
      severity: 'error',
    },
    { cause },
  );
}

function validExtension(value: unknown): value is string {
  if (typeof value !== 'string' || value.length === 0) return false;
  for (const character of value) {
    const code = character.codePointAt(0) ?? 0;
    const alphanumeric =
      (code >= 0x30 && code <= 0x39) ||
      (code >= 0x41 && code <= 0x5a) ||
      (code >= 0x61 && code <= 0x7a);
    if (!alphanumeric && character !== '_' && character !== '-') return false;
  }
  return true;
}

function validMimeType(value: unknown): value is string {
  if (typeof value !== 'string') return false;
  const slashIndex = value.indexOf('/');
  if (slashIndex <= 0 || slashIndex === value.length - 1) return false;
  for (const character of value) {
    const code = character.codePointAt(0) ?? 0;
    if (code <= 0x20 || code === 0x7f) return false;
  }
  return true;
}

function partExtension(partName: string): string | undefined {
  const slashIndex = partName.lastIndexOf('/');
  const dotIndex = partName.lastIndexOf('.');
  return dotIndex > slashIndex + 1
    ? partName.slice(dotIndex + 1).toLowerCase()
    : undefined;
}

export function parseXlsxContentTypes(value: unknown): XlsxContentTypeTable {
  const root = record(record(value)?.Types);
  if (!root || attributes(root).xmlns !== XLSX_CONTENT_TYPES_NAMESPACE) {
    contentTypeFailure(
      'Content types root is missing or has the wrong namespace',
    );
  }

  const defaultNodes = records(root.Default);
  const overrideNodes = records(root.Override);
  if (!defaultNodes || !overrideNodes) {
    contentTypeFailure('Content types contain an invalid entry collection');
  }

  const defaults = new Map<string, string>();
  for (const defaultNode of defaultNodes) {
    const attrs = attributes(defaultNode);
    if (!validExtension(attrs.Extension)) {
      contentTypeFailure('Content type default has an invalid extension');
    }
    if (!validMimeType(attrs.ContentType)) {
      contentTypeFailure('Content type default has an invalid MIME type');
    }
    const extension = attrs.Extension.toLowerCase();
    if (defaults.has(extension)) {
      contentTypeFailure('Content types contain a duplicate default extension');
    }
    defaults.set(extension, attrs.ContentType);
  }

  const overrides = new Map<string, string>();
  for (const overrideNode of overrideNodes) {
    const attrs = attributes(overrideNode);
    if (typeof attrs.PartName !== 'string' || attrs.PartName.length === 0) {
      contentTypeFailure('Content type override has an invalid part name');
    }
    let partName: string;
    try {
      partName = canonicalizeXlsxPartName(attrs.PartName);
    } catch (cause) {
      contentTypeFailure(
        'Content type override has an invalid part name',
        cause,
      );
    }
    if (!validMimeType(attrs.ContentType)) {
      contentTypeFailure('Content type override has an invalid MIME type');
    }
    if (overrides.has(partName)) {
      contentTypeFailure(
        'Content types contain a duplicate canonical part name',
      );
    }
    overrides.set(partName, attrs.ContentType);
  }

  return {
    defaults,
    overrides,
    contentTypeFor(partName: string): string | undefined {
      const canonical = canonicalizeXlsxPartName(partName);
      const override = overrides.get(canonical);
      if (override !== undefined) return override;
      const extension = partExtension(canonical);
      return extension === undefined ? undefined : defaults.get(extension);
    },
  };
}
