import { decodedBase64ByteLength } from '../../../common/binary/base64';
import { XlsxParseError } from '../errors';
import type { XlsxProtectionCredential } from '../types';

type XmlRecord = Record<string, unknown>;

export interface XlsxProtectionCredentialFields {
  algorithmName: string;
  hashValue: string;
  legacyHash: string;
  saltValue: string;
  spinCount: string;
}

export interface XlsxParsedProtectionCredential {
  credential?: XlsxProtectionCredential;
  textCharacters: number;
}

function valueFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function legacyHash(
  value: unknown,
  part: string,
  context: string,
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !/^[0-9A-Fa-f]{4}$/u.test(value)) {
    valueFailure(`${context} legacy password hash is invalid`, part);
  }
  return value.toUpperCase();
}

function positiveUnsignedInteger(
  value: unknown,
  part: string,
  context: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !/^[1-9]\d*$/u.test(value)) {
    valueFailure(`${context} spin count is invalid`, part);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > 0xffff_ffff) {
    valueFailure(`${context} spin count is invalid`, part);
  }
  return parsed;
}

function algorithmName(
  value: unknown,
  part: string,
  context: string,
): string | undefined {
  if (value === undefined) return undefined;
  if (
    typeof value !== 'string' ||
    !/^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$/u.test(value)
  ) {
    valueFailure(`${context} hash algorithm is invalid`, part);
  }
  return value;
}

function canonicalBase64(
  value: unknown,
  part: string,
  message: string,
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') valueFailure(message, part);
  try {
    if (decodedBase64ByteLength(value) === 0) valueFailure(message, part);
  } catch {
    valueFailure(message, part);
  }
  return value;
}

export function parseXlsxProtectionCredential(
  attrs: XmlRecord,
  fields: XlsxProtectionCredentialFields,
  part: string,
  context: string,
): XlsxParsedProtectionCredential {
  const legacy = legacyHash(attrs[fields.legacyHash], part, context);
  const algorithm = algorithmName(attrs[fields.algorithmName], part, context);
  const hash = canonicalBase64(
    attrs[fields.hashValue],
    part,
    `${context} hash value is invalid`,
  );
  const salt = canonicalBase64(
    attrs[fields.saltValue],
    part,
    `${context} salt value is invalid`,
  );
  const spin = positiveUnsignedInteger(attrs[fields.spinCount], part, context);
  const strongCount = [algorithm, hash, salt, spin].filter(
    (value) => value !== undefined,
  ).length;
  if (strongCount !== 0 && strongCount !== 4) {
    valueFailure(`${context} strong hash metadata is incomplete`, part);
  }
  const credential: XlsxProtectionCredential = {
    ...(legacy === undefined ? {} : { legacyHash: legacy }),
    ...(strongCount === 0
      ? {}
      : {
          strongHash: {
            algorithmName: algorithm!,
            hashValue: hash!,
            saltValue: salt!,
            spinCount: spin!,
          },
        }),
  };
  return {
    ...(legacy === undefined && strongCount === 0 ? {} : { credential }),
    textCharacters:
      (legacy?.length ?? 0) +
      (algorithm?.length ?? 0) +
      (hash?.length ?? 0) +
      (salt?.length ?? 0),
  };
}
