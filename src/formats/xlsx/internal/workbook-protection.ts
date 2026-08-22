import { XlsxParseError } from '../errors';
import type { XlsxWorkbookProtection } from '../types';
import { parseXlsxProtectionCredential } from './protection-hash';

type XmlRecord = Record<string, unknown>;

function structureFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-structure',
    message,
    part,
    severity: 'error',
  });
}

function valueFailure(message: string, part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function record(value: unknown): XmlRecord | undefined {
  return Object.prototype.toString.call(value) === '[object Object]'
    ? (value as XmlRecord)
    : undefined;
}

function optionalBoolean(
  value: unknown,
  part: string,
  message: string,
): boolean {
  if (value === undefined || value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  valueFailure(message, part);
}

export interface XlsxParsedWorkbookProtection {
  protection?: XlsxWorkbookProtection;
  textCharacters: number;
}

export function parseXlsxWorkbookProtection(
  value: unknown,
  part: string,
): XlsxParsedWorkbookProtection {
  if (value === undefined) return { textCharacters: 0 };
  const node = record(value);
  if (!node) {
    structureFailure('Workbook protection is invalid', part);
  }
  const attrs = record(node.attrs) ?? {};
  const workbook = parseXlsxProtectionCredential(
    attrs,
    {
      algorithmName: 'workbookAlgorithmName',
      hashValue: 'workbookHashValue',
      legacyHash: 'workbookPassword',
      saltValue: 'workbookSaltValue',
      spinCount: 'workbookSpinCount',
    },
    part,
    'Workbook protection',
  );
  const revisions = parseXlsxProtectionCredential(
    attrs,
    {
      algorithmName: 'revisionsAlgorithmName',
      hashValue: 'revisionsHashValue',
      legacyHash: 'revisionsPassword',
      saltValue: 'revisionsSaltValue',
      spinCount: 'revisionsSpinCount',
    },
    part,
    'Revision protection',
  );
  return {
    protection: {
      lockRevisions: optionalBoolean(
        attrs.lockRevision,
        part,
        'Workbook revision-lock flag is invalid',
      ),
      lockStructure: optionalBoolean(
        attrs.lockStructure,
        part,
        'Workbook structure-lock flag is invalid',
      ),
      lockWindows: optionalBoolean(
        attrs.lockWindows,
        part,
        'Workbook window-lock flag is invalid',
      ),
      ...(revisions.credential === undefined
        ? {}
        : { revisionsCredential: revisions.credential }),
      ...(workbook.credential === undefined
        ? {}
        : { workbookCredential: workbook.credential }),
    },
    textCharacters: workbook.textCharacters + revisions.textCharacters,
  };
}
