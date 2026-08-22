import fc from 'fast-check';
import { describe, expect, it } from 'vitest';

import {
  canonicalizeXlsxPartName,
  getXlsxRelationshipPartName,
  resolveXlsxPartTarget,
  resolveXlsxRootTarget,
} from '../../src/formats/xlsx/internal/package-identity';

function captureError(action: () => unknown): TypeError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(TypeError);
    return error as TypeError;
  }
  throw new Error('Expected XLSX package identity validation to fail');
}

function expectIdentityError(action: () => unknown, reason: string): void {
  expect(captureError(action).message).toBe(
    `Invalid XLSX package part name: ${reason}`,
  );
}

describe('XLSX package part identity', () => {
  it.each([
    ['xl/workbook.xml', 'xl/workbook.xml'],
    ['/xl/workbook.xml', 'xl/workbook.xml'],
    ['xl/shared%53trings.xml', 'xl/sharedStrings.xml'],
    ['xl/My%20Sheet.xml', 'xl/My Sheet.xml'],
    ['custom/数据.xml', 'custom/数据.xml'],
    ['folder/http:book.xml', 'folder/http:book.xml'],
    ['1http:book.xml', '1http:book.xml'],
  ])('canonicalizes %s', (input, expected) => {
    expect(canonicalizeXlsxPartName(input)).toBe(expected);
  });

  it.each([
    ['', 'empty rooted name'],
    ['/', 'empty rooted name'],
    ['//xl/workbook.xml', 'empty rooted name'],
    ['xl//workbook.xml', 'empty path segment'],
    ['xl/workbook.xml/', 'empty path segment'],
    ['xl/./workbook.xml', 'dot segment'],
    ['xl/../workbook.xml', 'dot segment'],
    ['xl/%2e/workbook.xml', 'dot segment'],
    ['xl/%2E%2E/workbook.xml', 'dot segment'],
    ['xl%2fworkbook.xml', 'encoded path delimiter'],
    ['xl%5Cworkbook.xml', 'encoded path delimiter'],
    ['xl\\workbook.xml', 'backslash delimiter'],
    ['xl/workbook.xml?query', 'query or fragment'],
    ['xl/workbook.xml#fragment', 'query or fragment'],
    ['xl/%00.xml', 'control character'],
    [`xl/${String.fromCodePoint(0x1f)}.xml`, 'control character'],
    [`xl/${String.fromCodePoint(0x7f)}.xml`, 'control character'],
    ['a:book.xml', 'URI scheme'],
    ['ab:book.xml', 'URI scheme'],
    ['a1:book.xml', 'URI scheme'],
    ['x+1.2-3:book.xml', 'URI scheme'],
  ])('rejects ambiguous or unsafe part name %s', (name, reason) => {
    expectIdentityError(() => canonicalizeXlsxPartName(name), reason);
  });

  it('preserves the percent-decoding cause', () => {
    const error = captureError(() => canonicalizeXlsxPartName('xl/%GG.xml'));

    expect(error.message).toBe(
      'Invalid XLSX package part name: percent encoding',
    );
    expect(error.cause).toBeInstanceOf(URIError);
  });

  it('keeps valid case-sensitive identities distinct', () => {
    expect(canonicalizeXlsxPartName('xl/Sheet.xml')).not.toBe(
      canonicalizeXlsxPartName('xl/sheet.xml'),
    );
  });

  it.each([
    ['xl/workbook.xml', 'worksheets/sheet1.xml', 'xl/worksheets/sheet1.xml'],
    [
      'xl/worksheets/sheet1.xml',
      '../drawings/drawing1.xml',
      'xl/drawings/drawing1.xml',
    ],
    ['custom/books/workbook.xml', '/xl/styles.xml', 'xl/styles.xml'],
    [
      'xl/workbook.xml',
      'worksheets/My%20Sheet.xml',
      'xl/worksheets/My Sheet.xml',
    ],
  ])('resolves %s + %s', (owner, target, expected) => {
    expect(resolveXlsxPartTarget(owner, target)).toBe(expected);
  });

  it.each([
    ['', 'empty or padded relationship target'],
    [' sheet.xml', 'empty or padded relationship target'],
    ['sheet.xml ', 'empty or padded relationship target'],
    ['../../outside.xml', 'relationship target escapes package root'],
    ['https://example.com/book.xml', 'external URI scheme'],
    ['//server/share.xml', 'network-path reference'],
    ['worksheets/%2e%2e/styles.xml', 'dot segment'],
    ['worksheets%2fsheet1.xml', 'encoded path delimiter'],
    ['sheet.xml?version=1', 'query or fragment'],
    ['sheet.xml#cell', 'query or fragment'],
    ['./sheet.xml', 'dot target segment'],
    ['..', 'empty relationship target'],
  ])('rejects unsafe target %s', (target, reason) => {
    expectIdentityError(
      () => resolveXlsxPartTarget('xl/workbook.xml', target),
      reason,
    );
  });

  it.each([
    ['xl/workbook.xml', 'xl/workbook.xml'],
    ['/custom/books/workbook.xml', 'custom/books/workbook.xml'],
    ['xl/My%20Workbook.xml', 'xl/My Workbook.xml'],
  ])('resolves package-root target %s', (target, expected) => {
    expect(resolveXlsxRootTarget(target)).toBe(expected);
  });

  it.each([
    ['', 'empty or padded relationship target'],
    [' xl/workbook.xml', 'empty or padded relationship target'],
    ['xl/workbook.xml ', 'empty or padded relationship target'],
    ['../workbook.xml', 'dot segment'],
    ['https://example.com/book.xml', 'external URI scheme'],
    ['//server/share.xml', 'network-path reference'],
    ['xl%2fworkbook.xml', 'encoded path delimiter'],
  ])('rejects unsafe package-root target %s', (target, reason) => {
    expectIdentityError(() => resolveXlsxRootTarget(target), reason);
  });

  it('derives owner-scoped relationship part names', () => {
    expect(getXlsxRelationshipPartName('xl/workbook.xml')).toBe(
      'xl/_rels/workbook.xml.rels',
    );
    expect(getXlsxRelationshipPartName('xl/worksheets/sheet1.xml')).toBe(
      'xl/worksheets/_rels/sheet1.xml.rels',
    );
    expect(getXlsxRelationshipPartName('workbook.xml')).toBe(
      '_rels/workbook.xml.rels',
    );
  });

  it('is idempotent for generated valid part names', () => {
    fc.assert(
      fc.property(
        fc.array(fc.stringMatching(/^[A-Za-z0-9_-]{1,12}$/), {
          minLength: 1,
          maxLength: 8,
        }),
        (segments) => {
          const value = segments.join('/');
          const canonical = canonicalizeXlsxPartName(value);
          expect(canonicalizeXlsxPartName(canonical)).toBe(canonical);
        },
      ),
      { numRuns: 200, seed: 2_026_081_601 },
    );
  });
});
