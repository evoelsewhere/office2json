import { describe, expect, it } from 'vitest';

import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import { validateXlsxOperationStyle } from '../../src/formats/xlsx/roundtrip/style-validation';

function capture(value: unknown): XlsxWriteError {
  try {
    validateXlsxOperationStyle(value, 'style-op');
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected style validation to fail');
}

describe('XLSX operation style validation', () => {
  it('normalizes the complete portable style contract', () => {
    const style = {
      alignment: {
        horizontal: 'distributed',
        indent: 250,
        justifyLastLine: true,
        readingOrder: 'left-to-right',
        relativeIndent: -15,
        shrinkToFit: true,
        textRotation: 255,
        vertical: 'justify',
        wrapText: true,
      },
      border: {
        bottom: { color: { kind: 'automatic' }, style: 'thin' },
        diagonal: { color: { argb: 'FFAABBCC', kind: 'rgb', tint: -1 } },
        diagonalDown: true,
        diagonalUp: true,
        end: { style: 'dashDot' },
        horizontal: { style: 'dashDotDot' },
        left: { style: 'dashed' },
        outline: false,
        right: { style: 'dotted' },
        start: { style: 'double' },
        top: { style: 'hair' },
        vertical: { style: 'medium' },
      },
      checkbox: true,
      fill: {
        backgroundColor: { index: 65, kind: 'indexed', tint: 1 },
        foregroundColor: { index: 11, kind: 'theme', tint: -0.5 },
        kind: 'pattern',
        pattern: 'gray0625',
      },
      font: {
        bold: true,
        charset: 255,
        color: { argb: 'FF000000', kind: 'rgb' },
        condense: true,
        extend: true,
        family: 5,
        italic: true,
        name: 'Agent',
        outline: true,
        scheme: 'major',
        shadow: true,
        size: 409,
        strike: true,
        underline: 'single-accounting',
        verticalAlignment: 'subscript',
      },
      numberFormat: '0.0000',
      protection: { hidden: true, locked: false },
    } as const;
    expect(validateXlsxOperationStyle(style, 'style-op')).toEqual(style);
    expect(validateXlsxOperationStyle({}, 'style-op')).toEqual({});
  });

  it('normalizes linear and path gradients without negative zero', () => {
    expect(
      validateXlsxOperationStyle(
        { alignment: { textRotation: 180 } },
        'style-op',
      ),
    ).toEqual({ alignment: { textRotation: 180 } });
    expect(
      validateXlsxOperationStyle(
        {
          fill: {
            angle: 360,
            kind: 'gradient',
            stops: [
              { color: { argb: 'FF000000', kind: 'rgb' }, position: -0 },
              { color: { index: 1, kind: 'theme' }, position: 1 },
            ],
            type: 'linear',
          },
        },
        'style-op',
      ),
    ).toEqual({
      fill: {
        angle: 360,
        kind: 'gradient',
        stops: [
          { color: { argb: 'FF000000', kind: 'rgb' }, position: 0 },
          { color: { index: 1, kind: 'theme' }, position: 1 },
        ],
        type: 'linear',
      },
    });
    expect(
      validateXlsxOperationStyle(
        {
          fill: {
            bottom: 1,
            kind: 'gradient',
            left: 0.25,
            right: 0.5,
            stops: [
              { color: { index: 0, kind: 'indexed' }, position: 0 },
              { color: { kind: 'automatic' }, position: 1 },
            ],
            top: 0.75,
            type: 'path',
          },
        },
        'style-op',
      ).fill,
    ).toMatchObject({ bottom: 1, left: 0.25, right: 0.5, top: 0.75 });
  });

  it.each([
    [null, 'XLSX set-cell-style style shape is invalid'],
    [[], 'XLSX set-cell-style style shape is invalid'],
    [{ extra: true }, 'XLSX set-cell-style style shape is invalid'],
    [{ checkbox: false }, 'XLSX checkbox style flag is invalid'],
    [{ numberFormat: '' }, 'XLSX style number format is invalid'],
    [{ alignment: {} }, 'XLSX style alignment must be normalized'],
    [
      { alignment: { horizontal: 'bad' } },
      'XLSX style horizontal alignment is invalid',
    ],
    [{ alignment: { indent: 0 } }, 'XLSX style alignment indent is invalid'],
    [
      { alignment: { justifyLastLine: false } },
      'XLSX style justify-last-line flag is invalid',
    ],
    [
      { alignment: { readingOrder: 'bad' } },
      'XLSX style reading order is invalid',
    ],
    [
      { alignment: { relativeIndent: 0 } },
      'XLSX style relative indent is invalid',
    ],
    [
      { alignment: { shrinkToFit: false } },
      'XLSX style shrink-to-fit flag is invalid',
    ],
    [
      { alignment: { textRotation: 181 } },
      'XLSX style text rotation is invalid',
    ],
    [
      { alignment: { vertical: 'bottom' } },
      'XLSX style vertical alignment is invalid',
    ],
    [
      { alignment: { wrapText: false } },
      'XLSX style wrap-text flag is invalid',
    ],
    [{ border: {} }, 'XLSX style border must be normalized'],
    [{ border: { bottom: {} } }, 'XLSX style border side must be normalized'],
    [
      { border: { bottom: { extra: true } } },
      'XLSX style border-side shape is invalid',
    ],
    [
      { border: { bottom: { style: 'bad' } } },
      'XLSX style border-side kind is invalid',
    ],
    [
      { border: { diagonalDown: false } },
      'XLSX style border diagonalDown flag is invalid',
    ],
    [
      { border: { outline: true } },
      'XLSX style border outline flag is invalid',
    ],
    [{ fill: {} }, 'XLSX style fill kind is invalid'],
    [
      { fill: { kind: 'pattern', pattern: 'bad' } },
      'XLSX pattern fill type is invalid',
    ],
    [
      { fill: { kind: 'gradient', stops: [], type: 'linear' } },
      'XLSX gradient fill stops are invalid',
    ],
    [
      {
        fill: {
          angle: 1,
          kind: 'gradient',
          stops: [
            { color: { kind: 'automatic' }, position: 0 },
            { color: { kind: 'automatic' }, position: 1 },
          ],
          type: 'path',
        },
      },
      'XLSX path gradient angle is invalid',
    ],
    [
      {
        fill: {
          kind: 'gradient',
          left: 0.5,
          stops: [
            { color: { kind: 'automatic' }, position: 0 },
            { color: { kind: 'automatic' }, position: 1 },
          ],
          type: 'linear',
        },
      },
      'XLSX linear gradient path bounds are invalid',
    ],
    [
      {
        fill: {
          kind: 'gradient',
          stops: [
            { color: { kind: 'automatic' }, position: 0.5 },
            { color: { kind: 'automatic' }, position: 0.5 },
          ],
          type: 'linear',
        },
      },
      'XLSX gradient stop positions are out of order',
    ],
    [{ font: {} }, 'XLSX style font must be normalized'],
    [{ font: { bold: false } }, 'XLSX style font bold flag is invalid'],
    [{ font: { charset: 256 } }, 'XLSX style font charset is invalid'],
    [{ font: { family: 6 } }, 'XLSX style font family is invalid'],
    [{ font: { name: '' } }, 'XLSX style font name is invalid'],
    [{ font: { scheme: 'none' } }, 'XLSX style font scheme is invalid'],
    [{ font: { size: 0 } }, 'XLSX style font size is invalid'],
    [{ font: { underline: 'none' } }, 'XLSX style font underline is invalid'],
    [
      { font: { verticalAlignment: 'baseline' } },
      'XLSX style font vertical alignment is invalid',
    ],
    [{ protection: {} }, 'XLSX style protection must be normalized'],
    [
      { protection: { hidden: false } },
      'XLSX style protection hidden flag is invalid',
    ],
    [
      { protection: { locked: true } },
      'XLSX style protection locked flag is invalid',
    ],
  ] as const)('rejects non-normalized style %#', (value, message) => {
    expect(capture(value).diagnostic).toMatchObject({
      code: 'invalid-roundtrip-json',
      message,
      operationId: 'style-op',
    });
  });

  it.each([
    [1, 'XLSX set-cell-style style shape is invalid'],
    [undefined, 'XLSX set-cell-style style shape is invalid'],
    [Object.create(null), 'XLSX set-cell-style style shape is invalid'],
    [{ alignment: null }, 'XLSX style alignment shape is invalid'],
    [{ alignment: { extra: true } }, 'XLSX style alignment shape is invalid'],
    [
      { alignment: { relativeIndent: -16 } },
      'XLSX style relative indent is invalid',
    ],
    [{ alignment: { textRotation: 0 } }, 'XLSX style text rotation is invalid'],
    [{ border: null }, 'XLSX style border shape is invalid'],
    [{ border: { extra: true } }, 'XLSX style border shape is invalid'],
    [{ border: { bottom: null } }, 'XLSX style border-side shape is invalid'],
    [{ fill: null }, 'XLSX style fill shape is invalid'],
    [
      { fill: { extra: true, kind: 'pattern', pattern: 'solid' } },
      'XLSX pattern fill shape is invalid',
    ],
    [
      {
        fill: {
          extra: true,
          kind: 'gradient',
          stops: [
            { color: { kind: 'automatic' }, position: 0 },
            { color: { kind: 'automatic' }, position: 1 },
          ],
          type: 'linear',
        },
      },
      'XLSX gradient fill shape is invalid',
    ],
    [
      {
        fill: {
          kind: 'gradient',
          stops: [
            { color: { kind: 'automatic' }, position: 0 },
            { color: { kind: 'automatic' }, position: 1 },
          ],
          type: 'bad',
        },
      },
      'XLSX gradient fill type is invalid',
    ],
    [
      {
        fill: {
          angle: 0,
          kind: 'gradient',
          stops: [
            { color: { kind: 'automatic' }, position: 0 },
            { color: { kind: 'automatic' }, position: 1 },
          ],
          type: 'linear',
        },
      },
      'XLSX gradient angle is invalid',
    ],
    [
      {
        fill: {
          angle: 361,
          kind: 'gradient',
          stops: [
            { color: { kind: 'automatic' }, position: 0 },
            { color: { kind: 'automatic' }, position: 1 },
          ],
          type: 'linear',
        },
      },
      'XLSX gradient angle is invalid',
    ],
    [
      {
        fill: {
          kind: 'gradient',
          left: 0,
          stops: [
            { color: { kind: 'automatic' }, position: 0 },
            { color: { kind: 'automatic' }, position: 1 },
          ],
          type: 'path',
        },
      },
      'XLSX gradient path bound is invalid',
    ],
    [
      {
        fill: {
          kind: 'gradient',
          stops: [null, { color: { kind: 'automatic' }, position: 1 }],
          type: 'linear',
        },
      },
      'XLSX gradient stop shape is invalid',
    ],
    [
      {
        fill: {
          kind: 'gradient',
          stops: [
            {
              color: { kind: 'automatic' },
              extra: true,
              position: 0,
            },
            { color: { kind: 'automatic' }, position: 1 },
          ],
          type: 'linear',
        },
      },
      'XLSX gradient stop shape is invalid',
    ],
    [
      {
        fill: {
          kind: 'gradient',
          stops: [
            { position: 0 },
            { color: { kind: 'automatic' }, position: 1 },
          ],
          type: 'linear',
        },
      },
      'XLSX gradient stop shape is invalid',
    ],
    [
      {
        fill: {
          kind: 'gradient',
          stops: [
            { color: { kind: 'automatic' } },
            { color: { kind: 'automatic' }, position: 1 },
          ],
          type: 'linear',
        },
      },
      'XLSX gradient stop shape is invalid',
    ],
    [
      {
        fill: {
          kind: 'gradient',
          stops: [
            { color: { kind: 'automatic' }, position: -1 },
            { color: { kind: 'automatic' }, position: 1 },
          ],
          type: 'linear',
        },
      },
      'XLSX gradient stop position is invalid',
    ],
    [{ font: null }, 'XLSX style font shape is invalid'],
    [{ font: { extra: true } }, 'XLSX style font shape is invalid'],
    [{ font: { size: 410 } }, 'XLSX style font size is invalid'],
    [{ protection: null }, 'XLSX style protection shape is invalid'],
    [{ protection: { extra: true } }, 'XLSX style protection shape is invalid'],
  ] as const)('rejects exact nested style shape %#', (value, message) => {
    expect(capture(value).diagnostic.message).toBe(message);
  });

  it.each([
    [
      { kind: 'automatic', tint: 0.1 },
      'XLSX automatic style color shape is invalid',
    ],
    [
      { argb: 'ffaabbcc', kind: 'rgb' },
      'XLSX RGB style color value is invalid',
    ],
    [
      { argb: 'XFFAABBCC', kind: 'rgb' },
      'XLSX RGB style color value is invalid',
    ],
    [
      { argb: 'FFAABBCCX', kind: 'rgb' },
      'XLSX RGB style color value is invalid',
    ],
    [
      { argb: 'FFAABBCC', extra: true, kind: 'rgb' },
      'XLSX RGB style color shape is invalid',
    ],
    [
      { argb: 'FFAABBCC', kind: 'rgb', tint: 0 },
      'XLSX RGB style color tint is invalid',
    ],
    [
      { extra: true, index: 1, kind: 'indexed' },
      'XLSX indexed style color shape is invalid',
    ],
    [{ index: 12, kind: 'theme' }, 'XLSX style color index is invalid'],
    [{ index: 66, kind: 'indexed' }, 'XLSX style color index is invalid'],
    [{ index: 1, kind: 'theme', tint: 0 }, 'XLSX style color tint is invalid'],
    [
      { index: 1, kind: 'theme', tint: Infinity },
      'XLSX style color tint is invalid',
    ],
    [{ kind: 'bad' }, 'XLSX style color kind is invalid'],
    [{ kind: 1 }, 'XLSX style color kind is invalid'],
    [null, 'XLSX style color shape is invalid'],
    [1, 'XLSX style color shape is invalid'],
  ] as const)('rejects invalid color %#', (invalidColor, message) => {
    expect(capture({ font: { color: invalidColor } }).diagnostic.message).toBe(
      message,
    );
  });
});
