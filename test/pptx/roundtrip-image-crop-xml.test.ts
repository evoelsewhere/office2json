import { describe, expect, it } from 'vitest';

import { patchPptxImageCropXml } from '../../src/formats/pptx/roundtrip/image-crop-xml';
import type { PptxRoundTripSetImageCropOperation } from '../../src/formats/pptx/roundtrip/types';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MARKUP_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';

const CROP = { bottom: -20, left: 30, right: 0, top: 10.125 };

function operation(
  expectedCrop: PptxRoundTripSetImageCropOperation['expectedCrop'] = null,
  value: PptxRoundTripSetImageCropOperation['value'] = CROP,
): PptxRoundTripSetImageCropOperation {
  return {
    expectedCrop,
    id: 'set-image-crop-1',
    kind: 'set-image-crop',
    targetKey: 'slide-1-element-1',
    value,
  };
}

function slideXml(
  sourceRect = '',
  blip = '<a:blip r:embed="rIdImage"/>',
  extra = '',
): string {
  return (
    `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}" xmlns:mc="${MARKUP_NAMESPACE}">` +
    '<p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="2"/></p:nvPicPr>' +
    `<p:blipFill>${blip}${sourceRect}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr/>${extra}</p:pic></p:spTree></p:cSld></p:sld>`
  );
}

describe('PowerPoint native image crop XML patching', () => {
  it('adds a crop after the complete blip and preserves its children', () => {
    const input = slideXml(
      '',
      '<a:blip r:embed="rIdImage"><a:alphaModFix amt="50000"/></a:blip>',
    );
    const output = patchPptxImageCropXml(input, '2', operation());

    expect(output).toContain(
      '<a:alphaModFix amt="50000"/></a:blip><a:srcRect l="30000" t="10125" r="0" b="-20000"/>',
    );
    expect(output).toContain('<a:stretch><a:fillRect/></a:stretch>');
  });

  it('replaces signed and omitted source edges with canonical attributes', () => {
    const input = slideXml('<a:srcRect l="+30000" t="10125" b="-20000"/>');
    const output = patchPptxImageCropXml(
      input,
      '2',
      operation(CROP, {
        bottom: 5,
        left: 10,
        right: 15,
        top: 20,
      }),
    );

    expect(output).not.toContain('l="+30000"');
    expect(output).toContain(
      '<a:srcRect l="10000" t="20000" r="15000" b="5000"/>',
    );
  });

  it('removes an existing crop without changing the blip or stretch', () => {
    const input = slideXml('<a:srcRect l="30000" t="10125" r="0" b="-20000"/>');
    const output = patchPptxImageCropXml(input, '2', operation(CROP, null));

    expect(output).not.toContain('srcRect');
    expect(output).toContain(
      '<a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch>',
    );
  });

  it('supports namespace aliases and explicit empty source rectangles', () => {
    const input = slideXml(
      '<a:srcRect   l = \'30000\' t = "0" r = \'0\' b = "0"></a:srcRect>',
    )
      .replace(
        `xmlns:a="${DRAWING_NAMESPACE}"`,
        `xmlns:drawing="${DRAWING_NAMESPACE}"`,
      )
      .replaceAll('<a:', '<drawing:')
      .replaceAll('</a:', '</drawing:');
    const expected = { bottom: 0, left: 30, right: 0, top: 0 };

    expect(
      patchPptxImageCropXml(input, '2', operation(expected, CROP)),
    ).toContain('<drawing:srcRect l="30000" t="10125" r="0" b="-20000"/>');
  });

  it.each([
    [
      'negative boundary',
      '<a:srcRect l="-100000"/>',
      { bottom: 0, left: -100, right: 0, top: 0 },
    ],
    [
      'positive boundary with extension',
      '<a:srcRect l="100000" r="-1000"/>',
      { bottom: 0, left: 100, right: -1, top: 0 },
    ],
  ])('accepts the exact %s', (_name, sourceRect, expectedCrop) => {
    expect(
      patchPptxImageCropXml(
        slideXml(sourceRect),
        '2',
        operation(expectedCrop, CROP),
      ),
    ).toContain('l="30000"');
  });

  it.each(['bottom', 'left', 'right', 'top'] as const)(
    'checks the stale %s crop edge independently',
    (edge) => {
      const expected = { ...CROP, [edge]: CROP[edge] + 1 };
      expect(() =>
        patchPptxImageCropXml(
          slideXml('<a:srcRect l="30000" t="10125" r="0" b="-20000"/>'),
          '2',
          operation(expected, null),
        ),
      ).toThrow('does not match its preview precondition');
    },
  );

  it.each([
    ['stale precondition', '<a:srcRect l="10000"/>', CROP, 'does not match'],
    [
      'non-integer',
      '<a:srcRect l="1.5"/>',
      { bottom: 0, left: 0.0015, right: 0, top: 0 },
      'percentage is invalid',
    ],
    [
      'unsafe integer',
      '<a:srcRect l="100001"/>',
      { bottom: 0, left: 100.001, right: 0, top: 0 },
      'percentage is unsafe',
    ],
    [
      'unsafe negative integer',
      '<a:srcRect l="-100001"/>',
      { bottom: 0, left: -100.001, right: 0, top: 0 },
      'percentage is unsafe',
    ],
    [
      'unknown attribute',
      '<a:srcRect diagonal="1"/>',
      null,
      'attributes are unsafe',
    ],
    [
      'unsupported Unicode attribute',
      '<a:srcRect é="1"/>',
      null,
      'attributes are invalid',
    ],
    [
      'duplicate attribute',
      '<a:srcRect l="1" l="2"/>',
      null,
      'slide root is unsupported',
    ],
    [
      'collapsed region',
      '<a:srcRect l="60000" r="40000"/>',
      null,
      'no positive visible region',
    ],
    [
      'collapsed vertical region',
      '<a:srcRect t="60000" b="40000"/>',
      null,
      'no positive visible region',
    ],
  ])('rejects %s', (_name, sourceRect, expectedCrop, message) => {
    expect(() =>
      patchPptxImageCropXml(
        slideXml(sourceRect),
        '2',
        operation(expectedCrop as never),
      ),
    ).toThrow(message);
  });

  it.each([
    ['PresentationML extension', '<p:extLst/>'],
    ['DrawingML extension', '<a:extLst/>'],
    ['alternate content', '<mc:AlternateContent/>'],
  ])('rejects %s', (_name, extra) => {
    expect(() =>
      patchPptxImageCropXml(slideXml('', undefined, extra), '2', operation()),
    ).toThrow(
      'PowerPoint image crop target contains unsupported compatibility markup',
    );
  });

  it('rejects ambiguous fills, source rectangles, blips, and target ids', () => {
    expect(() =>
      patchPptxImageCropXml(
        slideXml().replace('</p:pic>', '<p:blipFill></p:blipFill></p:pic>'),
        '2',
        operation(),
      ),
    ).toThrow('requires exactly one picture fill');
    expect(() =>
      patchPptxImageCropXml(
        slideXml('<a:srcRect/><a:srcRect/>'),
        '2',
        operation(),
      ),
    ).toThrow('requires at most one source rectangle');
    expect(() =>
      patchPptxImageCropXml(
        slideXml('', '<a:blip/><a:blip/>'),
        '2',
        operation(),
      ),
    ).toThrow('requires one picture blip');
    expect(() => patchPptxImageCropXml(slideXml(), '7', operation())).toThrow(
      'requires one unique picture for id 7',
    );
  });

  it('rejects crop removal when the source has no crop', () => {
    expect(() =>
      patchPptxImageCropXml(slideXml(), '2', operation(null, null)),
    ).toThrow('crop removal has no source crop');
  });
});
