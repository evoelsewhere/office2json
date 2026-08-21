import { describe, expect, it } from 'vitest';

import { parsePptx, readPptxRoundTrip } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  OFFICE_REL_NS,
  OFFICE_REL_TYPE,
  PACKAGE_REL_NS,
  PRESENTATION_NS,
} from './pptx-package';

function pictureXml(id: number, crop: string): string {
  return `<p:pic>
    <p:nvPicPr>
      <p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/>
    </p:nvPicPr>
    <p:blipFill>
      <a:blip r:embed="rIdImage"/>
      <a:srcRect ${crop}/>
      <a:stretch><a:fillRect/></a:stretch>
    </p:blipFill>
    <p:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    </p:spPr>
  </p:pic>`;
}

describe('PowerPoint picture crops through the public API', () => {
  it('preserves canonical crop percentages and omits malformed attributes', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            ${pictureXml(700, 't="10000" b="-20000" l="+30000" r="0"')}
            ${pictureXml(
              701,
              't="junk10000" b="10000junk" l="Infinity" r="1.5"',
            )}
          </p:spTree></p:cSld>
        </p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdImage" Type="${OFFICE_REL_TYPE}image" Target="../media/crop.png"/>
        </Relationships>`,
    });

    const result = await parsePptx(input, {
      errorMode: 'strict',
      imageMode: 'none',
    });
    const byId = Object.fromEntries(
      (result.slides[0]?.elements ?? []).map((element) => [
        element.id,
        element,
      ]),
    );

    expect(byId['700']).toMatchObject({
      rect: { b: -20, l: 30, r: 0, t: 10 },
      type: 'image',
    });
    expect(byId['701']).toMatchObject({ type: 'image' });
    expect(byId['701']).not.toHaveProperty('rect');
    expect(result).not.toHaveProperty('diagnostics');

    const snapshot = await readPptxRoundTrip(input);
    expect(snapshot.document.slides[0]?.elements).toMatchObject([
      {
        crop: { bottom: -20, left: 30, right: 0, top: 10 },
        type: 'image',
      },
      { type: 'image' },
    ]);
    expect(snapshot.document.slides[0]?.elements[1]).not.toHaveProperty('crop');
  });
});
