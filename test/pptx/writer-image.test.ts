import { describe, expect, it } from 'vitest';

import type {
  PptxSceneImageElement,
  PptxSceneTransform,
} from '../../src/formats/pptx/scene-types';
import { serializePicture } from '../../src/formats/pptx/writer/image';

const TRANSFORM: PptxSceneTransform = {
  height: 40,
  width: 100,
  x: 10,
  y: 20,
};

function imageElement(): PptxSceneImageElement {
  return {
    authored: {},
    key: 'image-1',
    mediaKey: 'media-1',
    resolved: { hidden: false },
    type: 'image',
  };
}

describe('native PowerPoint image serialization', () => {
  it('serializes a complete picture with an embedded relationship', () => {
    expect(serializePicture(imageElement(), TRANSFORM, 2, 'rId2')).toBe(
      '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 2"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="127000" y="254000"/><a:ext cx="1270000" cy="508000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>',
    );
  });

  it('escapes metadata, relationship ids, and authored visibility', () => {
    const element = imageElement();
    element.name = `Image <&"' _x0041_`;
    element.description = `Description <&"'`;
    element.title = `Title <&"'`;
    element.authored.hidden = true;
    const xml = serializePicture(element, TRANSFORM, 7, `rId<&"'`);

    expect(xml).toContain(
      '<p:cNvPr id="7" name="Image &lt;&amp;&quot;&apos; _x005F_x0041_" descr="Description &lt;&amp;&quot;&apos;" title="Title &lt;&amp;&quot;&apos;" hidden="1"/>',
    );
    expect(xml).toContain('r:embed="rId&lt;&amp;&quot;&apos;"');
  });

  it('serializes rotation and explicit flip values through shared transforms', () => {
    const xml = serializePicture(
      imageElement(),
      {
        ...TRANSFORM,
        flipHorizontal: false,
        flipVertical: true,
        rotation: -30,
      },
      2,
      'rId2',
    );

    expect(xml).toContain('rot="-1800000" flipH="0" flipV="1"');
  });

  it('serializes exact signed crop percentages before image stretching', () => {
    const element = imageElement();
    element.crop = { bottom: -20, left: 30, right: 0, top: 10.125 };

    expect(serializePicture(element, TRANSFORM, 2, 'rId2')).toContain(
      '<a:blip r:embed="rId2"/><a:srcRect l="30000" t="10125" r="0" b="-20000"/><a:stretch>',
    );
  });

  it('distinguishes explicit visible state from authored absence', () => {
    const visible = imageElement();
    visible.authored.hidden = false;
    expect(serializePicture(visible, TRANSFORM, 2, 'rId2')).toContain(
      ' hidden="0"/>',
    );

    const inherited = imageElement();
    inherited.resolved.hidden = true;
    expect(serializePicture(inherited, TRANSFORM, 2, 'rId2')).not.toContain(
      ' hidden=',
    );
  });
});
