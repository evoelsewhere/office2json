import { describe, expect, it } from 'vitest';

import { resolvePptxResourceLimits } from '../../src/formats/pptx/internal/resource-limits';
import type { PptxRoundTripSnapshot } from '../../src/formats/pptx/roundtrip/types';
import { validatePptxRoundTripSnapshot } from '../../src/formats/pptx/roundtrip/validate';

const HASH = 'a'.repeat(64);

function snapshot(): PptxRoundTripSnapshot {
  return {
    consistency: {
      canonicalizationVersion: 'canonical-json-v1',
      capabilityProfileVersion: 'pptx-roundtrip-text-v1',
      contractVersion: '1',
      hashAlgorithm: 'sha256',
      keyAlgorithmVersion: 'pptx-scene-key-v1',
      operationsSha256: HASH,
      semanticPreviewSha256: HASH,
      sourceManifestSha256: HASH,
    },
    document: {
      layouts: [],
      masters: [],
      media: [],
      schemaVersion: 2,
      size: { height: 540, width: 960 },
      slides: [],
      themes: [],
    },
    format: 'pptx',
    operations: [],
    schemaVersion: 1,
    source: {
      byteLength: 1,
      conformance: 'unknown',
      data: new Uint8Array([1]),
      kind: 'bytes',
      sha256: HASH,
    },
    supportProfile: {
      effectiveLevel: 'R0',
      id: 'pptx-roundtrip-r0',
      producerMatrix: [],
      version: '1',
    },
  };
}

function rootRecord(value: PptxRoundTripSnapshot): Record<string, unknown> {
  return value as unknown as Record<string, unknown>;
}

function unknownRecord(value: unknown): Record<string, unknown> {
  return value as Record<string, unknown>;
}

function nestedRecord(
  value: PptxRoundTripSnapshot,
  key: 'consistency' | 'source' | 'supportProfile',
): Record<string, unknown> {
  return value[key] as unknown as Record<string, unknown>;
}

function expectInvalid(value: unknown, code: string, message: string): void {
  let received: unknown;
  try {
    validatePptxRoundTripSnapshot(value, resolvePptxResourceLimits());
  } catch (error) {
    received = error;
  }
  expect(received).toMatchObject({ code, message });
}

describe('PowerPoint round-trip snapshot contract validation', () => {
  it('accepts the exact R0 runtime shape without replacing it', () => {
    const value = snapshot();
    expect(
      validatePptxRoundTripSnapshot(value, resolvePptxResourceLimits()),
    ).toBe(value);
  });

  it.each([
    [
      'primitive root',
      () => null,
      'invalid-snapshot',
      'PowerPoint round-trip snapshot has an invalid root shape',
    ],
    [
      'extra root field',
      () => {
        const value = snapshot();
        rootRecord(value).extra = true;
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip snapshot has an invalid root shape',
    ],
    [
      'missing root field',
      () => {
        const value = snapshot();
        delete rootRecord(value).document;
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip snapshot has an invalid root shape',
    ],
    [
      'substituted root field',
      () => {
        const value = snapshot();
        delete rootRecord(value).document;
        rootRecord(value).extra = true;
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip snapshot has an invalid root shape',
    ],
    [
      'schema version',
      () => {
        const value = snapshot();
        rootRecord(value).schemaVersion = 2;
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip snapshot schema version is unsupported',
    ],
    [
      'format',
      () => {
        const value = snapshot();
        rootRecord(value).format = 'docx';
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip snapshot format must be pptx',
    ],
    [
      'operations type',
      () => {
        const value = snapshot();
        rootRecord(value).operations = {};
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip snapshot operations must be an array',
    ],
    [
      'nonempty operations',
      () => {
        const value = snapshot();
        rootRecord(value).operations = [{ type: 'replaceText' }];
        return value;
      },
      'invalid-snapshot',
      'PowerPoint round-trip operation 1 kind is unsupported',
    ],
  ])('rejects %s', (_name, create, code, message) => {
    expectInvalid(create(), code, message);
  });

  it('accepts a text edit with an exact preview precondition and R2 profile', () => {
    const value = snapshot();
    value.document.slides = [
      {
        elements: [
          {
            authored: {},
            key: 'text-1',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [{ key: 'run-1', text: 'Before', type: 'run' }],
                  key: 'paragraph-1',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'slide-1',
      },
    ];
    value.operations = [
      {
        expectedText: 'Before',
        id: 'replace-text-1',
        kind: 'replace-text',
        targetKey: 'run-1',
        value: 'After',
      },
    ];
    value.supportProfile = {
      effectiveLevel: 'R2',
      id: 'pptx-roundtrip-text-v1',
      producerMatrix: [],
      version: '1',
    };

    expect(
      validatePptxRoundTripSnapshot(value, resolvePptxResourceLimits()),
    ).toBe(value);
  });

  function transformSnapshot(): PptxRoundTripSnapshot {
    const value = snapshot();
    value.document.slides = [
      {
        elements: [
          {
            authored: {},
            key: 'text-1',
            resolved: {
              hidden: false,
              transform: {
                flipHorizontal: false,
                flipVertical: false,
                height: 40,
                rotation: 0,
                width: 160,
                x: 10,
                y: 20,
              },
            },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [{ key: 'run-1', text: 'Before', type: 'run' }],
                  key: 'paragraph-1',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'slide-1',
      },
    ];
    value.operations = [
      {
        expectedTransform: {
          flipHorizontal: false,
          flipVertical: false,
          height: 40,
          rotation: 0,
          width: 160,
          x: 10,
          y: 20,
        },
        id: 'set-transform-1',
        kind: 'set-transform',
        targetKey: 'text-1',
        value: {
          flipHorizontal: true,
          flipVertical: false,
          height: 50,
          rotation: 15,
          width: 170,
          x: 30,
          y: 40,
        },
      },
    ];
    value.supportProfile = {
      effectiveLevel: 'R2',
      id: 'pptx-roundtrip-text-v1',
      producerMatrix: [],
      version: '1',
    };
    return value;
  }

  function replaceSnapshot(replacement = 'After'): PptxRoundTripSnapshot {
    const value = transformSnapshot();
    value.operations = [
      {
        expectedText: 'Before',
        id: 'replace-text-1',
        kind: 'replace-text',
        targetKey: 'run-1',
        value: replacement,
      },
    ];
    return value;
  }

  function imageCropSnapshot(): PptxRoundTripSnapshot {
    const value = snapshot();
    value.consistency.capabilityProfileVersion = 'pptx-roundtrip-native-v1';
    value.document.slides = [
      {
        elements: [
          {
            authored: {},
            crop: { bottom: -20, left: 30, right: 0, top: 10 },
            key: 'image-1',
            resolved: {
              hidden: false,
              transform: {
                flipHorizontal: false,
                flipVertical: false,
                height: 40,
                rotation: 0,
                width: 160,
                x: 10,
                y: 20,
              },
            },
            type: 'image',
          },
        ],
        key: 'slide-1',
      },
    ];
    value.operations = [
      {
        expectedCrop: { bottom: -20, left: 30, right: 0, top: 10 },
        id: 'set-image-crop-1',
        kind: 'set-image-crop',
        targetKey: 'image-1',
        value: { bottom: 5, left: 10, right: 15, top: 20 },
      },
    ];
    value.supportProfile = {
      effectiveLevel: 'R2',
      id: 'pptx-roundtrip-native-v1',
      producerMatrix: [],
      version: '1',
    };
    return value;
  }

  it('accepts a transform edit with an exact preview precondition', () => {
    const value = transformSnapshot();

    expect(
      validatePptxRoundTripSnapshot(value, resolvePptxResourceLimits()),
    ).toBe(value);
  });

  it('accepts image crop replacement, removal, and same-target transform', () => {
    const replacement = imageCropSnapshot();
    expect(
      validatePptxRoundTripSnapshot(replacement, resolvePptxResourceLimits()),
    ).toBe(replacement);

    const removal = imageCropSnapshot();
    const removalOperation = removal.operations[0];
    if (removalOperation?.kind !== 'set-image-crop') {
      throw new Error('Expected crop operation');
    }
    removalOperation.value = null;
    expect(
      validatePptxRoundTripSnapshot(removal, resolvePptxResourceLimits()),
    ).toBe(removal);

    const composed = imageCropSnapshot();
    composed.operations.push({
      expectedTransform: {
        flipHorizontal: false,
        flipVertical: false,
        height: 40,
        rotation: 0,
        width: 160,
        x: 10,
        y: 20,
      },
      id: 'set-transform-2',
      kind: 'set-transform',
      targetKey: 'image-1',
      value: {
        flipHorizontal: true,
        flipVertical: false,
        height: 50,
        rotation: 10,
        width: 170,
        x: 30,
        y: 40,
      },
    });
    expect(
      validatePptxRoundTripSnapshot(composed, resolvePptxResourceLimits()),
    ).toBe(composed);

    for (const boundary of [
      { bottom: 1.015, left: 0, right: 0, top: 0 },
      { bottom: 0, left: -100, right: 0, top: 0 },
      { bottom: 0, left: 100, right: -1, top: 0 },
    ]) {
      const value = imageCropSnapshot();
      const operation = value.operations[0];
      if (operation?.kind !== 'set-image-crop') {
        throw new Error('Expected crop operation');
      }
      operation.value = boundary;
      expect(
        validatePptxRoundTripSnapshot(value, resolvePptxResourceLimits()),
      ).toBe(value);
    }
  });

  it.each([
    [
      'missing target',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation !== undefined) operation.targetKey = 'missing';
      },
      'PowerPoint round-trip image crop target does not exist',
    ],
    [
      'stale precondition',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-image-crop' && operation.expectedCrop) {
          operation.expectedCrop.left = 31;
        }
      },
      'PowerPoint round-trip image crop precondition does not match the preview',
    ],
    [
      'no-op',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-image-crop') {
          operation.value = structuredClone(operation.expectedCrop);
        }
      },
      'PowerPoint round-trip image crop must change the value',
    ],
    [
      'invalid expected crop',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-image-crop' && operation.expectedCrop) {
          delete unknownRecord(operation.expectedCrop).top;
        }
      },
      'PowerPoint round-trip operation 1 expectedCrop is invalid',
    ],
    [
      'missing edge',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-image-crop' && operation.value) {
          delete unknownRecord(operation.value).top;
        }
      },
      'PowerPoint round-trip operation 1 value is not a valid image crop',
    ],
    [
      'invalid percentage',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-image-crop' && operation.value) {
          operation.value.left = 101;
          operation.value.right = -2;
        }
      },
      'PowerPoint round-trip operation 1 value is not a valid image crop',
    ],
    [
      'low percentage',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-image-crop' && operation.value) {
          operation.value.bottom = -100.001;
        }
      },
      'PowerPoint round-trip operation 1 value is not a valid image crop',
    ],
    [
      'excess precision',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-image-crop' && operation.value) {
          operation.value.left = 0.0001;
        }
      },
      'PowerPoint round-trip operation 1 value is not a valid image crop',
    ],
    [
      'collapsed region',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-image-crop' && operation.value) {
          operation.value.left = 60;
          operation.value.right = 40;
        }
      },
      'PowerPoint round-trip operation 1 value is not a valid image crop',
    ],
    [
      'collapsed vertical region',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-image-crop' && operation.value) {
          operation.value.top = 60;
          operation.value.bottom = 40;
        }
      },
      'PowerPoint round-trip operation 1 value is not a valid image crop',
    ],
  ])('rejects image crop operation with %s', (_name, mutate, message) => {
    const value = imageCropSnapshot();
    mutate(value);
    expectInvalid(value, 'invalid-snapshot', message);
  });

  it.each(['bottom', 'left', 'right', 'top'] as const)(
    'rejects a non-numeric image crop %s edge',
    (edge) => {
      const value = imageCropSnapshot();
      const operation = value.operations[0];
      if (operation?.kind !== 'set-image-crop' || operation.value === null) {
        throw new Error('Expected crop operation');
      }
      unknownRecord(operation.value)[edge] = '1';
      expectInvalid(
        value,
        'invalid-snapshot',
        'PowerPoint round-trip operation 1 value is not a valid image crop',
      );
    },
  );

  it('reports the exact index for a malformed second image crop', () => {
    const value = imageCropSnapshot();
    const first = value.operations[0];
    if (first?.kind !== 'set-image-crop') {
      throw new Error('Expected crop operation');
    }
    value.document.slides[0]?.elements.push({
      authored: {},
      key: 'image-2',
      resolved: { hidden: false },
      type: 'image',
    });
    value.operations.push({
      ...structuredClone(first),
      expectedCrop: null,
      id: 'set-image-crop-2',
      targetKey: 'image-2',
      value: { bottom: 0, left: 0, right: 0 } as never,
    });
    expectInvalid(
      value,
      'invalid-snapshot',
      'PowerPoint round-trip operation 2 value is not a valid image crop',
    );
  });

  it('reports the exact index of a malformed second operation', () => {
    const value = transformSnapshot();
    rootRecord(value).operations = [
      ...value.operations,
      { kind: 'set-transform' },
    ];
    expectInvalid(
      value,
      'invalid-snapshot',
      'PowerPoint round-trip operation 2 has an invalid shape',
    );
  });

  it.each([
    [
      'missing target',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation !== undefined) operation.targetKey = 'missing';
      },
      'PowerPoint round-trip transform target does not exist',
    ],
    [
      'stale precondition',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-transform')
          operation.expectedTransform.x = 11;
      },
      'PowerPoint round-trip transform precondition does not match the preview',
    ],
    [
      'no-op',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-transform') {
          operation.value = structuredClone(operation.expectedTransform);
        }
      },
      'PowerPoint round-trip transform must change the value',
    ],
    [
      'missing expected field',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-transform') {
          delete unknownRecord(operation.expectedTransform).x;
        }
      },
      'PowerPoint round-trip operation 1 expectedTransform is invalid',
    ],
    [
      'zero width',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'set-transform') operation.value.width = 0;
      },
      'PowerPoint round-trip operation 1 value is not a valid transform',
    ],
  ])('rejects transform operation with %s', (_name, mutate, message) => {
    const value = transformSnapshot();
    mutate(value);
    expectInvalid(value, 'invalid-snapshot', message);
  });

  it.each([null, [], 7])(
    'rejects primitive operation value %j',
    (operation) => {
      const value = transformSnapshot();
      rootRecord(value).operations = [operation];
      expectInvalid(
        value,
        'invalid-snapshot',
        'PowerPoint round-trip operation 1 has an invalid shape',
      );
    },
  );

  it.each([
    ['id', 7, 'PowerPoint round-trip operation 1 id must be a string'],
    [
      'targetKey',
      7,
      'PowerPoint round-trip operation 1 targetKey must be a string',
    ],
    ['id', '', 'PowerPoint round-trip operation ids must be unique'],
    [
      'targetKey',
      '',
      'PowerPoint round-trip text edit targets must be non-empty and unique',
    ],
  ])('rejects transform %s=%j', (key, replacement, message) => {
    const value = transformSnapshot();
    const operation = value.operations[0];
    if (operation === undefined) throw new Error('Expected operation');
    unknownRecord(operation)[key] = replacement;
    expectInvalid(value, 'invalid-snapshot', message);
  });

  it('rejects duplicate operation ids before resolving the second target', () => {
    const value = transformSnapshot();
    const operation = value.operations[0];
    if (operation?.kind !== 'set-transform')
      throw new Error('Expected transform');
    value.operations.push({
      ...structuredClone(operation),
      targetKey: 'missing',
      value: { ...operation.value, x: 31 },
    });
    expectInvalid(
      value,
      'invalid-snapshot',
      'PowerPoint round-trip operation ids must be unique',
    );
  });

  it('rejects duplicate operation targets with unique ids', () => {
    const value = transformSnapshot();
    const operation = value.operations[0];
    if (operation?.kind !== 'set-transform')
      throw new Error('Expected transform');
    value.operations.push({
      ...structuredClone(operation),
      id: 'set-transform-2',
      value: { ...operation.value, x: 31 },
    });
    expectInvalid(
      value,
      'invalid-snapshot',
      'PowerPoint round-trip text edit targets must be non-empty and unique',
    );
  });

  it.each(['x', 'y', 'width', 'height', 'rotation'] as const)(
    'rejects non-numeric transform field %s',
    (key) => {
      for (const replacement of ['1', Number.NaN, Number.POSITIVE_INFINITY]) {
        const value = transformSnapshot();
        const operation = value.operations[0];
        if (operation?.kind !== 'set-transform') {
          throw new Error('Expected transform');
        }
        unknownRecord(operation.value)[key] = replacement;
        expectInvalid(
          value,
          'invalid-snapshot',
          typeof replacement === 'number'
            ? 'PowerPoint round-trip snapshot contains a non-JSON value'
            : 'PowerPoint round-trip operation 1 value is not a valid transform',
        );
      }
    },
  );

  it.each(['width', 'height'] as const)(
    'rejects non-positive transform %s',
    (key) => {
      for (const replacement of [0, -1]) {
        const value = transformSnapshot();
        const operation = value.operations[0];
        if (operation?.kind !== 'set-transform') {
          throw new Error('Expected transform');
        }
        operation.value[key] = replacement;
        expectInvalid(
          value,
          'invalid-snapshot',
          'PowerPoint round-trip operation 1 value is not a valid transform',
        );
      }
    },
  );

  it.each(['flipHorizontal', 'flipVertical'] as const)(
    'rejects non-boolean transform field %s',
    (key) => {
      const value = transformSnapshot();
      const operation = value.operations[0];
      if (operation?.kind !== 'set-transform') {
        throw new Error('Expected transform');
      }
      unknownRecord(operation.value)[key] = 1;
      expectInvalid(
        value,
        'invalid-snapshot',
        'PowerPoint round-trip operation 1 value is not a valid transform',
      );
    },
  );

  it('rejects duplicate transform keys at semantic scene validation', () => {
    const value = transformSnapshot();
    const slide = value.document.slides[0];
    const element = slide?.elements[0];
    if (slide === undefined || element?.type !== 'text') {
      throw new Error('Expected text element');
    }
    const duplicate = structuredClone(element);
    const paragraph = duplicate.text.paragraphs[0];
    const child = paragraph?.children[0];
    if (child === undefined) throw new Error('Expected text child');
    child.key = 'unique-run';
    slide.elements.push(duplicate);
    expectInvalid(
      value,
      'invalid-snapshot',
      'PowerPoint round-trip snapshot semantic preview is invalid',
    );
  });

  it.each([
    [
      'missing target',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation !== undefined) operation.targetKey = 'missing';
      },
      'PowerPoint round-trip text edit target does not exist',
    ],
    [
      'stale precondition',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'replace-text')
          operation.expectedText = 'Stale';
      },
      'PowerPoint round-trip text edit precondition does not match the preview',
    ],
    [
      'no-op',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'replace-text') operation.value = 'Before';
      },
      'PowerPoint round-trip text edit must change the value',
    ],
    [
      'unsafe XML',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation?.kind === 'replace-text') operation.value = 'bad\u0000';
      },
      'PowerPoint round-trip text edit value is not safe XML text',
    ],
    [
      'non-string expected text',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation !== undefined) unknownRecord(operation).expectedText = 7;
      },
      'PowerPoint round-trip operation 1 expectedText must be a string',
    ],
    [
      'non-string value',
      (value: PptxRoundTripSnapshot) => {
        const operation = value.operations[0];
        if (operation !== undefined) unknownRecord(operation).value = 7;
      },
      'PowerPoint round-trip operation 1 value must be a string',
    ],
  ])('rejects replace operation with %s', (_name, mutate, message) => {
    const value = replaceSnapshot();
    mutate(value);
    expectInvalid(value, 'invalid-snapshot', message);
  });

  it('enforces text byte limits at exact UTF-16 and UTF-8 boundaries', () => {
    const limits = resolvePptxResourceLimits();
    limits.maxXmlBytes = 4;
    expect(() =>
      validatePptxRoundTripSnapshot(replaceSnapshot('1234'), limits),
    ).not.toThrow();
    for (const replacement of ['12345', '😀😀']) {
      expect(() =>
        validatePptxRoundTripSnapshot(replaceSnapshot(replacement), limits),
      ).toThrow(
        'PowerPoint round-trip text edit value exceeds the XML part byte limit',
      );
    }
  });

  it.each([
    ['break', 'break-1'],
    ['unsupported element', 'unsupported'],
  ])('does not expose %s as an editable run', (_name, targetKey) => {
    const value = replaceSnapshot();
    const slide = value.document.slides[0];
    const element = slide?.elements[0];
    if (slide === undefined || element?.type !== 'text') {
      throw new Error('Expected text element');
    }
    const paragraph = element.text.paragraphs[0];
    if (paragraph === undefined) throw new Error('Expected paragraph');
    paragraph.children.push({ key: 'break-1', type: 'break' });
    slide.elements.push({
      authored: {},
      feature: 'shape',
      key: 'unsupported',
      resolved: { hidden: false },
      type: 'unsupported',
    });
    const operation = value.operations[0];
    if (operation !== undefined) operation.targetKey = targetKey;
    expectInvalid(
      value,
      'invalid-snapshot',
      'PowerPoint round-trip text edit target does not exist',
    );
  });

  it.each([
    ['unsupported element', 'unsupported'],
    ['text without geometry', 'no-transform'],
  ])('does not expose %s as an editable transform', (_name, targetKey) => {
    const value = transformSnapshot();
    const slide = value.document.slides[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements.push(
      {
        authored: {},
        feature: 'shape',
        key: 'unsupported',
        resolved: {
          hidden: false,
          transform: { height: 10, width: 10, x: 1, y: 1 },
        },
        type: 'unsupported',
      },
      {
        authored: {},
        key: 'no-transform',
        resolved: { hidden: false },
        text: {
          body: {},
          paragraphs: [
            {
              children: [{ key: 'other-run', text: 'Other', type: 'run' }],
              key: 'other-paragraph',
            },
          ],
        },
        type: 'text',
      },
    );
    const operation = value.operations[0];
    if (operation?.kind === 'set-transform') {
      operation.targetKey = targetKey;
      operation.expectedTransform = {
        flipHorizontal: false,
        flipVertical: false,
        height: 10,
        rotation: 0,
        width: 10,
        x: 1,
        y: 1,
      };
    }
    expectInvalid(
      value,
      'invalid-snapshot',
      'PowerPoint round-trip transform target does not exist',
    );
  });

  it.each([
    [
      'shape',
      (value: PptxRoundTripSnapshot) => {
        rootRecord(value).source = {};
      },
      'PowerPoint round-trip snapshot source has an invalid shape',
    ],
    [
      'extra field',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').extra = true;
      },
      'PowerPoint round-trip snapshot source has an invalid shape',
    ],
    [
      'kind',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').kind = 'base64';
      },
      'PowerPoint round-trip snapshot source kind must be bytes',
    ],
    [
      'data',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').data = {};
      },
      'PowerPoint round-trip snapshot source data must be Uint8Array or Blob',
    ],
    [
      'zero length',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').byteLength = 0;
        nestedRecord(value, 'source').data = new Uint8Array();
      },
      'PowerPoint round-trip snapshot source byteLength must be a positive safe integer',
    ],
    [
      'fractional length',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').byteLength = 1.5;
      },
      'PowerPoint round-trip snapshot source byteLength must be a positive safe integer',
    ],
    [
      'negative length',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').byteLength = -1;
      },
      'PowerPoint round-trip snapshot source byteLength must be a positive safe integer',
    ],
    [
      'unsafe length',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').byteLength = Number.MAX_SAFE_INTEGER + 1;
      },
      'PowerPoint round-trip snapshot source byteLength must be a positive safe integer',
    ],
    [
      'declared length',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').byteLength = 2;
      },
      'PowerPoint round-trip snapshot source byteLength does not match its data',
    ],
    [
      'conformance',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').conformance = 'future';
      },
      'PowerPoint round-trip snapshot source conformance is invalid',
    ],
    [
      'SHA-256',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').sha256 = 'A'.repeat(64);
      },
      'PowerPoint round-trip snapshot source SHA-256 is invalid',
    ],
    [
      'SHA-256 prefix',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').sha256 = `x${HASH}`;
      },
      'PowerPoint round-trip snapshot source SHA-256 is invalid',
    ],
    [
      'SHA-256 suffix',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'source').sha256 = `${HASH}x`;
      },
      'PowerPoint round-trip snapshot source SHA-256 is invalid',
    ],
  ])('rejects source %s', (_name, mutate, message) => {
    const value = snapshot();
    mutate(value);
    expectInvalid(value, 'invalid-snapshot', message);
  });

  it('accepts every declared source conformance and Blob transport', () => {
    for (const conformance of ['strict', 'transitional', 'unknown'] as const) {
      const value = snapshot();
      value.source.conformance = conformance;
      value.source.data = new Blob([new Uint8Array([1]).buffer]);
      expect(() =>
        validatePptxRoundTripSnapshot(value, resolvePptxResourceLimits()),
      ).not.toThrow();
    }
  });

  it('applies the runtime input byte limit to snapshot source data', () => {
    expect(() =>
      validatePptxRoundTripSnapshot(
        snapshot(),
        resolvePptxResourceLimits({ maxInputBytes: 1 }),
      ),
    ).not.toThrow();
    const value = snapshot();
    value.source.byteLength = 2;
    value.source.data = new Uint8Array([1, 2]);
    expect(() =>
      validatePptxRoundTripSnapshot(
        value,
        resolvePptxResourceLimits({ maxInputBytes: 1 }),
      ),
    ).toThrow('maxInputBytes exceeded');
  });

  it.each([
    [
      'shape',
      (value: PptxRoundTripSnapshot) => {
        rootRecord(value).supportProfile = {};
      },
      'PowerPoint round-trip snapshot support profile has an invalid shape',
    ],
    [
      'level',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'supportProfile').effectiveLevel = 'R1';
      },
      'PowerPoint round-trip snapshot support level must be R0',
    ],
    [
      'id',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'supportProfile').id = 'other';
      },
      'PowerPoint round-trip snapshot support profile id is unsupported',
    ],
    [
      'version',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'supportProfile').version = '2';
      },
      'PowerPoint round-trip snapshot support profile version is unsupported',
    ],
    [
      'producer matrix type',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'supportProfile').producerMatrix = {};
      },
      'PowerPoint round-trip snapshot producer matrix must be an array',
    ],
    [
      'producer claim',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'supportProfile').producerMatrix = ['PowerPoint'];
      },
      'PowerPoint round-trip snapshot cannot claim unverified producer evidence',
    ],
  ])('rejects support profile %s', (_name, mutate, message) => {
    const value = snapshot();
    mutate(value);
    expectInvalid(value, 'invalid-snapshot', message);
  });

  it('reports the R2 support-level requirement for edited snapshots', () => {
    const value = transformSnapshot();
    value.supportProfile.effectiveLevel = 'R1';
    expectInvalid(
      value,
      'invalid-snapshot',
      'PowerPoint text edit snapshot support level must be R2',
    );
  });

  it.each([
    [
      'shape',
      (value: PptxRoundTripSnapshot) => {
        rootRecord(value).consistency = {};
      },
      'PowerPoint round-trip snapshot consistency has an invalid shape',
    ],
    [
      'canonicalization version',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').canonicalizationVersion = 'future';
      },
      'PowerPoint round-trip snapshot canonicalization version is unsupported',
    ],
    [
      'capability profile version',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').capabilityProfileVersion = 'future';
      },
      'PowerPoint round-trip snapshot capability profile version is unsupported',
    ],
    [
      'contract version',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').contractVersion = '2';
      },
      'PowerPoint round-trip snapshot contract version is unsupported',
    ],
    [
      'hash algorithm',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').hashAlgorithm = 'sha512';
      },
      'PowerPoint round-trip snapshot hash algorithm is unsupported',
    ],
    [
      'key algorithm version',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').keyAlgorithmVersion = 'future';
      },
      'PowerPoint round-trip snapshot key algorithm version is unsupported',
    ],
    [
      'operations digest',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').operationsSha256 = 'short';
      },
      'PowerPoint round-trip snapshot operations SHA-256 is invalid',
    ],
    [
      'preview digest',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').semanticPreviewSha256 = 'short';
      },
      'PowerPoint round-trip snapshot semantic preview SHA-256 is invalid',
    ],
    [
      'source digest',
      (value: PptxRoundTripSnapshot) => {
        nestedRecord(value, 'consistency').sourceManifestSha256 = 'short';
      },
      'PowerPoint round-trip snapshot source manifest SHA-256 is invalid',
    ],
  ])('rejects consistency %s', (_name, mutate, message) => {
    const value = snapshot();
    mutate(value);
    expectInvalid(value, 'invalid-snapshot', message);
  });

  it('rejects an invalid semantic preview with its scene issues', () => {
    const value = snapshot();
    value.document.size.width = 0;

    let received: unknown;
    try {
      validatePptxRoundTripSnapshot(value, resolvePptxResourceLimits());
    } catch (error) {
      received = error;
    }
    expect(received).toMatchObject({
      code: 'invalid-snapshot',
      issues: [
        {
          code: 'invalid-numeric-value',
          message: 'Expected a positive finite number',
          path: '$.size.width',
        },
      ],
      message: 'PowerPoint round-trip snapshot semantic preview is invalid',
    });
  });
});
