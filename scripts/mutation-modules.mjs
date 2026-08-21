export const mutationModules = [
  {
    estimatedSeconds: 158,
    name: 'orchestration',
    source: 'src/formats/pptx/roundtrip/orchestration.ts',
    tests: [
      'test/pptx/roundtrip-orchestration.test.ts',
      'test/pptx/roundtrip-patch-text.test.ts',
      'test/pptx/roundtrip-transform.test.ts',
      'test/black-box/roundtrip-shape-transform.test.ts',
      'test/black-box/roundtrip-image-transform.test.ts',
      'test/black-box/roundtrip-image-crop.test.ts',
      'test/black-box/roundtrip-table-transform.test.ts',
      'test/black-box/roundtrip-table-text.test.ts',
      'test/black-box/roundtrip-group-transform.test.ts',
      'test/black-box/roundtrip-chart-transform.test.ts',
    ],
  },
  {
    estimatedSeconds: 150,
    name: 'package-preservation',
    source: 'src/formats/pptx/roundtrip/package-preservation.ts',
    tests: ['test/pptx/roundtrip-patch-package.test.ts'],
  },
  {
    estimatedSeconds: 30,
    name: 'patch-error',
    source: 'src/formats/pptx/roundtrip/patch-error.ts',
    tests: ['test/pptx/roundtrip-patch-error.test.ts'],
  },
  {
    estimatedSeconds: 150,
    name: 'relationships',
    source: 'src/formats/pptx/roundtrip/relationships.ts',
    tests: ['test/pptx/roundtrip-relationships.test.ts'],
  },
  {
    estimatedSeconds: 236,
    name: 'shape-range',
    source: 'src/formats/pptx/roundtrip/shape-range.ts',
    tests: ['test/pptx/roundtrip-shape-range.test.ts'],
  },
  {
    estimatedSeconds: 150,
    name: 'text-xml',
    source: 'src/formats/pptx/roundtrip/text-xml.ts',
    tests: ['test/pptx/roundtrip-text-xml.test.ts'],
  },
  {
    estimatedSeconds: 636,
    name: 'transform-xml',
    source: 'src/formats/pptx/roundtrip/transform-xml.ts',
    tests: [
      'test/pptx/roundtrip-transform-xml.test.ts',
      'test/pptx/roundtrip-group-transform-xml.test.ts',
      'test/black-box/roundtrip-chart-transform.test.ts',
    ],
  },
  {
    estimatedSeconds: 77,
    name: 'roundtrip-consistency',
    source: 'src/formats/pptx/roundtrip/consistency.ts',
    tests: [
      'test/pptx/roundtrip-consistency.test.ts',
      'test/black-box/roundtrip-shape-transform.test.ts',
      'test/black-box/roundtrip-chart-transform.test.ts',
    ],
  },
  {
    estimatedSeconds: 465,
    name: 'roundtrip-edit',
    source: 'src/formats/pptx/roundtrip/edit.ts',
    tests: [
      'test/pptx/roundtrip-edit.test.ts',
      'test/pptx/roundtrip-transform.test.ts',
      'test/black-box/roundtrip-shape-transform.test.ts',
      'test/black-box/roundtrip-image-transform.test.ts',
      'test/black-box/roundtrip-image-crop.test.ts',
      'test/black-box/roundtrip-table-transform.test.ts',
      'test/black-box/roundtrip-table-text.test.ts',
      'test/black-box/roundtrip-group-transform.test.ts',
      'test/black-box/roundtrip-chart-transform.test.ts',
    ],
  },
  {
    estimatedSeconds: 329,
    name: 'roundtrip-preview',
    source: 'src/formats/pptx/roundtrip/preview.ts',
    tests: [
      'test/pptx/roundtrip-preview.test.ts',
      'test/black-box/picture-crop.test.ts',
      'test/black-box/roundtrip-image-crop.test.ts',
      'test/black-box/roundtrip-shape-transform.test.ts',
      'test/black-box/roundtrip-chart-transform.test.ts',
    ],
  },
  {
    estimatedSeconds: 240,
    name: 'group-preview',
    source: 'src/formats/pptx/roundtrip/group-preview.ts',
    tests: ['test/pptx/roundtrip-group-preview.test.ts'],
  },
  {
    estimatedSeconds: 180,
    name: 'image-crop-xml',
    source: 'src/formats/pptx/roundtrip/image-crop-xml.ts',
    tests: [
      'test/pptx/roundtrip-image-crop-xml.test.ts',
      'test/black-box/roundtrip-image-crop.test.ts',
    ],
  },
  {
    estimatedSeconds: 420,
    name: 'table-preview',
    source: 'src/formats/pptx/roundtrip/table-preview.ts',
    tests: ['test/pptx/roundtrip-table-preview.test.ts'],
  },
  {
    estimatedSeconds: 420,
    name: 'table-preview-cell',
    source: 'src/formats/pptx/roundtrip/table-preview-cell.ts',
    tests: ['test/pptx/roundtrip-table-preview.test.ts'],
  },
  {
    estimatedSeconds: 300,
    name: 'scene-table-validation',
    source: 'src/formats/pptx/scene-table-validation.ts',
    tests: ['test/pptx/scene-table-validation.test.ts'],
  },
  {
    estimatedSeconds: 300,
    name: 'scene-table-cell-validation',
    source: 'src/formats/pptx/scene-table-cell-validation.ts',
    tests: ['test/pptx/scene-table-validation.test.ts'],
  },
  {
    estimatedSeconds: 300,
    name: 'scene-table-merge-validation',
    source: 'src/formats/pptx/scene-table-merge-validation.ts',
    tests: ['test/pptx/scene-table-validation.test.ts'],
  },
  {
    estimatedSeconds: 559,
    name: 'roundtrip-validation',
    source: 'src/formats/pptx/roundtrip/validate.ts',
    tests: [
      'test/pptx/roundtrip-validation.test.ts',
      'test/black-box/roundtrip-shape-transform.test.ts',
      'test/black-box/roundtrip-image-transform.test.ts',
      'test/black-box/roundtrip-image-crop.test.ts',
      'test/black-box/roundtrip-table-transform.test.ts',
      'test/black-box/roundtrip-table-text.test.ts',
      'test/black-box/roundtrip-group-transform.test.ts',
      'test/black-box/roundtrip-chart-transform.test.ts',
    ],
  },
  {
    estimatedSeconds: 105,
    name: 'writer-chart',
    source: 'src/formats/pptx/writer/chart.ts',
    tests: ['test/pptx/writer-chart.test.ts'],
  },
  {
    estimatedSeconds: 105,
    name: 'writer-shape',
    source: 'src/formats/pptx/writer/shape.ts',
    tests: [
      'test/pptx/writer-shape.test.ts',
      'test/pptx/writer-text-shape.test.ts',
    ],
  },
  {
    estimatedSeconds: 180,
    name: 'writer-group',
    source: 'src/formats/pptx/writer/group.ts',
    tests: ['test/pptx/writer-group.test.ts'],
  },
  {
    estimatedSeconds: 180,
    name: 'writer-group-verify',
    source: 'src/formats/pptx/writer/group-verify.ts',
    tests: ['test/pptx/writer-group-verify.test.ts'],
  },
  {
    estimatedSeconds: 61,
    name: 'writer-image',
    source: 'src/formats/pptx/writer/image.ts',
    tests: ['test/pptx/writer-image.test.ts'],
  },
  {
    estimatedSeconds: 232,
    name: 'writer-table',
    source: 'src/formats/pptx/writer/table.ts',
    tests: ['test/pptx/writer-table.test.ts'],
  },
  {
    estimatedSeconds: 300,
    name: 'writer-table-verify',
    source: 'src/formats/pptx/writer/table-verify.ts',
    tests: ['test/pptx/writer-table-verify.test.ts'],
  },
  {
    estimatedSeconds: 690,
    name: 'font-style',
    source: 'src/formats/pptx/internal/font-style.ts',
    tests: ['test/pptx/font-style.test.ts'],
  },
];

export function resolveMutationModule(name) {
  const result = mutationModules.find((candidate) => candidate.name === name);
  if (result === undefined) {
    throw new Error(
      `Unknown mutation module ${JSON.stringify(name)}; expected one of ${mutationModules
        .map((candidate) => candidate.name)
        .join(', ')}`,
    );
  }
  return result;
}
