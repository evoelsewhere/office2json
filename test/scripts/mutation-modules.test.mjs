import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import {
  mutationModules,
  resolveMutationModule,
} from '../../scripts/mutation-modules.mjs';
import { mutatedFiles } from '../../scripts/mutation-scope.mjs';
import {
  fileMutationTimeoutMs,
  focusedMutationTimeoutMs,
} from '../../scripts/mutation-timeouts.mjs';
import strykerConfig from '../../stryker.config.mjs';
import vitestStrykerConfig from '../../vitest.stryker.config.ts';

const projectRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '../..',
);

describe('focused mutation modules', () => {
  it('bounds focused mutants tightly and gives static file mutants headroom', () => {
    expect(strykerConfig.vitest).toEqual({
      configFile: 'vitest.stryker.config.ts',
    });
    expect(vitestStrykerConfig.test?.bail).toBe(1);
    expect(vitestStrykerConfig.test?.exclude).toContain('**/test/browser/**');
    expect(strykerConfig.timeoutMS).toBe(focusedMutationTimeoutMs);
    expect(fileMutationTimeoutMs).toBeGreaterThan(focusedMutationTimeoutMs);
    expect(fileMutationTimeoutMs).toBeLessThan(4 * 60_000);
    expect(
      fs.readFileSync(
        path.join(projectRoot, 'stryker.shard.config.mjs'),
        'utf8',
      ),
    ).toContain('timeoutMS: fileMutationTimeoutMs');
  });

  it('maps every patch responsibility to an independent source and test set', () => {
    expect(mutationModules.map(({ name }) => name)).toEqual([
      'orchestration',
      'package-preservation',
      'patch-error',
      'relationships',
      'shape-range',
      'text-xml',
      'transform-xml',
      'roundtrip-consistency',
      'roundtrip-edit',
      'roundtrip-preview',
      'group-preview',
      'image-crop-xml',
      'table-preview',
      'table-preview-cell',
      'scene-table-validation',
      'scene-table-cell-validation',
      'scene-table-merge-validation',
      'roundtrip-validation',
      'writer-chart',
      'writer-shape',
      'writer-group',
      'writer-group-verify',
      'writer-image',
      'writer-table',
      'writer-table-verify',
      'font-style',
    ]);
    expect(new Set(mutationModules.map(({ source }) => source)).size).toBe(
      mutationModules.length,
    );
    for (const module of mutationModules) {
      expect(module.estimatedSeconds).toBeGreaterThan(0);
      expect(mutatedFiles).toContain(module.source);
      expect(fs.existsSync(path.join(projectRoot, module.source))).toBe(true);
      expect(module.tests.length).toBeGreaterThan(0);
      for (const test of module.tests) {
        expect(fs.existsSync(path.join(projectRoot, test))).toBe(true);
      }
    }
  });

  it('resolves known names and rejects unknown names', () => {
    expect(resolveMutationModule('relationships')).toMatchObject({
      source: 'src/formats/pptx/roundtrip/relationships.ts',
    });
    expect(() => resolveMutationModule('missing')).toThrow(
      'Unknown mutation module "missing"',
    );
  });
});
