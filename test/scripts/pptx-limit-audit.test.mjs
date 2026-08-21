import { readFile } from 'node:fs/promises';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

describe('PowerPoint remaining-limit audit', () => {
  it('tracks native gaps separately from deliberate security boundaries', async () => {
    const [audit, readme, guide] = await Promise.all([
      readFile(path.resolve('docs', 'pptx-limit-audit.md'), 'utf8'),
      readFile(path.resolve('README.md'), 'utf8'),
      readFile(path.resolve('docs', 'pptx-usage-guide.md'), 'utf8'),
    ]);

    for (const capability of [
      'Shape-owned rich text',
      'Image replacement',
      'Table structure',
      'Chart data',
      'Chart formatting',
      'Advanced charts',
      'Speaker notes',
      'Transitions',
      'Accessibility metadata',
    ]) {
      expect(audit, capability).toContain(capability);
    }
    for (const boundary of [
      'encrypted or password-protected packages',
      'signatures',
      'VBA',
      'external relationships are never fetched',
      'Markup Compatibility',
      'Claiming arbitrary PPTX creation or editing',
      'pixel-identical rendering',
    ]) {
      expect(audit, boundary).toContain(boundary);
    }
    expect(audit).toContain('slide-1-element-2-row-3-cell-1-run-1');
    expect(audit).toContain('slide-1-element-2-element-1-run-1');
    expect(audit).toMatch(
      /\| Tables\s+\|[^\n]+single plain cell-run replacement\s+\| Complete\s+\|/,
    );
    expect(audit).toMatch(
      /\| Groups\s+\|[^\n]+nested plain text\s+\| Complete\s+\|/,
    );
    expect(audit).toMatch(
      /\| Images\s+\|[^\n]+crop add\/replace\/remove[^\n]+Certification pending\s+\|/,
    );
    expect(audit).toContain('24,318 mutants');
    expect(audit).toContain('zero missed');
    expect(audit).toContain('strict mutation threshold at 100%');
    expect(readme).toContain(
      '[PowerPoint remaining-limit audit](docs/pptx-limit-audit.md)',
    );
    expect(guide).toContain('Arbitrary PPTX creation/editing');
    expect((audit.match(/^```/gm)?.length ?? 0) % 2).toBe(0);
  });
});
