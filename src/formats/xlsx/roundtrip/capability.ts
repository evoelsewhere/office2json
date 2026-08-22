import type {
  XlsxCapabilityEntry,
  XlsxCapabilityManifest,
  XlsxEditOperation,
} from './types';

function capabilityDomains(): string[] {
  return [
    'active-content',
    'calculation',
    'cells',
    'charts',
    'comments',
    'conditional-formatting',
    'connections',
    'defined-names',
    'document-properties',
    'drawings-images',
    'external-links',
    'filters-sorts',
    'formulas',
    'hyperlinks',
    'known-extensions',
    'merges',
    'modern-cell-metadata',
    'pivots',
    'print-layout',
    'protection',
    'rows-columns',
    'shared-strings',
    'sheet-metadata',
    'sparklines',
    'styles',
    'tables',
    'unknown-extensions',
    'validation',
    'views',
    'workbook-sheets',
  ];
}

function capabilityOperations(): Array<XlsxEditOperation['kind']> {
  return [
    'add-worksheet',
    'clear-cell',
    'delete-columns',
    'delete-rows',
    'delete-worksheet',
    'insert-columns',
    'insert-rows',
    'rename-worksheet',
    'set-cell',
    'set-cell-style',
    'set-column',
    'set-hyperlink',
    'set-row',
  ];
}

export function createXlsxCapabilityManifest(): XlsxCapabilityManifest {
  const domains: XlsxCapabilityEntry[] = capabilityDomains().map((domain) => ({
    domain,
    level:
      domain === 'cells' ||
      domain === 'formulas' ||
      domain === 'hyperlinks' ||
      domain === 'rows-columns' ||
      domain === 'styles'
        ? 'verified-R2'
        : 'preservation-only',
  }));
  return {
    domains,
    effectiveLevel: 'R2',
    id: 'xlsx-agent-ready',
    operations: capabilityOperations().map((operation) => ({
      ...(operation === 'clear-cell' || operation === 'set-cell'
        ? {
            constraints: [
              'existing-explicit-cell',
              'clean-supported-package-closure',
              'no-unaffected-formulas-or-defined-names',
              'no-grouped-formula-target',
              'no-date-or-rich-text-value',
              'no-external-capable-formula',
            ],
            level: 'verified-R2' as const,
          }
        : operation === 'set-cell-style'
          ? {
              constraints: [
                'existing-explicit-cell',
                'existing-or-append-normalized-style',
                'existing-styles-part-for-append',
                'no-new-checkbox-style',
                'clean-supported-package-closure',
                'no-non-anchor-merged-cell',
              ],
              level: 'verified-R2' as const,
            }
          : operation === 'set-row' || operation === 'set-column'
            ? {
                constraints: [
                  operation === 'set-row'
                    ? 'existing-explicit-row'
                    : 'existing-exact-column-range',
                  'size-and-visibility-only',
                  'clean-supported-package-closure',
                ],
                level: 'verified-R2' as const,
              }
            : operation === 'delete-columns' ||
                operation === 'delete-rows' ||
                operation === 'insert-columns' ||
                operation === 'insert-rows'
              ? {
                  constraints: [
                    'reference-free-simple-worksheet',
                    'explicit-row-and-cell-references',
                    'structural-operations-only-batch',
                    'no-explicit-column-definitions-for-column-shifts',
                    'no-grid-overflow',
                    'clean-supported-package-closure',
                  ],
                  level: 'verified-R2' as const,
                }
              : operation === 'set-hyperlink'
                ? {
                    constraints: [
                      'existing-explicit-cell',
                      'safe-internal-or-external-target',
                      'deterministic-relationship-allocation',
                      'no-overlapping-multi-cell-hyperlink',
                      'http-https-mailto-only',
                      'no-url-credentials',
                      'clean-supported-package-closure',
                    ],
                    level: 'verified-R2' as const,
                  }
                : { level: 'unsupported' as const }),
      operation,
    })),
    producerEvidence: [],
    version: '1',
  };
}
