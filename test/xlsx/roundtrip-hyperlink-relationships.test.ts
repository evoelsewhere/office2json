import { describe, expect, it } from 'vitest';

import { XlsxWriteError } from '../../src/formats/xlsx/roundtrip/errors';
import {
  patchXlsxHyperlinkRelationships,
  planXlsxExternalHyperlinkRelationships,
} from '../../src/formats/xlsx/roundtrip/hyperlink-relationships';
import type { XlsxPackageGraphRelationship } from '../../src/formats/xlsx/roundtrip/internal/package-graph';
import { defaultXlsxWriteLimits } from '../../src/formats/xlsx/roundtrip/write-limits';
import { defaultXlsxResourceLimits } from '../../src/formats/xlsx/internal/resource-limits';
import type { XlsxHyperlink } from '../../src/formats/xlsx/types';
import {
  XLSX_OFFICE_REL_NS,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

const SHEET = 'xl/worksheets/sheet1.xml';
const PART = 'xl/worksheets/_rels/sheet1.xml.rels';
const TYPE = `${XLSX_OFFICE_REL_NS}/hyperlink`;

function bytes(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function hyperlink(
  reference: string,
  target: XlsxHyperlink['target'],
): XlsxHyperlink {
  const parts = reference.split(':');
  const start = parts[0]!;
  const end = parts[1] ?? start;
  const cell = (value: string) => ({
    column: value.charCodeAt(0) - 64,
    row: Number(value.slice(1)),
  });
  return {
    range: { end: cell(end), reference, start: cell(start) },
    selectionRelation: 'full-sheet',
    target,
  };
}

function capture(action: () => unknown): XlsxWriteError {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected relationship patch to fail');
}

describe('XLSX external hyperlink relationships', () => {
  it('reuses source IDs and appends after the maximum numeric ID', () => {
    const worksheet = bytes(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><hyperlinks><hyperlink ref="A1" r:id="rId2"/></hyperlinks></worksheet>`,
    );
    const relationships: XlsxPackageGraphRelationship[] = [
      {
        id: 'rId2',
        mode: 'external',
        owner: SHEET,
        target: 'https://old.invalid/',
        type: TYPE,
      },
      {
        id: 'rId9',
        mode: 'external',
        owner: SHEET,
        target: 'mailto:unused@example.invalid',
        type: TYPE,
      },
      {
        id: 'rId99x',
        mode: 'external',
        owner: SHEET,
        target: 'mailto:suffix@example.invalid',
        type: TYPE,
      },
      {
        id: 'xrId88',
        mode: 'external',
        owner: SHEET,
        target: 'mailto:prefix@example.invalid',
        type: TYPE,
      },
    ];
    const plan = planXlsxExternalHyperlinkRelationships(
      worksheet,
      relationships,
      [
        hyperlink('A1', { kind: 'external', url: 'https://new.invalid/' }),
        hyperlink('B2', {
          kind: 'external',
          url: 'mailto:new@example.invalid',
        }),
        hyperlink('C3', { kind: 'internal', location: 'Sheet2!A1' }),
      ],
      SHEET,
    );
    expect([...plan.idsByCell]).toEqual([
      ['A1', 'rId2'],
      ['B2', 'rId10'],
    ]);
    expect([...plan.targets]).toEqual([
      ['rId2', 'https://new.invalid/'],
      ['rId10', 'mailto:new@example.invalid'],
    ]);
    expect([...plan.removeIds]).toEqual([]);
    expect(plan.changed).toBe(true);
  });

  it('does not reuse a worksheet ID owned by a non-hyperlink relationship', () => {
    const worksheet = bytes(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><hyperlinks><hyperlink ref="A1" r:id="rId2"/></hyperlinks></worksheet>`,
    );
    const plan = planXlsxExternalHyperlinkRelationships(
      worksheet,
      [
        {
          id: 'rId2',
          mode: 'external',
          owner: SHEET,
          target: 'https://old.invalid/',
          type: 'urn:not-a-hyperlink',
        },
      ],
      [hyperlink('A1', { kind: 'external', url: 'https://new.invalid/' })],
      SHEET,
    );
    expect([...plan.idsByCell]).toEqual([['A1', 'rId3']]);
    expect([...plan.removeIds]).toEqual(['rId2']);
    expect([...plan.targets]).toEqual([['rId3', 'https://new.invalid/']]);
  });

  it('distinguishes unchanged targets from removal-only plans', () => {
    const worksheet = bytes(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><hyperlinks><hyperlink ref="A1" r:id="rId2"/></hyperlinks></worksheet>`,
    );
    const relationships: XlsxPackageGraphRelationship[] = [
      {
        id: 'rId2',
        mode: 'external',
        owner: SHEET,
        target: 'https://same.invalid/',
        type: TYPE,
      },
    ];
    const unchanged = planXlsxExternalHyperlinkRelationships(
      worksheet,
      relationships,
      [hyperlink('A1', { kind: 'external', url: 'https://same.invalid/' })],
      SHEET,
    );
    expect(unchanged.changed).toBe(false);
    expect([...unchanged.removeIds]).toEqual([]);
    const removed = planXlsxExternalHyperlinkRelationships(
      worksheet,
      relationships,
      [hyperlink('A1', { kind: 'internal', location: 'Sheet2!A1' })],
      SHEET,
    );
    expect(removed.changed).toBe(true);
    expect([...removed.removeIds]).toEqual(['rId2']);
    expect([...removed.targets]).toEqual([]);
  });

  it('marks a mixed unchanged and changed target plan as changed', () => {
    const worksheet = bytes(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}"><hyperlinks><hyperlink ref="A1" r:id="rId1"/><hyperlink ref="B2" r:id="rId2"/></hyperlinks></worksheet>`,
    );
    const relationships: XlsxPackageGraphRelationship[] = [
      {
        id: 'rId1',
        mode: 'external',
        owner: SHEET,
        target: 'https://same.invalid/',
        type: TYPE,
      },
      {
        id: 'rId2',
        mode: 'external',
        owner: SHEET,
        target: 'https://old.invalid/',
        type: TYPE,
      },
    ];
    const plan = planXlsxExternalHyperlinkRelationships(
      worksheet,
      relationships,
      [
        hyperlink('A1', { kind: 'external', url: 'https://same.invalid/' }),
        hyperlink('B2', { kind: 'external', url: 'https://new.invalid/' }),
      ],
      SHEET,
    );
    expect(plan.changed).toBe(true);
  });

  it('ignores nested and foreign worksheet hyperlink ID lookalikes', () => {
    const worksheet = bytes(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_NS}" xmlns:x="urn:foreign"><x:hyperlinks><x:hyperlink ref="A1" r:id="rId99"/></x:hyperlinks><hyperlinks><wrapper><hyperlink ref="A1" r:id="rId88"/></wrapper><hyperlink ref="A1" r:id="rId2"/></hyperlinks></worksheet>`,
    );
    const relationships: XlsxPackageGraphRelationship[] = [
      {
        id: 'rId2',
        mode: 'external',
        owner: SHEET,
        target: 'https://same.invalid/',
        type: TYPE,
      },
    ];
    const plan = planXlsxExternalHyperlinkRelationships(
      worksheet,
      relationships,
      [
        hyperlink('A1', { kind: 'external', url: 'https://same.invalid/' }),
        hyperlink('B2', { kind: 'external', url: 'https://new.invalid/' }),
      ],
      SHEET,
    );
    expect([...plan.idsByCell]).toEqual([
      ['A1', 'rId2'],
      ['B2', 'rId3'],
    ]);
  });

  it('patches changed targets, removes owned IDs, adds new IDs, and preserves unrelated nodes', () => {
    const source = bytes(
      `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship TargetMode="External" Target="https://old.invalid/" Type="${TYPE}" Id="rId2"/><Relationship Id="keep" Type="urn:keep" Target="../keep.xml"/><Relationship Id="rId4" Type="${TYPE}" Target="https://remove.invalid/" TargetMode="External"/></Relationships>`,
    );
    const result = patchXlsxHyperlinkRelationships(
      source,
      {
        changed: true,
        idsByCell: new Map(),
        removeIds: new Set(['rId4']),
        targets: new Map([
          ['rId2', 'https://updated.invalid/'],
          ['rId3', 'https://new.invalid/?a=1&b=2'],
        ]),
      },
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    const xml = new TextDecoder().decode(result.data);
    expect(xml).toBe(
      `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="rId2" Type="${TYPE}" Target="https://updated.invalid/" TargetMode="External"/><Relationship Id="keep" Type="urn:keep" Target="../keep.xml"/><Relationship Id="rId3" Type="${TYPE}" Target="https://new.invalid/?a=1&amp;b=2" TargetMode="External"/></Relationships>`,
    );
    expect(result.patchCount).toBe(3);
  });

  it('creates a deterministic relationship part and expands a self-closing root', () => {
    const plan = {
      changed: true,
      idsByCell: new Map([['A1', 'rId1']]),
      removeIds: new Set<string>(),
      targets: new Map([['rId1', 'https://example.invalid/']]),
    };
    const created = patchXlsxHyperlinkRelationships(
      null,
      plan,
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(new TextDecoder().decode(created.data)).toBe(
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="rId1" Type="${TYPE}" Target="https://example.invalid/" TargetMode="External"/></Relationships>`,
    );
    expect(created.patchCount).toBe(1);
    const expanded = patchXlsxHyperlinkRelationships(
      bytes(
        `<?xml version="1.0"?><Relationships xmlns="${XLSX_PACKAGE_REL_NS}"/>`,
      ),
      plan,
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(new TextDecoder().decode(expanded.data)).toBe(
      `<?xml version="1.0"?><Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="rId1" Type="${TYPE}" Target="https://example.invalid/" TargetMode="External"/></Relationships>`,
    );
    const trapped = patchXlsxHyperlinkRelationships(
      bytes(`<Relationships xmlns="${XLSX_PACKAGE_REL_NS}" note="/>" /   >`),
      plan,
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(new TextDecoder().decode(trapped.data)).toBe(
      `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}" note="/>" ><Relationship Id="rId1" Type="${TYPE}" Target="https://example.invalid/" TargetMode="External"/></Relationships>`,
    );
  });

  it('creates an empty relationship part without reporting a patch', () => {
    const result = patchXlsxHyperlinkRelationships(
      null,
      {
        changed: false,
        idsByCell: new Map(),
        removeIds: new Set(),
        targets: new Map(),
      },
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    const expected = `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${XLSX_PACKAGE_REL_NS}"></Relationships>`;
    expect(new TextDecoder().decode(result.data)).toBe(expected);
    expect(result.patchBytes).toBe(bytes(expected).byteLength);
    expect(result.patchCount).toBe(0);
  });

  it('joins multiple generated relationships without text between them', () => {
    const result = patchXlsxHyperlinkRelationships(
      null,
      {
        changed: true,
        idsByCell: new Map(),
        removeIds: new Set(),
        targets: new Map([
          ['rId1', 'https://one.invalid/'],
          ['rId2', 'https://two.invalid/'],
        ]),
      },
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toContain(
      `Target="https://one.invalid/" TargetMode="External"/><Relationship Id="rId2" Type="${TYPE}"`,
    );
  });

  it('patches one existing target without an addition', () => {
    const source = bytes(
      `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="rId1" Type="${TYPE}" Target="https://old.invalid/" TargetMode="External"/></Relationships>`,
    );
    const result = patchXlsxHyperlinkRelationships(
      source,
      {
        changed: true,
        idsByCell: new Map([['A1', 'rId1']]),
        removeIds: new Set(),
        targets: new Map([['rId1', 'https://new.invalid/']]),
      },
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(result.patchCount).toBe(1);
    expect(new TextDecoder().decode(result.data)).toContain(
      'Target="https://new.invalid/"',
    );
  });

  it('patches only direct Relationship elements at the root depth', () => {
    const source = bytes(
      `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><wrapper><Relationship Id="rId1" Type="${TYPE}" Target="https://nested.invalid/" TargetMode="External"/></wrapper><Other Id="rId1"/><Relationship TargetMode="External" Target="https://old.invalid/" Type="${TYPE}" Id="rId1"/></Relationships>`,
    );
    const result = patchXlsxHyperlinkRelationships(
      source,
      {
        changed: true,
        idsByCell: new Map(),
        removeIds: new Set(),
        targets: new Map([['rId1', 'https://updated.invalid/']]),
      },
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    const xml = new TextDecoder().decode(result.data);
    expect(xml).toContain('Target="https://nested.invalid/"');
    expect(xml).toContain('<Other Id="rId1"/>');
    expect(xml).toContain('Target="https://updated.invalid/"');
  });

  it('preserves non-Relationship elements even when their attributes look owned', () => {
    const source = bytes(
      `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Other Id="rId1" Type="${TYPE}" Target="https://foreign.invalid/"/></Relationships>`,
    );
    const result = patchXlsxHyperlinkRelationships(
      source,
      {
        changed: true,
        idsByCell: new Map(),
        removeIds: new Set(),
        targets: new Map([['rId1', 'https://added.invalid/']]),
      },
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Other Id="rId1" Type="${TYPE}" Target="https://foreign.invalid/"/><Relationship Id="rId1" Type="${TYPE}" Target="https://added.invalid/" TargetMode="External"/></Relationships>`,
    );
  });

  it('patches prefixed relationship roots with matching element names', () => {
    const source = bytes(
      `<p:Relationships xmlns:p="${XLSX_PACKAGE_REL_NS}"><p:Relationship Id="rId1" Type="${TYPE}" Target="https://old.invalid/" TargetMode="External"/></p:Relationships>`,
    );
    const result = patchXlsxHyperlinkRelationships(
      source,
      {
        changed: true,
        idsByCell: new Map(),
        removeIds: new Set(),
        targets: new Map([['rId1', 'https://new.invalid/']]),
      },
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<p:Relationships xmlns:p="${XLSX_PACKAGE_REL_NS}"><p:Relationship Id="rId1" Type="${TYPE}" Target="https://new.invalid/" TargetMode="External"/></p:Relationships>`,
    );
  });

  it('accounts for removals before enforcing the final relationship limit', () => {
    const source = bytes(
      `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="rId1" Type="${TYPE}" Target="https://remove.invalid/" TargetMode="External"/><Relationship Id="keep" Type="urn:keep" Target="../keep.xml"/></Relationships>`,
    );
    const result = patchXlsxHyperlinkRelationships(
      source,
      {
        changed: true,
        idsByCell: new Map(),
        removeIds: new Set(['rId1']),
        targets: new Map([['rId2', 'https://added.invalid/']]),
      },
      TYPE,
      defaultXlsxWriteLimits(),
      { ...defaultXlsxResourceLimits(), maxRelationships: 2 },
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="keep" Type="urn:keep" Target="../keep.xml"/><Relationship Id="rId2" Type="${TYPE}" Target="https://added.invalid/" TargetMode="External"/></Relationships>`,
    );
  });

  it('appends multiple relationships to an existing root in map order', () => {
    const result = patchXlsxHyperlinkRelationships(
      bytes(`<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"></Relationships>`),
      {
        changed: true,
        idsByCell: new Map(),
        removeIds: new Set(),
        targets: new Map([
          ['rId1', 'https://one.invalid/'],
          ['rId2', 'https://two.invalid/'],
        ]),
      },
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toBe(
      `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="rId1" Type="${TYPE}" Target="https://one.invalid/" TargetMode="External"/><Relationship Id="rId2" Type="${TYPE}" Target="https://two.invalid/" TargetMode="External"/></Relationships>`,
    );
  });

  it('escapes every relationship target attribute character', () => {
    const result = patchXlsxHyperlinkRelationships(
      null,
      {
        changed: true,
        idsByCell: new Map(),
        removeIds: new Set(),
        targets: new Map([['rId1', '&<"\r\n\t']]),
      },
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(new TextDecoder().decode(result.data)).toContain(
      'Target="&amp;&lt;&quot;&#13;&#10;&#9;"',
    );
  });

  it('enforces relationship and generated-byte limits exactly', () => {
    const plan = {
      changed: true,
      idsByCell: new Map([['A1', 'rId1']]),
      removeIds: new Set<string>(),
      targets: new Map([['rId1', 'https://example.invalid/']]),
    };
    const successful = patchXlsxHyperlinkRelationships(
      null,
      plan,
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(() =>
      patchXlsxHyperlinkRelationships(
        null,
        plan,
        TYPE,
        {
          ...defaultXlsxWriteLimits(),
          maxGeneratedXmlBytes: successful.data.byteLength,
        },
        { ...defaultXlsxResourceLimits(), maxRelationships: 1 },
        PART,
      ),
    ).not.toThrow();
    expect(
      capture(() =>
        patchXlsxHyperlinkRelationships(
          null,
          plan,
          TYPE,
          {
            ...defaultXlsxWriteLimits(),
            maxGeneratedXmlBytes: successful.data.byteLength - 1,
          },
          defaultXlsxResourceLimits(),
          PART,
        ),
      ).diagnostic.limitName,
    ).toBe('maxGeneratedXmlBytes');
    const relationshipLimitError = capture(() =>
      patchXlsxHyperlinkRelationships(
        null,
        plan,
        TYPE,
        defaultXlsxWriteLimits(),
        { ...defaultXlsxResourceLimits(), maxRelationships: 1 - 1 },
        PART,
      ),
    );
    expect(relationshipLimitError.diagnostic).toMatchObject({
      limitName: 'maxRelationships',
      message: 'XLSX generated relationships exceed the reader limit',
    });
  });

  it('enforces existing-part aggregate limits and rejects invalid roots', () => {
    const source = bytes(
      `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}"><Relationship Id="rId1" Type="${TYPE}" Target="https://old.invalid/" TargetMode="External"/></Relationships>`,
    );
    const plan = {
      changed: true,
      idsByCell: new Map(),
      removeIds: new Set<string>(),
      targets: new Map([
        ['rId1', 'https://updated.invalid/'],
        ['rId2', 'https://added.invalid/'],
      ]),
    };
    const successful = patchXlsxHyperlinkRelationships(
      source,
      plan,
      TYPE,
      defaultXlsxWriteLimits(),
      defaultXlsxResourceLimits(),
      PART,
    );
    expect(successful.patchBytes).toBe(330);
    expect(() =>
      patchXlsxHyperlinkRelationships(
        source,
        plan,
        TYPE,
        {
          ...defaultXlsxWriteLimits(),
          maxGeneratedXmlBytes: successful.data.byteLength,
          maxPatchBytes: successful.patchBytes,
        },
        { ...defaultXlsxResourceLimits(), maxRelationships: 2 },
        PART,
      ),
    ).not.toThrow();
    for (const [limitName, limit] of [
      ['maxGeneratedXmlBytes', successful.data.byteLength - 1],
      ['maxPatchBytes', successful.patchBytes - 1],
    ] as const) {
      expect(
        capture(() =>
          patchXlsxHyperlinkRelationships(
            source,
            plan,
            TYPE,
            { ...defaultXlsxWriteLimits(), [limitName]: limit },
            defaultXlsxResourceLimits(),
            PART,
          ),
        ).diagnostic,
      ).toMatchObject({ limit, limitName, part: PART });
    }
    expect(
      capture(() =>
        patchXlsxHyperlinkRelationships(
          source,
          plan,
          TYPE,
          defaultXlsxWriteLimits(),
          { ...defaultXlsxResourceLimits(), maxRelationships: 1 },
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({
      actual: 2,
      limit: 1,
      limitName: 'maxRelationships',
      message: 'XLSX generated relationships exceed the reader limit',
    });
    expect(
      capture(() =>
        patchXlsxHyperlinkRelationships(
          bytes('<outer/>'),
          plan,
          TYPE,
          defaultXlsxWriteLimits(),
          defaultXlsxResourceLimits(),
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({
      featureClass: 'relationship-xml',
      message: 'XLSX relationship root cannot be patched',
    });
    expect(
      capture(() =>
        patchXlsxHyperlinkRelationships(
          bytes(
            `<outer><Relationships xmlns="${XLSX_PACKAGE_REL_NS}"/></outer>`,
          ),
          plan,
          TYPE,
          defaultXlsxWriteLimits(),
          defaultXlsxResourceLimits(),
          PART,
        ),
      ).diagnostic,
    ).toMatchObject({
      featureClass: 'relationship-xml',
      message: 'XLSX relationship root cannot be patched',
    });
  });

  it('rejects exhausted relationship ID allocation', () => {
    const worksheet = bytes(
      `<worksheet xmlns="${XLSX_SPREADSHEET_NS}"><hyperlinks/></worksheet>`,
    );
    expect(
      capture(() =>
        planXlsxExternalHyperlinkRelationships(
          worksheet,
          [
            {
              id: `rId${Number.MAX_SAFE_INTEGER}`,
              mode: 'external',
              owner: SHEET,
              target: 'https://old.invalid/',
              type: TYPE,
            },
          ],
          [hyperlink('A1', { kind: 'external', url: 'https://new.invalid/' })],
          SHEET,
        ),
      ).diagnostic,
    ).toMatchObject({
      code: 'identifier-allocation-failed',
      featureClass: 'relationship-id',
      message: 'XLSX relationship IDs are exhausted',
    });
  });
});
