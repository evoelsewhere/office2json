import { describe, expect, it } from 'vitest';

import {
  applyXlsxEdits,
  createXlsxCapabilityManifest,
  readXlsxRoundTrip,
  validateXlsxRoundTripJson,
  writeXlsxRoundTrip,
  XlsxWriteError,
} from '../../src/formats/xlsx';
import type { XlsxRoundTripSnapshot } from '../../src/formats/xlsx/roundtrip';
import { sha256XlsxBytes } from '../../src/formats/xlsx/roundtrip/digest';
import {
  createIndependentXlsx,
  XLSX_CONTENT_TYPES_NS,
  XLSX_OFFICE_REL_TYPE,
  XLSX_PACKAGE_REL_NS,
  XLSX_SPREADSHEET_NS,
} from '../black-box/xlsx-package';

function portableClone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

async function captureWriteError(
  action: () => Promise<unknown>,
): Promise<XlsxWriteError> {
  try {
    await action();
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxWriteError);
    return error as XlsxWriteError;
  }
  throw new Error('Expected XLSX round-trip action to fail');
}

describe('XLSX exact R0 round-trip', () => {
  it('round-trips standalone JSON to byte-identical source with copy evidence', async () => {
    const bytes = await createIndependentXlsx();
    const before = bytes.slice();
    const snapshot = await readXlsxRoundTrip(bytes);

    expect(snapshot).toMatchObject({
      format: 'xlsx-roundtrip',
      keyAlgorithmVersion: 'xlsx-snapshot-key-v1',
      operations: [],
      preservation: {
        containsActiveContent: false,
        containsDigitalSignatures: false,
        containsExternalRelationships: false,
        securityMode: 'reject-active',
      },
      schemaVersion: '1',
      source: {
        byteLength: bytes.byteLength,
        conformance: 'transitional',
      },
      supportProfile: {
        effectiveLevel: 'R2',
        id: 'xlsx-agent-ready',
        producerEvidence: [],
        version: '1',
      },
    });
    expect(snapshot.source.sha256).toMatch(/^[0-9a-f]{64}$/u);
    expect(snapshot.baseDocumentHash).toMatch(/^[0-9a-f]{64}$/u);
    expect(snapshot.stateHash).toBe(snapshot.baseDocumentHash);
    expect(snapshot.sourceManifestHash).toMatch(/^[0-9a-f]{64}$/u);
    expect(snapshot.document.key).toMatch(/^xlsx:workbook:[0-9a-f]{32}$/u);
    expect(snapshot.document.sheets[0]?.key).toMatch(
      /^xlsx:sheet:[0-9a-f]{32}$/u,
    );
    expect(bytes).toEqual(before);

    const json = JSON.stringify(snapshot);
    const validated = await validateXlsxRoundTripJson(JSON.parse(json));
    const result = await writeXlsxRoundTrip(validated);
    expect(result.data).toEqual(bytes);
    expect(result.data).not.toBe(bytes);
    expect(result.report).toMatchObject({
      diagnostics: [],
      level: 'R0',
      outputSha256: snapshot.source.sha256,
      sourceSha256: snapshot.source.sha256,
    });
    expect(result.report.parts.length).toBeGreaterThan(0);
    expect(
      result.report.parts.every((part) => part.disposition === 'copy'),
    ).toBe(true);
    expect(result.report.parts.map((part) => part.name)).toEqual(
      [...result.report.parts.map((part) => part.name)].sort(),
    );
    expect(await sha256XlsxBytes(result.data)).toBe(snapshot.source.sha256);
  });

  it('treats Uint8Array views, ArrayBuffer, and Blob inputs identically', async () => {
    const bytes = await createIndependentXlsx();
    const padded = new Uint8Array(bytes.byteLength + 2);
    padded.set(bytes, 1);
    const inputs = [
      bytes,
      padded.subarray(1, padded.byteLength - 1),
      bytes.buffer.slice(
        bytes.byteOffset,
        bytes.byteOffset + bytes.byteLength,
      ) as ArrayBuffer,
      new Blob([
        bytes.buffer.slice(
          bytes.byteOffset,
          bytes.byteOffset + bytes.byteLength,
        ) as ArrayBuffer,
      ]),
    ];
    const snapshots = await Promise.all(
      inputs.map((input) => readXlsxRoundTrip(input)),
    );
    for (const snapshot of snapshots) expect(snapshot).toEqual(snapshots[0]);
  });

  it('is deterministic and isolated across repeated concurrent reads and writes', async () => {
    const bytes = await createIndependentXlsx();
    const snapshots = await Promise.all(
      Array.from({ length: 8 }, () => readXlsxRoundTrip(bytes)),
    );
    for (const snapshot of snapshots) expect(snapshot).toEqual(snapshots[0]);
    expect(new Set(snapshots).size).toBe(snapshots.length);
    const outputs = await Promise.all(
      snapshots.map((item) => writeXlsxRoundTrip(item)),
    );
    for (const output of outputs) expect(output.data).toEqual(bytes);
    expect(new Set(outputs.map((output) => output.data)).size).toBe(
      outputs.length,
    );
  });

  it.each([
    [
      (snapshot: XlsxRoundTripSnapshot) => {
        snapshot.source.sha256 = '0'.repeat(64);
      },
      'source-package-mismatch',
      'XLSX source bytes do not match their snapshot identity',
    ],
    [
      (snapshot: XlsxRoundTripSnapshot) => {
        const first = snapshot.source.packageBase64[0]!;
        snapshot.source.packageBase64 = `${first === 'A' ? 'B' : 'A'}${snapshot.source.packageBase64.slice(1)}`;
      },
      'source-package-mismatch',
      'XLSX source bytes do not match their snapshot identity',
    ],
    [
      (snapshot: XlsxRoundTripSnapshot) => {
        snapshot.sourceManifestHash = '0'.repeat(64);
      },
      'source-package-mismatch',
      'XLSX source package graph does not match the snapshot',
    ],
    [
      (snapshot: XlsxRoundTripSnapshot) => {
        snapshot.source.conformance = 'strict';
      },
      'source-package-mismatch',
      'XLSX source package graph does not match the snapshot',
    ],
    [
      (snapshot: XlsxRoundTripSnapshot) => {
        snapshot.baseDocumentHash = '0'.repeat(64);
      },
      'snapshot-integrity-failed',
      'XLSX semantic preview does not match its source',
    ],
    [
      (snapshot: XlsxRoundTripSnapshot) => {
        snapshot.stateHash = '0'.repeat(64);
      },
      'snapshot-integrity-failed',
      'XLSX semantic preview does not match its source',
    ],
    [
      (snapshot: XlsxRoundTripSnapshot) => {
        snapshot.document.workbook.dateSystem = '1904';
      },
      'snapshot-integrity-failed',
      'XLSX semantic preview does not match its source',
    ],
    [
      (snapshot: XlsxRoundTripSnapshot) => {
        snapshot.preservation.containsOpaqueContent =
          !snapshot.preservation.containsOpaqueContent;
      },
      'source-package-mismatch',
      'XLSX preservation inventory does not match the source',
    ],
  ] as const)(
    'rejects tampered bound snapshot %#',
    async (tamper, code, message) => {
      const snapshot = portableClone(
        await readXlsxRoundTrip(await createIndependentXlsx()),
      );
      tamper(snapshot);
      const error = await captureWriteError(() => writeXlsxRoundTrip(snapshot));
      expect(error.diagnostic).toMatchObject({ code, message });
    },
  );

  it('rejects source bytes, Base64, shape, versions, profiles, and operations', async () => {
    const source = await readXlsxRoundTrip(await createIndependentXlsx());

    const badBase64 = portableClone(source);
    badBase64.source.packageBase64 = `!${badBase64.source.packageBase64.slice(1)}`;
    expect(
      (await captureWriteError(() => validateXlsxRoundTripJson(badBase64)))
        .diagnostic.code,
    ).toBe('invalid-roundtrip-json');

    const badLength = portableClone(source);
    badLength.source.byteLength += 1;
    expect(
      (await captureWriteError(() => validateXlsxRoundTripJson(badLength)))
        .diagnostic.code,
    ).toBe('invalid-roundtrip-json');

    const unknownRoot = portableClone(source) as XlsxRoundTripSnapshot & {
      unknown?: boolean;
    };
    unknownRoot.unknown = true;
    expect(
      (await captureWriteError(() => validateXlsxRoundTripJson(unknownRoot)))
        .diagnostic.code,
    ).toBe('invalid-roundtrip-json');

    const version = portableClone(source);
    version.schemaVersion = '2' as '1';
    expect(
      (await captureWriteError(() => validateXlsxRoundTripJson(version)))
        .diagnostic.code,
    ).toBe('unsupported-snapshot-version');

    const profile = portableClone(source);
    profile.supportProfile.version = '2' as '1';
    expect(
      (await captureWriteError(() => validateXlsxRoundTripJson(profile)))
        .diagnostic.code,
    ).toBe('unsupported-snapshot-version');

    const operation = portableClone(source);
    operation.operations.push({
      cell: 'A1',
      kind: 'clear-cell',
      operationId: 'operation-1',
      sheetKey: operation.document.sheets[0]!.key,
    });
    expect(
      (await captureWriteError(() => validateXlsxRoundTripJson(operation)))
        .diagnostic.code,
    ).toBe('snapshot-integrity-failed');
  });

  it('enforces source and snapshot limits at exact boundaries', async () => {
    const bytes = await createIndependentXlsx();
    const snapshot = await readXlsxRoundTrip(bytes);
    await expect(
      validateXlsxRoundTripJson(portableClone(snapshot), {
        limits: {
          maxOutputBytes: bytes.byteLength,
          maxSourcePackageBytes: bytes.byteLength,
        },
      }),
    ).resolves.toEqual(snapshot);
    const error = await captureWriteError(() =>
      validateXlsxRoundTripJson(portableClone(snapshot), {
        limits: {
          maxOutputBytes: bytes.byteLength,
          maxSourcePackageBytes: bytes.byteLength - 1,
        },
      }),
    );
    expect(error.diagnostic).toMatchObject({
      actual: bytes.byteLength,
      code: 'resource-limit-exceeded',
      limit: bytes.byteLength - 1,
      limitName: 'maxSourcePackageBytes',
    });
  });

  it('publishes a deterministic constrained R2 cell capability manifest', () => {
    const manifest = createXlsxCapabilityManifest();
    expect(manifest.domains.length).toBeGreaterThan(20);
    expect(manifest.effectiveLevel).toBe('R2');
    expect(
      manifest.domains.filter((entry) => entry.level === 'verified-R2'),
    ).toEqual([
      { domain: 'cells', level: 'verified-R2' },
      { domain: 'formulas', level: 'verified-R2' },
      { domain: 'hyperlinks', level: 'verified-R2' },
      { domain: 'rows-columns', level: 'verified-R2' },
      { domain: 'styles', level: 'verified-R2' },
    ]);
    expect(
      manifest.operations.filter((entry) => entry.level === 'verified-R2'),
    ).toEqual([
      expect.objectContaining({ operation: 'clear-cell' }),
      expect.objectContaining({ operation: 'delete-columns' }),
      expect.objectContaining({ operation: 'delete-rows' }),
      expect.objectContaining({ operation: 'insert-columns' }),
      expect.objectContaining({ operation: 'insert-rows' }),
      expect.objectContaining({ operation: 'set-cell' }),
      expect.objectContaining({ operation: 'set-cell-style' }),
      expect.objectContaining({ operation: 'set-column' }),
      expect.objectContaining({ operation: 'set-hyperlink' }),
      expect.objectContaining({ operation: 'set-row' }),
    ]);
    expect(
      manifest.operations
        .filter(
          (entry) =>
            entry.operation === 'clear-cell' || entry.operation === 'set-cell',
        )
        .every(
          (entry) =>
            JSON.stringify(entry.constraints) ===
            JSON.stringify([
              'existing-explicit-cell',
              'clean-supported-package-closure',
              'no-unaffected-formulas-or-defined-names',
              'no-grouped-formula-target',
              'no-date-or-rich-text-value',
              'no-external-capable-formula',
            ]),
        ),
    ).toBe(true);
    expect(
      manifest.operations.find((entry) => entry.operation === 'set-cell-style'),
    ).toEqual({
      constraints: [
        'existing-explicit-cell',
        'existing-or-append-normalized-style',
        'existing-styles-part-for-append',
        'no-new-checkbox-style',
        'clean-supported-package-closure',
        'no-non-anchor-merged-cell',
      ],
      level: 'verified-R2',
      operation: 'set-cell-style',
    });
    expect(
      manifest.operations
        .filter((entry) =>
          [
            'delete-columns',
            'delete-rows',
            'insert-columns',
            'insert-rows',
          ].includes(entry.operation),
        )
        .every(
          (entry) =>
            JSON.stringify(entry.constraints) ===
            JSON.stringify([
              'reference-free-simple-worksheet',
              'explicit-row-and-cell-references',
              'structural-operations-only-batch',
              'no-explicit-column-definitions-for-column-shifts',
              'no-grid-overflow',
              'clean-supported-package-closure',
            ]),
        ),
    ).toBe(true);
    expect(
      manifest.operations.find((entry) => entry.operation === 'set-column'),
    ).toEqual({
      constraints: [
        'existing-exact-column-range',
        'size-and-visibility-only',
        'clean-supported-package-closure',
      ],
      level: 'verified-R2',
      operation: 'set-column',
    });
    expect(
      manifest.operations.find((entry) => entry.operation === 'set-hyperlink'),
    ).toEqual({
      constraints: [
        'existing-explicit-cell',
        'safe-internal-or-external-target',
        'deterministic-relationship-allocation',
        'no-overlapping-multi-cell-hyperlink',
        'http-https-mailto-only',
        'no-url-credentials',
        'clean-supported-package-closure',
      ],
      level: 'verified-R2',
      operation: 'set-hyperlink',
    });
    expect(
      manifest.operations.find((entry) => entry.operation === 'set-row'),
    ).toEqual({
      constraints: [
        'existing-explicit-row',
        'size-and-visibility-only',
        'clean-supported-package-closure',
      ],
      level: 'verified-R2',
      operation: 'set-row',
    });
    expect(new Set(manifest.domains.map((entry) => entry.domain)).size).toBe(
      manifest.domains.length,
    );
    expect(
      new Set(manifest.operations.map((entry) => entry.operation)).size,
    ).toBe(manifest.operations.length);
    expect(JSON.parse(JSON.stringify(manifest))).toEqual(manifest);
  });

  it('accepts empty and supported typed edit batches honestly', async () => {
    const snapshot = await readXlsxRoundTrip(await createIndependentXlsx());
    await expect(applyXlsxEdits(snapshot, [])).resolves.toEqual(snapshot);
    await expect(
      applyXlsxEdits(snapshot, [
        {
          cell: 'A1',
          kind: 'clear-cell',
          operationId: 'clear-a1',
          sheetKey: snapshot.document.sheets[0]!.key,
        },
      ]),
    ).resolves.toMatchObject({
      operations: [{ kind: 'clear-cell', operationId: 'clear-a1' }],
    });
    await expect(applyXlsxEdits(snapshot, {} as never)).rejects.toThrow(
      'XLSX edit operations must be an array',
    );
  });

  it('rejects active content by default and preserves it exactly when acknowledged by policy', async () => {
    const contentTypes = `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
      <Override PartName="/xl/activeX/activeX1.bin" ContentType="application/vnd.ms-office.activeX"/>
    </Types>`;
    const bytes = await createIndependentXlsx({
      '[Content_Types].xml': contentTypes,
      'xl/activeX/activeX1.bin': new Uint8Array([1, 2, 3]),
    });
    expect(
      (await captureWriteError(() => readXlsxRoundTrip(bytes))).diagnostic,
    ).toMatchObject({
      code: 'preservation-conflict',
      featureClass: 'active-content',
      message: 'XLSX round-trip source contains active or embedded content',
    });

    const snapshot = await readXlsxRoundTrip(bytes, {
      securityMode: 'preserve-opaque',
    });
    expect(snapshot.preservation).toMatchObject({
      containsActiveContent: true,
      containsOpaqueContent: true,
      securityMode: 'preserve-opaque',
    });
    await expect(
      writeXlsxRoundTrip(portableClone(snapshot)),
    ).resolves.toMatchObject({
      data: bytes,
      report: { level: 'R0' },
    });
    const rejected = portableClone(snapshot);
    rejected.preservation.securityMode = 'reject-active';
    expect(
      (await captureWriteError(() => writeXlsxRoundTrip(rejected))).diagnostic,
    ).toMatchObject({
      code: 'preservation-conflict',
      featureClass: 'active-content',
      message: 'XLSX source contains active or embedded content',
    });
  });

  it('round-trips a Strict source and records exact conformance', async () => {
    const strictSheetNs = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const strictRelNs =
      'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const bytes = await createIndependentXlsx({
      '[Content_Types].xml': `<Types xmlns="${XLSX_CONTENT_TYPES_NS}">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      </Types>`,
      '_rels/.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
        <Relationship Id="root" Type="${strictRelNs}/officeDocument" Target="xl/workbook.xml"/>
      </Relationships>`,
      'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
        <Relationship Id="sheet" Type="${strictRelNs}/worksheet" Target="worksheets/sheet1.xml"/>
      </Relationships>`,
      'xl/sharedStrings.xml': null,
      'xl/styles.xml': null,
      'xl/workbook.xml': `<s:workbook xmlns:s="${strictSheetNs}" xmlns:r="${strictRelNs}">
        <s:sheets><s:sheet name="Strict" sheetId="1" r:id="sheet"/></s:sheets>
      </s:workbook>`,
      'xl/worksheets/sheet1.xml': `<s:worksheet xmlns:s="${strictSheetNs}"><s:sheetData/></s:worksheet>`,
    });
    const snapshot = await readXlsxRoundTrip(bytes);
    expect(snapshot.source.conformance).toBe('strict');
    expect((await writeXlsxRoundTrip(portableClone(snapshot))).data).toEqual(
      bytes,
    );
  });

  it('inventories safe external relationships without dereferencing them', async () => {
    const worksheet = `<worksheet xmlns="${XLSX_SPREADSHEET_NS}" xmlns:r="${XLSX_OFFICE_REL_TYPE.slice(0, -1)}">
      <sheetData/><hyperlinks><hyperlink ref="A1" r:id="link"/></hyperlinks>
    </worksheet>`;
    const bytes = await createIndependentXlsx({
      'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${XLSX_PACKAGE_REL_NS}">
        <Relationship Id="link" Type="${XLSX_OFFICE_REL_TYPE}hyperlink"
          Target="https://example.com/never-fetched" TargetMode="External"/>
      </Relationships>`,
      'xl/worksheets/sheet1.xml': worksheet,
    });
    const snapshot = await readXlsxRoundTrip(bytes);
    expect(snapshot.preservation.containsExternalRelationships).toBe(true);
    expect((await writeXlsxRoundTrip(snapshot)).data).toEqual(bytes);
  });

  it('preserves signed-package parts exactly at R0', async () => {
    const bytes = await createIndependentXlsx({
      '_xmlsignatures/sig1.xml': '<Signature>opaque-signature</Signature>',
    });
    const snapshot = await readXlsxRoundTrip(bytes);
    expect(snapshot.preservation.containsDigitalSignatures).toBe(true);
    const result = await writeXlsxRoundTrip(portableClone(snapshot));
    expect(result.data).toEqual(bytes);
    expect(
      result.report.parts.find(
        (part) => part.name === '_xmlsignatures/sig1.xml',
      ),
    ).toMatchObject({ disposition: 'copy' });
  });

  it('validates read and write option objects and values exactly', async () => {
    const bytes = await createIndependentXlsx();
    const snapshot = await readXlsxRoundTrip(bytes);
    const invalidReads: Array<readonly [unknown, string]> = [
      [null, 'XLSX round-trip read options must be a plain object'],
      [[], 'XLSX round-trip read options must be a plain object'],
      [new Date(), 'XLSX round-trip read options must be a plain object'],
      [{ unknown: true }, 'Unknown XLSX round-trip read option unknown'],
      [{ securityMode: 'unsafe' }, 'XLSX round-trip security mode is invalid'],
    ];
    invalidReads.push([
      Object.defineProperty({}, Symbol.toStringTag, { value: 'Spoofed' }),
      'XLSX round-trip read options must be a plain object',
    ]);
    invalidReads.push([
      Object.create(null),
      'XLSX round-trip read options must be a plain object',
    ]);
    for (const [options, message] of invalidReads) {
      await expect(readXlsxRoundTrip(bytes, options as never)).rejects.toThrow(
        message,
      );
    }
    for (const securityMode of ['preserve-opaque', 'reject-active'] as const) {
      await expect(
        readXlsxRoundTrip(bytes, { securityMode }),
      ).resolves.toMatchObject({ preservation: { securityMode } });
    }
    const invalidWrites: Array<readonly [unknown, string]> = [
      [null, 'XLSX write options must be a plain object'],
      [[], 'XLSX write options must be a plain object'],
      [new Date(), 'XLSX write options must be a plain object'],
      [{ unknown: true }, 'Unknown XLSX write option unknown'],
      [
        { acknowledgeOpaqueContent: 'yes' },
        'XLSX acknowledgeOpaqueContent must be boolean',
      ],
      [
        { minimumEditedFidelity: 'R0' },
        'XLSX minimumEditedFidelity is invalid',
      ],
    ];
    invalidWrites.push([
      Object.defineProperty({}, Symbol.toStringTag, { value: 'Spoofed' }),
      'XLSX write options must be a plain object',
    ]);
    invalidWrites.push([
      Object.create(null),
      'XLSX write options must be a plain object',
    ]);
    for (const [options, message] of invalidWrites) {
      await expect(
        writeXlsxRoundTrip(snapshot, options as never),
      ).rejects.toThrow(message);
    }
    for (const minimumEditedFidelity of ['R1', 'R2', 'R3'] as const) {
      await expect(
        writeXlsxRoundTrip(snapshot, {
          acknowledgeOpaqueContent: false,
          limits: { maxOperations: 1 },
          minimumEditedFidelity,
          readerLimits: { maxReturnedCells: 10, maxScannedCells: 10 },
        }),
      ).resolves.toMatchObject({ report: { level: 'R0' } });
    }
  });

  it('passes reader limits through strict read and write verification', async () => {
    const bytes = await createIndependentXlsx({
      'xl/worksheets/sheet1.xml': `<worksheet xmlns="${XLSX_SPREADSHEET_NS}">
        <sheetData><row><c r="A1"/><c r="B1"/></row></sheetData>
      </worksheet>`,
    });
    await expect(
      readXlsxRoundTrip(bytes, {
        limits: { maxReturnedCells: 1, maxScannedCells: 2 },
      }),
    ).rejects.toMatchObject({
      diagnostic: { limitName: 'maxReturnedCells' },
      name: 'XlsxParseError',
    });
    const snapshot = await readXlsxRoundTrip(bytes);
    const error = await captureWriteError(() =>
      writeXlsxRoundTrip(snapshot, {
        readerLimits: { maxReturnedCells: 1, maxScannedCells: 2 },
      }),
    );
    expect(error.diagnostic).toMatchObject({
      code: 'source-package-mismatch',
      message: 'XLSX source package failed strict verification',
    });
  });
});
