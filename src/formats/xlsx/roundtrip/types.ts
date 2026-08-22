import type {
  XlsxCellValue,
  XlsxDocument,
  XlsxHyperlinkTarget,
  XlsxInput,
  XlsxResourceLimits,
  XlsxSheet,
  XlsxStyle,
} from '../types';

export type XlsxFidelityLevel = 'R0' | 'R1' | 'R2' | 'R3';
export type XlsxCapabilityLevel =
  | 'editable-R1'
  | 'preservation-only'
  | 'producer-R3'
  | 'unsupported'
  | 'verified-R2';

export interface XlsxCapabilityEntry {
  domain: string;
  level: XlsxCapabilityLevel;
}

export interface XlsxOperationCapability {
  constraints?: string[];
  level: XlsxCapabilityLevel;
  operation: XlsxEditOperation['kind'];
}

export interface XlsxProducerEvidence {
  producer: 'excel-macos' | 'excel-windows' | 'google-sheets' | 'libreoffice';
  version: string;
}

export interface XlsxCapabilityManifest {
  domains: XlsxCapabilityEntry[];
  effectiveLevel: XlsxFidelityLevel;
  id: 'xlsx-agent-ready';
  operations: XlsxOperationCapability[];
  producerEvidence: XlsxProducerEvidence[];
  version: '1';
}

export interface XlsxEditOperationBase {
  ifMatch?: string;
  operationId: string;
  sheetKey: string;
}

export type XlsxEditOperation =
  | (XlsxEditOperationBase & {
      cell: string;
      kind: 'clear-cell';
    })
  | (XlsxEditOperationBase & {
      cell: string;
      content:
        | { kind: 'formula'; expression: string }
        | { kind: 'value'; value: XlsxCellValue };
      kind: 'set-cell';
    })
  | (XlsxEditOperationBase & {
      cell: string;
      kind: 'set-cell-style';
      style: XlsxStyle;
    })
  | (XlsxEditOperationBase & {
      end: number;
      kind: 'set-column';
      start: number;
      width?: number;
      hidden?: boolean;
    })
  | (XlsxEditOperationBase & {
      cell: string;
      kind: 'set-hyperlink';
      target: XlsxHyperlinkTarget | null;
    })
  | (XlsxEditOperationBase & {
      height?: number;
      hidden?: boolean;
      kind: 'set-row';
      row: number;
    })
  | (XlsxEditOperationBase & {
      count: number;
      index: number;
      kind: 'delete-columns' | 'delete-rows' | 'insert-columns' | 'insert-rows';
    })
  | (XlsxEditOperationBase & {
      kind: 'rename-worksheet';
      name: string;
    })
  | (Omit<XlsxEditOperationBase, 'sheetKey'> & {
      kind: 'add-worksheet';
      name: string;
    })
  | (XlsxEditOperationBase & { kind: 'delete-worksheet' });

export type XlsxRoundTripSheet = XlsxSheet & { key: string };

export interface XlsxRoundTripDocument extends Omit<XlsxDocument, 'sheets'> {
  key: string;
  sheets: XlsxRoundTripSheet[];
}

export interface XlsxRoundTripSource {
  byteLength: number;
  conformance: 'strict' | 'transitional';
  packageBase64: string;
  sha256: string;
}

export interface XlsxRoundTripPreservation {
  containsActiveContent: boolean;
  containsDigitalSignatures: boolean;
  containsExternalRelationships: boolean;
  containsOpaqueContent: boolean;
  securityMode: 'preserve-opaque' | 'reject-active';
}

export interface XlsxRoundTripSnapshot {
  baseDocumentHash: string;
  document: XlsxRoundTripDocument;
  format: 'xlsx-roundtrip';
  keyAlgorithmVersion: 'xlsx-snapshot-key-v1';
  operations: XlsxEditOperation[];
  preservation: XlsxRoundTripPreservation;
  schemaVersion: '1';
  source: XlsxRoundTripSource;
  sourceManifestHash: string;
  stateHash: string;
  supportProfile: XlsxCapabilityManifest;
}

export interface XlsxRoundTripReadOptions {
  limits?: XlsxResourceLimits;
  securityMode?: 'preserve-opaque' | 'reject-active';
}

export interface XlsxWriteLimits {
  maxDependencyEdges?: number;
  maxDirtyParts?: number;
  maxFormulaRewriteTokens?: number;
  maxGeneratedMediaBytes?: number;
  maxGeneratedXmlBytes?: number;
  maxOperationBytes?: number;
  maxOperations?: number;
  maxOutputBytes?: number;
  maxPatchBytes?: number;
  maxPatchCount?: number;
  maxPatchedParts?: number;
  maxReferenceUpdates?: number;
  maxSnapshotDepth?: number;
  maxSnapshotJsonBytes?: number;
  maxSnapshotObjects?: number;
  maxSourcePackageBytes?: number;
  maxTotalOperationBytes?: number;
  maxValidationPasses?: number;
}

export interface XlsxWriteOptions {
  acknowledgeOpaqueContent?: boolean;
  limits?: XlsxWriteLimits;
  minimumEditedFidelity?: 'R1' | 'R2' | 'R3';
  readerLimits?: XlsxResourceLimits;
}

export interface XlsxPartFidelity {
  byteLength: number;
  disposition: 'add' | 'copy' | 'patch';
  name: string;
  sha256: string;
  sourceByteLength?: number;
  sourceSha256?: string;
}

export interface XlsxWriteReport {
  diagnostics: XlsxWriteDiagnostic[];
  level: 'R0' | 'R2';
  outputSha256: string;
  parts: XlsxPartFidelity[];
  sourceSha256: string;
  supportProfile: XlsxCapabilityManifest;
}

export interface XlsxWriteResult {
  data: Uint8Array;
  report: XlsxWriteReport;
}

export type XlsxWriteDiagnosticCode =
  | 'formula-rewrite-unsupported'
  | 'generated-package-invalid'
  | 'identifier-allocation-failed'
  | 'invalid-roundtrip-json'
  | 'opaque-content-conflict'
  | 'operation-precondition-failed'
  | 'preservation-conflict'
  | 'producer-verification-failed'
  | 'recalculation-required'
  | 'relationship-graph-invalid'
  | 'resource-limit-exceeded'
  | 'semantic-verification-failed'
  | 'signed-package-conflict'
  | 'snapshot-integrity-failed'
  | 'source-package-mismatch'
  | 'unsupported-edit-operation'
  | 'unsupported-snapshot-version';

export interface XlsxWriteDiagnostic {
  actual?: number;
  cell?: string;
  code: XlsxWriteDiagnosticCode;
  featureClass?: string;
  fidelity?: XlsxFidelityLevel;
  limit?: number;
  limitName?: keyof XlsxWriteLimits | keyof XlsxResourceLimits;
  message: string;
  objectKey?: string;
  operationId?: string;
  part?: string;
  range?: string;
  severity: 'error' | 'warning';
  sheetKey?: string;
}

export type XlsxRoundTripInput = XlsxInput;

export interface ResolvedXlsxWriteLimits {
  maxDependencyEdges: number;
  maxDirtyParts: number;
  maxFormulaRewriteTokens: number;
  maxGeneratedMediaBytes: number;
  maxGeneratedXmlBytes: number;
  maxOperationBytes: number;
  maxOperations: number;
  maxOutputBytes: number;
  maxPatchBytes: number;
  maxPatchCount: number;
  maxPatchedParts: number;
  maxReferenceUpdates: number;
  maxSnapshotDepth: number;
  maxSnapshotJsonBytes: number;
  maxSnapshotObjects: number;
  maxSourcePackageBytes: number;
  maxTotalOperationBytes: number;
  maxValidationPasses: number;
}
