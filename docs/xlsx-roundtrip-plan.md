# XLSX JSON Round-Trip Plan

> Status: proposed architecture and delivery plan. OAKit does not currently
> read, write, or round-trip XLSX files. This document defines the additional
> preservation and writer contracts required after the XLSX reader described in
> [xlsx-reader-plan.md](xlsx-reader-plan.md) exists.

## Direct answer: what “100%” can mean

The normalized reader plan alone cannot guarantee:

```text
XLSX -> normalized JSON -> XLSX
```

for every workbook and every possible JSON edit. Normalization intentionally
removes XML storage details, relationship IDs, producer extensions, package
ordering, unused styles, calculation caches, and other information a writer may
need to recreate the source package.

A truthful round-trip design needs separate, measurable guarantees:

| Guarantee                                                                           | Can be 100%?                                                         | Required contract                                                                                                                                         |
| ----------------------------------------------------------------------------------- | -------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Unchanged JSON returns the exact source `.xlsx` bytes.                              | Yes                                                                  | Preserve the complete source package as opaque Base64 in the round-trip JSON and return it after hash verification when there are no edits.               |
| Unchanged semantic model is rebuilt into a newly serialized but equivalent package. | Only for a declared support profile                                  | Preserve all unsupported parts or reject; validate every rebuilt part and compare normalized semantics. Byte identity is not promised.                    |
| Supported edits preserve all unaffected source content.                             | Yes, within a versioned operation/profile matrix                     | Copy untouched parts, patch or rebuild only a proven dirty closure, preserve opaque content only when it cannot be affected, and reject ambiguity.        |
| Arbitrary edits to arbitrary future Excel features retain everything.               | No                                                                   | Unknown extensions can reference changed cells, ranges, IDs, signatures, caches, or external services in ways the writer cannot safely infer.             |
| Edited output is byte-for-byte identical to the input.                              | No                                                                   | Any real edit changes at least content bytes, CRCs, ZIP metadata, signatures, or relationships.                                                           |
| Excel-calculated results are always current after an offline edit.                  | No without a compatible calculation engine or caller-provided caches | Preserve formulas, invalidate affected caches, and request recalculation on open. External/volatile formulas may still require Excel or another producer. |

Therefore the target is not “always emit something.” The target is:

> For an unchanged round-trip snapshot, return exactly the original bytes. For
> an edit inside the published capability profile, return a package that passes
> structural, semantic, producer, security, and preservation checks. Otherwise
> reject with a typed explanation before emitting output.

This is a fail-closed 100% guarantee over a declared domain. Refusing an
operation outside that domain is correct; silently dropping or approximating
content is not.

## Round-trip fidelity levels

Every public write result reports one fidelity level. Callers must never infer
fidelity from the absence of an exception alone.

| Level | Name                   | Promise                                                                                                                                                           |
| ----- | ---------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `R0`  | Exact no-op            | After JSON serialization/deserialization and no edits, output SHA-256 and bytes equal the source package exactly.                                                 |
| `R1`  | Part-preserving edit   | Every untouched uncompressed package part is byte-identical; changed parts are valid and listed; ZIP container bytes may differ.                                  |
| `R2`  | Verified semantic edit | Strictly reparsing output produces the requested normalized semantics for every affected supported feature; unaffected supported semantics remain equal.          |
| `R3`  | Producer-verified edit | In addition to R1/R2, the output opens without repair and preserves required behavior in the declared Excel, LibreOffice Calc, and Google Sheets producer matrix. |

`R0` is possible for every accepted source because it returns the verified
opaque source package. `R1` through `R3` are possible only when the source,
operations, affected dependency graph, and security mode meet the versioned
support manifest.

`R0` is a terminal no-op class rather than a lower edited-fidelity score. For
edited workbooks, R1, R2, and R3 are cumulative: R2 includes R1 evidence, and R3
includes R1/R2 evidence. The writer returns R0 for no-op or the highest verified
edited level plus the evidence used. A release must not advertise general
edited round-trip support until the required feature and operation rows reach
R3.

## Architectural principles

### Semantic JSON and preservation data are separate

The existing `XlsxDocument` remains a clean, normalized model. Raw XML,
relationship IDs, ZIP entries, and caches do not leak into it.

Round-trip mode wraps that model in a separate, explicitly opaque envelope:

```text
XlsxRoundTripSnapshot
├── normalized semantic document
├── ordered, typed edit operations
├── source and state integrity metadata
└── opaque source package bytes
```

The preservation package is not an agent prompt or semantic document. Agent
adapters must strip it before sending workbook content to a model. It is
untrusted binary data owned only by the round-trip pipeline.

### Operations, not arbitrary object diffs, drive edits

An arbitrary mutation of normalized JSON does not identify:

- whether a missing field means delete, unsupported, omitted by selection, or
  intentionally unchanged;
- which formulas, names, tables, drawings, pivots, caches, or extensions depend
  on a changed range;
- how stable IDs and relationship ownership should change;
- which package parts must be rebuilt;
- whether the requested transformation is safe for opaque content.

The public writer therefore accepts a versioned discriminated union of edit
operations. Operations are ordered, atomic, validated, bounded, and replayable.
The semantic preview is derived from those operations rather than used as an
untrusted diff source.

### Writer behavior is strict

The tolerant reader can omit malformed optional content with diagnostics. The
writer cannot serialize a partially understood or internally inconsistent
model and hope Excel repairs it.

The writer either:

1. proves the operation is within the declared capability profile;
2. produces and validates the complete dirty closure;
3. returns bytes plus a fidelity report;

or throws `XlsxWriteError` without returning a package.

There is no initial `bestEffort`, `force`, or `ignoreValidation` option. A later
lossy export API, if needed, must be named and documented separately and can
never report round-trip fidelity.

### Reader, editor, and writer stay separate

```text
strict round-trip reader
  -> preservation snapshot
    -> pure edit planner and semantic preview
      -> strict writer and package validator
```

- the reader discovers and normalizes source state;
- the editor validates operations and calculates affected semantic objects;
- the writer owns IDs, relationships, content types, XML serialization, part
  copying/patching/rebuilding, and package validation;
- the validator reparses output independently from writer construction state.

No writer function should be a reverse traversal of tolerant reader code.

## Proposed public API

Names are provisional until black-box contract tests freeze them:

```ts
export async function readXlsxRoundTrip(
  input: XlsxInput,
  options?: XlsxRoundTripReadOptions,
): Promise<XlsxRoundTripSnapshot>;

export async function applyXlsxEdits(
  snapshot: XlsxRoundTripSnapshot,
  operations: readonly XlsxEditOperation[],
): Promise<XlsxRoundTripSnapshot>;

export async function writeXlsxRoundTrip(
  snapshot: XlsxRoundTripSnapshot,
  options?: XlsxWriteOptions,
): Promise<XlsxWriteResult>;

export async function validateXlsxRoundTripJson(
  value: unknown,
): Promise<XlsxRoundTripSnapshot>;

export interface XlsxRoundTripReadOptions {
  securityMode?: 'reject-active' | 'preserve-opaque';
  limits?: XlsxResourceLimits;
}

export interface XlsxWriteOptions {
  minimumEditedFidelity?: 'R1' | 'R2' | 'R3';
  acknowledgeOpaqueContent?: boolean;
  limits?: XlsxWriteLimits;
}
```

All functions are deterministic and do not mutate caller-owned inputs,
snapshots, operations, byte arrays, or option objects.

Round-trip reads are always strict and complete within resource limits:

- tolerant recovery is not accepted because omitted source semantics cannot be
  edited or verified safely;
- structurally valid but unsupported optional content is not treated as a
  tolerant omission: it remains inside the exact opaque source package, is
  inventoried for conflict analysis, and contributes no invented normalized
  semantics; policy may reject it before snapshot creation;
- sheet/range selection is not accepted because the writer needs a complete
  supported dependency graph and semantic base hash;
- Blob/object URL media modes are not accepted because URLs are process-local
  and non-deterministic;
- binary media remains in the opaque source package and semantic media metadata
  remains in the document;
- pivot records and other large optional semantic payloads follow the support
  profile, while their untouched source parts remain opaque and exact;
- if full required parsing exceeds limits, round-trip edit mode rejects instead
  of creating an incomplete snapshot.

A source-free `createXlsx(document)` API would be a separate writer contract.
It can create a valid package from supported semantics but cannot claim
preservation of source information that normalized JSON never contained.

### Round-trip snapshot

```ts
export interface XlsxRoundTripSnapshot {
  format: 'xlsx-roundtrip';
  schemaVersion: string;
  supportProfile: string;
  source: {
    packageBase64: string;
    byteLength: number;
    sha256: string;
    conformance: 'transitional' | 'strict';
  };
  baseDocumentHash: string;
  stateHash: string;
  document: XlsxRoundTripDocument;
  operations: XlsxEditOperation[];
  preservation: {
    securityMode: 'reject-active' | 'preserve-opaque';
    containsOpaqueContent: boolean;
    containsDigitalSignatures: boolean;
  };
}

export interface XlsxRoundTripDocument extends XlsxDocument {
  // Every editable workbook, sheet, table, drawing, chart, pivot, comment,
  // style, and other operation target carries an OAKit-owned snapshot key.
}
```

The envelope is JSON-compatible. Binary data is Base64 because a standalone
JSON file must carry all information required to reproduce an unchanged source
package. This adds approximately one-third encoding overhead before JSON
escaping, so snapshot byte and text limits are mandatory.

`schemaVersion` defines JSON validation and migration. `supportProfile`
identifies the exact feature/operation/producer matrix used for fidelity claims.
The writer rejects unknown future versions rather than interpreting them as an
older shape.

### Snapshot integrity

The opaque package is authoritative source evidence. The writer:

1. Base64-decodes it with strict lexical and size validation;
2. verifies `byteLength` and `sha256`;
3. reparses it through the strict package and XLSX reader;
4. computes canonical normalized `baseDocumentHash`;
5. replays ordered operations;
6. computes the derived semantic document and `stateHash`;
7. compares the derived document with the snapshot preview.

If the semantic preview was manually changed without a matching validated
operation sequence, the writer throws `snapshot-integrity-failed`. Editing the
operation list must go through `validateXlsxRoundTripJson` or
`applyXlsxEdits`, which regenerates the preview and hashes.

The canonical hash algorithm, property ordering, numeric formatting, Unicode
handling, and schema version are part of the public round-trip contract. It
must not depend on engine object-key accidents, locale, or `JSON.stringify`
implementation details.

Round-trip object keys are deterministic within the snapshot and are included
in semantic hashing, but they are not OOXML relationship/style/shared-string
IDs. A fresh parse of the same exact source and support profile must generate
the same keys. Added objects derive keys deterministically from the operation
sequence and reject collisions.

### Edit operations

Every operation has an ID for diagnostics and an optional precondition for
optimistic concurrency:

```ts
export interface XlsxEditOperationBase {
  operationId: string;
  ifMatch?: string;
}

export type XlsxEditOperation =
  | XlsxSetCellOperation
  | XlsxClearCellOperation
  | XlsxSetCellStyleOperation
  | XlsxSetCommentOperation
  | XlsxSetHyperlinkOperation
  | XlsxSetRowOperation
  | XlsxSetColumnOperation
  | XlsxInsertRowsOperation
  | XlsxDeleteRowsOperation
  | XlsxInsertColumnsOperation
  | XlsxDeleteColumnsOperation
  | XlsxAddWorksheetOperation
  | XlsxRenameWorksheetOperation
  | XlsxDeleteWorksheetOperation
  | XlsxSetTableOperation
  | XlsxSetValidationOperation
  | XlsxSetConditionalFormatOperation
  | XlsxSetDrawingOperation
  | XlsxSetChartOperation
  | XlsxSetPivotOperation;
```

Operations target a round-trip `sheetKey`, not a sheet name or OOXML
relationship ID. A `sheetKey` is an OAKit-owned stable identifier scoped to the
snapshot. It survives a rename and never exposes package IDs. Cell/range
coordinates are interpreted in the sequential state produced by earlier
operations.

`ifMatch` is a canonical hash of the targeted semantic object before the
operation. A mismatch aborts the entire edit sequence. This prevents an agent,
retry, or concurrent process from applying an operation to stale workbook
state.

## Source layout

Round-trip code remains inside the XLSX domain:

```text
src/formats/xlsx/
├── index.ts                         Reader public API
├── types.ts                         Normalized reader types
├── roundtrip/
│   ├── index.ts                     Round-trip public API
│   ├── types.ts                     Snapshot, operations, reports, limits
│   ├── errors.ts                    XlsxWriteError and write diagnostics
│   ├── read-snapshot.ts             Strict read plus opaque package capture
│   ├── validate-snapshot.ts         Schema, hash, and Base64 validation
│   ├── apply-edits.ts               Pure ordered operation planner
│   ├── impact-graph.ts              Semantic and package dirty closure
│   ├── capability.ts                Versioned feature/operation checks
│   ├── write.ts                     Strict writer orchestration
│   └── internal/
│       ├── package-graph.ts          Internal parts/content types/rels graph
│       ├── part-copy.ts              Byte-preserving unchanged part handling
│       ├── xml-patch.ts              Validated local source-byte patches
│       ├── xml-write.ts              Namespace-aware strict serialization
│       ├── identifiers.ts            ID and relationship allocation
│       ├── workbook-write.ts         Workbook and sheet manifest writes
│       ├── worksheet-write.ts        Cell/row/column serialization
│       ├── formula-rewrite.ts        Token-aware reference transformations
│       ├── style-write.ts            Style identity and append-only allocation
│       ├── shared-string-write.ts    Shared/inline string strategy
│       ├── relationship-write.ts     Owner-scoped relationship updates
│       ├── content-types-write.ts    Content type updates
│       ├── drawing-write.ts          Drawing anchor and media changes
│       ├── chart-write.ts            Chart updates and cache policy
│       ├── pivot-write.ts            Pivot metadata/cache policy
│       ├── calculation-write.ts      Cache/chain/recalculation policy
│       ├── extension-guard.ts        Unknown/opaque impact detection
│       ├── signature-guard.ts        Digital signature policy
│       └── validate-output.ts        Reopen, graph, and semantic comparison
└── internal/                         Reader-owned modules
```

Only format-neutral ZIP writing, strict XML serialization, OPC graph, hashing,
and binary helpers may move to `common`, and only after a second writer format
proves the same contract. XLSX formulas, ranges, styles, tables, drawings,
pivots, and preservation rules remain XLSX-owned.

## Internal preservation model

The public snapshot stores one opaque source package instead of exposing raw
parts. During each write, the strict writer reconstructs an internal package
graph from those bytes:

```text
PackageGraph
├── canonical part identity
├── original uncompressed bytes and SHA-256
├── content type
├── relationship owner and targets
├── supported semantic owner
├── opaque/unknown feature markers
├── signature coverage
└── dirty/copy/patch/rebuild disposition
```

This graph is parse-scoped and never serialized into the normalized document.
It is discarded after the write result is produced.

### Part dispositions

Every source and generated part has exactly one disposition:

| Disposition | Meaning                                                                       | Fidelity rule                                                                                                     |
| ----------- | ----------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------- |
| `copy`      | Operation cannot affect the part.                                             | Uncompressed bytes and content type remain identical.                                                             |
| `patch`     | A bounded local supported node changes without restructuring opaque siblings. | Apply validated non-overlapping byte/token patches, then reparse and compare expected semantics.                  |
| `rebuild`   | The part's supported structure changes materially.                            | Serialize from strict writer state; allowed only when the entire affected feature/extension closure is supported. |
| `add`       | A new part is required.                                                       | Allocate canonical part name, content type, relationship, and IDs deterministically.                              |
| `remove`    | A supported operation makes the part unreachable.                             | Remove only after proving no remaining relationship or opaque owner depends on it.                                |
| `block`     | Safe preservation cannot be proven.                                           | Throw before creating output.                                                                                     |

Unknown content is never normalized into a fake supported model. It is copied
only when the impact graph proves the requested operations cannot affect its
owner, references, IDs, or ranges.

### Patch safety

Local XML patching can preserve producer-specific markup in an otherwise dirty
part, but it must not be broad string replacement.

- source bytes are decoded and structurally validated first;
- the parser records exact token/value spans for supported patch points;
- replacement text is serialized and escaped according to its XML context;
- patches are sorted, non-overlapping, and applied from the end of the part;
- namespace declarations, encodings, line endings, and untouched bytes remain
  unchanged;
- the complete patched part is decoded, reparsed, schema/semantic checked, and
  bounded again;
- if a safe span is unavailable, the writer rebuilds only when allowed or
  blocks the operation.

Formula and relationship targets are never changed with unvalidated textual
search/replace.

## Dirty-part and dependency graph

An edit is not confined to the XML node visibly changed by the user. The
planner computes a transitive semantic and package closure before writing.

### Example impact rules

| Operation                   | Minimum affected closure                                                                                                                                                                                         |
| --------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Set literal cell value      | Owning worksheet; shared strings when applicable; dependent formula/chart/pivot caches; calculation properties.                                                                                                  |
| Set formula                 | Worksheet; formula groups; dependency/calc chain; cached value policy; workbook calculation flags.                                                                                                               |
| Set cell style              | Worksheet and styles part; conditional formatting remains separate.                                                                                                                                              |
| Add comment                 | Worksheet relationships, comments, authors/persons, optional VML/threaded-comment parts, content types.                                                                                                          |
| Set hyperlink               | Worksheet and possibly worksheet relationships; protocol policy.                                                                                                                                                 |
| Add image                   | Worksheet relationships, drawing part/rels, media part, content types, anchor IDs.                                                                                                                               |
| Edit chart                  | Chart part and relationships; series formula/cache policy; drawing owner.                                                                                                                                        |
| Rename worksheet            | Workbook, scoped names, formulas, charts, validations, conditional formats, print areas/titles, pivots, external-link metadata, and supported extensions containing sheet references.                            |
| Insert/delete row or column | Every affected A1/range/structured reference, merges, tables, filters, validations, conditional formats, drawings, comments, names, charts, pivots, print settings, shared/array formulas, and known extensions. |
| Delete worksheet            | Workbook, relationships, content types, app properties, names, references from other sheets/charts/pivots, owned subgraph, and orphan analysis.                                                                  |

The graph is conservative. If the writer cannot prove that an opaque extension
or signed part is independent of the closure, it blocks the edit. This is how
the writer avoids false 100% claims.

### Stable allocation

- preserve existing part names, relationship IDs, shape IDs, style indexes,
  shared-string indexes, table IDs, pivot IDs, and workbook sheet IDs whenever
  their semantics do not change;
- append new IDs deterministically without renumbering existing content;
- never reuse a deleted ID in the same operation sequence unless the relevant
  OOXML contract explicitly permits and tests it;
- validate uniqueness, lexical form, range, and owner scope before writing;
- retain authored order unless an operation explicitly changes it;
- remove orphan parts only when the complete relationship graph proves they are
  unreferenced and not covered by opaque/signature rules.

## Writer pipeline

### 1. Validate JSON and resource budgets

- validate the snapshot schema before decoding large strings;
- enforce JSON depth, object/array counts, string bytes, operation count,
  operation payload, Base64 lexical form, decoded source bytes, media bytes,
  formulas, styles, ranges, and expected output limits;
- reject unknown fields in security-sensitive discriminated unions;
- reject duplicate operation IDs, sheet keys, object IDs, and ambiguous keys;
- validate every number in lexical/finite/safe/range order;
- never execute getters or accept class instances as trusted plain JSON.

### 2. Verify and strictly reparse the source

- verify source length and SHA-256;
- open the OPC archive with actual expansion accounting;
- validate canonical part names, content types, relationships, XML, limits, and
  the declared XLSX main part;
- discover active, opaque, signed, external, and unsupported content;
- strictly reparse the normalized base document;
- verify canonical `baseDocumentHash` and support profile.

No edits are applied until source integrity and capability checks pass.

### 3. Replay operations atomically

- validate operation discriminants and preconditions;
- resolve stable OAKit object keys;
- apply operations sequentially to immutable semantic state;
- validate workbook invariants after each operation and at the final state;
- calculate changed semantic objects and dependency edges;
- derive and verify the snapshot `document` and `stateHash`;
- abort the entire sequence on the first failure without mutating the source
  snapshot.

### 4. Classify fidelity and conflicts

If there are no operations and the derived state equals the source, select
`R0` and return the exact decoded source bytes after verification.

For edited state:

- match every feature and operation against the support profile;
- calculate copy/patch/rebuild/add/remove/block dispositions;
- expand dirty closure through relationships and semantic references;
- check opaque extensions, active content, signatures, external links, caches,
  producer constraints, and writer limits;
- block before serialization if the requested fidelity cannot be proven.

### 5. Plan formulas and recalculation

- tokenize formulas without evaluating them;
- rewrite references only for supported structural operations;
- preserve untouched formula text and cached values byte-for-byte when their
  precedents are unaffected;
- invalidate affected cached results, calculation chains, chart caches, and
  pivot caches according to explicit policies;
- set workbook recalculation flags when producer recalculation is required;
- never invent a formula result;
- require a caller-provided cached result or block when the target profile
  requires an immediately current cache that OAKit cannot calculate.

### 6. Allocate IDs, strings, styles, and parts

- preserve and append rather than globally renumber;
- deduplicate only when semantic identity rules permit it;
- distinguish named styles, differential styles, and visually equal but
  separately named records;
- preserve the source shared-string/inline-string strategy when practical;
- update counts and unique counts from actual writer state;
- allocate new part names and relationships from the owning part;
- prevent collisions after OPC canonicalization.

### 7. Copy, patch, rebuild, add, and remove

- copy untouched part bytes;
- apply verified local patches;
- strictly serialize fully supported dirty parts;
- add new parts with explicit content types and owner relationships;
- remove parts only after orphan and opaque/signature analysis;
- preserve source Strict/Transitional conformance for dirty parts unless a
  separately named conversion operation is requested;
- write a bounded deterministic ZIP container.

### 8. Validate the complete output graph

Before semantic comparison:

- reopen the generated archive from bytes;
- repeat ZIP, OPC, relationship, XML, numeric, and resource-limit checks;
- require one valid main workbook part;
- verify every internal relationship target exists;
- verify every reachable part has the correct content type;
- reject duplicate canonical names, duplicate relationship IDs, dangling
  targets, invalid roots, and forbidden active content for the selected mode;
- verify no unintended parts were added, removed, or changed;
- compare hashes of all parts classified `copy`.

### 9. Verify semantic intent

Strictly parse the output through a fresh round-trip reader instance, including
its opaque inventory, and compare:

- every edited semantic object with the operation-derived target;
- every unaffected supported semantic object with the source;
- workbook/sheet order, stable object identity, formulas, cached-state markers,
  styles, tables, drawings, charts, pivots, comments, external metadata, and
  diagnostics required by the support profile;
- no `NaN`, `Infinity`, raw parser state, invalid URL, or unsafe content escape;
- deterministic output across repeated and concurrent writes.

Writer construction objects are not accepted as validation evidence. Reopen
and comparison use independent state.

### 10. Return bytes and fidelity evidence

```ts
export interface XlsxWriteResult {
  bytes: Uint8Array;
  report: {
    level: 'R0' | 'R1' | 'R2' | 'R3';
    sourceSha256: string;
    outputSha256: string;
    copiedParts: string[];
    patchedParts: string[];
    rebuiltParts: string[];
    addedParts: string[];
    removedParts: string[];
    invalidatedCaches: string[];
    recalculationRequired: boolean;
    producerEvidence: string[];
  };
}
```

Part names in a fidelity report are bounded provenance strings, not access to
raw XML. Reports never contain cell text, formulas, credentials, connection
strings, or opaque payload bytes.

## Security and preservation modes

Security and perfect preservation can conflict. Sanitizing content changes the
file; preserving opaque active content can carry risk forward. The API must make
that trade-off explicit.

### `reject-active` default

- reject macro-enabled, binary, encrypted, executable, ActiveX, OLE, script,
  and other active package content;
- never preserve a source part that violates the declared safe XLSX profile;
- preserve safe external-link metadata without fetching it;
- redact secrets from diagnostics and reports;
- provide the strongest agent-safe default, but refuse rather than sanitize a
  source that cannot be preserved under this policy.

### `preserve-opaque` opt-in

- copy allowed opaque parts byte-for-byte without decoding, executing, or
  exposing them semantically;
- mark the snapshot and report as containing opaque content;
- block any edit whose impact closure may affect an opaque owner or reference;
- never claim the output was sanitized;
- require explicit caller acknowledgement on write;
- remain unavailable to agent adapters unless the host has an independent
  isolation and content policy.

This mode still rejects executable main workbook formats such as `.xlsm` or
`.xlsb` until those formats have separate public and security contracts.

### Digital signatures

- `R0` exact no-op preserves signatures because it returns the original bytes;
- any edited package can invalidate an OPC signature even if most parts are
  copied;
- the writer cannot re-sign without caller identity, credentials, algorithms,
  certificate policy, and a separate signing API;
- the strict default blocks edits to signed packages;
- an explicit future `remove-signatures` conversion can remove signatures but
  is lossy and must never report signature preservation or R1.

### External data and formulas

- never fetch external workbooks, web functions, linked images, connections,
  queries, linked data types, DDE, RTD, or pivot sources;
- preserve unaffected safe metadata under the support profile;
- block edits requiring refresh or unknown external resolution;
- omit credentials and connection strings from semantic JSON and reports;
- treat formulas, names, links, comments, and all cell content as untrusted
  data rather than instructions.

## Formula and cache fidelity

Formulas are the largest semantic gap between editing XML and producing a
workbook whose displayed values are immediately current.

### Formula source of truth

- formula expression is authoritative;
- cached value is producer-generated evidence, not the formula definition;
- workbook calculation properties describe recalculation intent;
- calculation chain is an optimization and may be stale;
- chart/pivot series caches are derived content and may also be stale.

### Edit policy

| Situation                                                 | Writer behavior                                                                                   |
| --------------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| Formula and all precedents unchanged                      | Preserve formula bytes and cached result when possible.                                           |
| Literal precedent changed                                 | Preserve formula expression, invalidate affected cache/chain, set recalculation flags.            |
| Formula changed with caller-provided compatible cache     | Write both, mark cache provenance as caller-provided, validate type.                              |
| Formula changed without cache                             | Omit invalid cache and require recalculation on open.                                             |
| Structural edit with fully supported reference rewrite    | Token-aware rewrite formulas/names/ranges, invalidate affected caches.                            |
| Structural edit reaches unsupported/opaque formula syntax | Block the operation.                                                                              |
| External/volatile formula needs a current result          | Preserve expression, require caller cache or report recalculation required; never fetch/evaluate. |

`R2` verifies formula expressions and cache-state policy. `R3` additionally
requires producer evidence that recalculation/open behavior matches the support
profile. OAKit must not claim that an invalidated cache is a current result.

## Unknown extensions and preservation conflicts

Unknown valid extensions are acceptable for `R0` because exact source bytes are
returned. They may also be copied in an edited package only when independence
is proven.

Conservative conflict rules include:

- block a structural edit to a worksheet containing an unknown extension that
  may encode cells, ranges, IDs, drawings, controls, or formula dependencies;
- block renaming/deleting a sheet when an opaque part may contain sheet names or
  sheet IDs;
- block relationship/ID renumbering in any owner with opaque children or
  opaque related parts;
- block pivot/table/chart source edits when an unsupported cache or extension
  depends on the changed range;
- copy a completely unrelated opaque custom property or image part only when
  the package graph proves no dirty owner/reference connection;
- report the exact blocking owner, operation ID, feature class, and reason
  without dumping raw XML.

An `unsafePreserveUnknown` escape hatch is intentionally excluded from the
round-trip API. It would turn an unproven result into a misleading fidelity
claim.

## Writer feature and operation matrix

Round-trip support is the intersection of reader feature coverage and edit
operation coverage. A reader feature marked supported is not automatically
writable.

| Domain                            | No-op preservation                  | Initial editable operations                                                        | Full target evidence                                                     |
| --------------------------------- | ----------------------------------- | ---------------------------------------------------------------------------------- | ------------------------------------------------------------------------ |
| Workbook/sheets                   | Exact through R0                    | Add/rename/delete/reorder supported sheets                                         | Reference rewrite, graph validity, producer open without repair          |
| Literal cells                     | Exact through R0                    | Set/clear text, number, boolean, error, ISO date, blank                            | Types, styles, shared/inline strings, exact boundaries                   |
| Formulas                          | Exact through R0                    | Set formula/cache policy; supported reference rewrites                             | Token/property tests, recalculation evidence, external/volatile policy   |
| Styles/number formats             | Exact through R0                    | Apply existing or append normalized style                                          | Identity/precedence, locale/time-zone determinism, no global renumbering |
| Rows/columns/merges               | Exact through R0                    | Size/visibility, then insert/delete and merge operations                           | Complete reference transformation and opaque conflict checks             |
| Names/tables/filters              | Exact through R0                    | Set/remove supported definitions and rules                                         | Scope, formulas, totals, structured references, producer corpus          |
| Validation/conditional formatting | Exact through R0                    | Set/remove normalized rules                                                        | Priority/order, differential styles, extension forms                     |
| Comments/hyperlinks               | Exact through R0                    | Set/remove safe normalized content                                                 | Authors/persons/VML/rels, protocol and redaction tests                   |
| Drawings/images                   | Exact through R0                    | Add/remove/update supported anchors and embedded images                            | IDs/rels/media/content types, browser lifecycle, producer render         |
| Charts                            | Exact through R0                    | Update supported chart semantics and series references                             | ChartML validation, cache policy, render/producer corpus                 |
| Sparklines                        | Exact through R0                    | Update supported groups/ranges                                                     | Extension namespaces and structural rewrite tests                        |
| Pivots/slicers/timelines          | Exact through R0                    | Metadata/style edits first; source/range edits only after cache policy is complete | Cache graph, refresh flags, Excel/Calc evidence                          |
| Print/protection/views            | Exact through R0                    | Set supported normalized properties                                                | Units, algorithms, page/view producer checks                             |
| External links/connections        | Exact safe/opaque copy under policy | Metadata edits only after redaction and no-refresh rules                           | No network, secret handling, relationship graph                          |
| Unknown extensions                | Exact R0                            | No edit unless proven independent and copied                                       | Versioned support manifest or typed block                                |
| Digital signatures                | Exact R0                            | None without separate signing/removal contract                                     | Cryptographic validation and credential policy                           |

The support manifest records each operation/domain pair as:

```text
unsupported | preservation-only | editable-R1 | verified-R2 | producer-R3
```

Marketing, README, and API documentation may claim only the lowest level that
all advertised operation/domain pairs have reached.

## Resource limits

Writer input is untrusted JSON plus an embedded untrusted Office package. It
needs all reader limits and additional writer limits:

```ts
export interface XlsxWriteLimits {
  maxSnapshotJsonBytes?: number;
  maxSnapshotDepth?: number;
  maxSnapshotObjects?: number;
  maxSourcePackageBytes?: number;
  maxOperations?: number;
  maxOperationBytes?: number;
  maxTotalOperationBytes?: number;
  maxDirtyParts?: number;
  maxPatchedParts?: number;
  maxPatchCount?: number;
  maxPatchBytes?: number;
  maxGeneratedXmlBytes?: number;
  maxGeneratedMediaBytes?: number;
  maxOutputBytes?: number;
  maxFormulaRewriteTokens?: number;
  maxReferenceUpdates?: number;
  maxDependencyEdges?: number;
  maxValidationPasses?: number;
}
```

The writer also reapplies archive, XML, worksheet, cell, text, formula, style,
media, chart, pivot, and relationship limits to source and output.

Every limit needs invalid, below, at, and one-over tests. Limit violations are
fatal. There is no partial output. Accounting includes copied, patched,
generated, and validation passes so an attacker cannot hide work in unchanged
source bytes or repeated reopen cycles.

## Errors and diagnostics

```ts
export type XlsxWriteDiagnosticCode =
  | 'invalid-roundtrip-json'
  | 'unsupported-snapshot-version'
  | 'snapshot-integrity-failed'
  | 'source-package-mismatch'
  | 'operation-precondition-failed'
  | 'unsupported-edit-operation'
  | 'preservation-conflict'
  | 'opaque-content-conflict'
  | 'signed-package-conflict'
  | 'formula-rewrite-unsupported'
  | 'recalculation-required'
  | 'identifier-allocation-failed'
  | 'relationship-graph-invalid'
  | 'generated-package-invalid'
  | 'semantic-verification-failed'
  | 'producer-verification-failed'
  | 'resource-limit-exceeded';
```

`XlsxWriteError` carries a typed diagnostic with bounded fields such as
operation ID, public object key, sheet/cell/range, owner part, fidelity level,
feature class, and limit metadata. It never includes source XML, opaque bytes,
credentials, full formulas, or unbounded cell text.

`recalculation-required` is a report warning for a successfully written
supported operation when the profile permits producer recalculation. It is an
error only when the requested fidelity profile requires current cached values.

## Independent test strategy

### No-op exactness

For every accepted corpus workbook:

```text
source bytes
  -> readXlsxRoundTrip
  -> JSON.stringify
  -> JSON.parse
  -> validateXlsxRoundTripJson
  -> writeXlsxRoundTrip(no operations)
  -> exact byte equality and SHA-256 equality
```

Test empty, small, large, Strict, Transitional, Excel, LibreOffice Calc, Google
Sheets, unknown-extension, external-link, opaque-policy, and signed-package
fixtures. R0 must not depend on ZIP reserialization.

### Part preservation after edits

- independently unzip source and output;
- compare complete canonical part-name sets;
- require exact bytes for every `copy` part;
- verify report dispositions match actual part hashes;
- verify no undeclared part changes, additions, or removals;
- verify content types and relationships independently from writer helpers;
- compare source/output opaque and extension parts expected to remain copied.

### Semantic edit assertions

Every operation fixture defines literal:

- source semantic state;
- operation and optional precondition;
- expected dirty closure;
- expected copied/changed part hashes;
- expected output semantic state;
- formula/cache/recalculation behavior;
- expected fidelity level or typed blocking diagnostic.

Tests call public APIs and strictly reparse output. Expected formulas, ranges,
dates, styles, and graph changes are not calculated with writer helpers.

### Producer verification

R3 requires a versioned producer matrix:

- Microsoft Excel for Windows;
- Microsoft Excel for macOS;
- LibreOffice Calc;
- Google Sheets import/export where the feature is representable.

Evidence includes:

- opens without repair/recovery prompts;
- no producer repair log;
- ordered sheet names and visibility;
- representative values, formulas, calculated results after permitted
  recalculation, styles, tables, comments, charts, pivots, and validations;
- visual/render checks for features whose contract is visual;
- save/reopen semantic checks where the producer rewrites valid package syntax.

Open XML SDK/package validation and LibreOffice headless checks can run in CI.
Excel application validation may require a controlled licensed Windows/macOS
environment; it is a release gate for R3 claims, not something silently
replaced by another library.

### Metamorphic properties

- no-op JSON round trip preserves exact bytes;
- applying an empty operation list equals no-op;
- applying one operation in isolation equals the corresponding one-operation
  batch;
- sequential operation order is deterministic;
- non-overlapping commutative operations produce equal semantic state even if
  package serialization differs;
- `parseStrict(write(apply(read(source), ops)))` equals the derived target for
  supported semantics;
- applying an operation and its declared inverse restores semantic state, while
  byte identity is required only when the writer can return the untouched
  source snapshot;
- repeated and concurrent writes do not share IDs, caches, dirty state, or
  object URLs.

### Adversarial and property tests

Fuzz and preserve minimized cases for:

- malformed/oversized snapshot JSON and deeply nested operations;
- invalid Base64, source hashes, state hashes, Unicode, and duplicate keys;
- ZIP/OPC/XML attacks embedded in the preservation package;
- operation sequences at grid, ID, range, formula, and size boundaries;
- stale `ifMatch` preconditions and duplicate operation IDs;
- formula token rewriting across absolute, mixed, named, structured, 3D,
  external, array, dynamic-array, and string-literal boundaries;
- insert/delete transforms through merges, tables, validations, conditional
  formatting, comments, drawings, charts, pivots, and print ranges;
- relationship and content-type allocation collisions;
- unknown extension, signature, opaque, external, and active-content conflicts;
- patch span overlap, encoding/chunk boundaries, escaping, and reparsing;
- output-size and validation-pass accounting;
- writer failure after temporary media/object URL creation and cleanup.

### Mutation testing

Mutation scope must include:

- schema and hash validation;
- preconditions and operation ordering;
- impact/dirty closure edges;
- copy/patch/rebuild/block classification;
- ID and relationship allocation;
- formula/reference transformations;
- cache invalidation and recalculation flags;
- opaque/signature/security guards;
- part hash comparison;
- output graph and semantic verification;
- resource accounting and cleanup.

No writer target may introduce `Survived` or `NoCoverage` mutants. Tests kill
mutants through public output, exact part hashes, semantic results, or typed
blocking diagnostics rather than private implementation shape.

## Performance and browser requirements

- exact no-op should decode/verify source bytes without rebuilding the ZIP;
- edit memory is bounded by source limits, dirty parts, dependency state, and
  output bytes, not the theoretical workbook grid;
- copied parts should not be parsed into semantic trees unless validation or
  impact analysis requires it;
- patch/rebuild passes are bounded and counted;
- browser code contains no filesystem, process, Buffer-only, or Node stream
  dependencies;
- hashing, Base64, ZIP writing, XML writing, and validation behave identically
  in Node and Chromium;
- browser object URLs and temporary buffers are released on success/failure;
- the `./xlsx/roundtrip` subpath has its own reviewed bundle-size and startup
  baseline;
- public upload hosts still enforce outer worker/process time and memory limits.

Scale suites cover large source Base64, many copied parts, dense local edits,
wide dependency closures, many formulas/styles/relationships, media, and
one-over failures that prove early abort.

## Package and CLI surface

Only after R0/R1/R2 contracts pass:

- expose `@evoelsewhere/oakit/xlsx/roundtrip` separately from the lightweight
  reader subpath;
- add ESM, CommonJS, declarations, browser, and tree-shaking smoke tests;
- keep round-trip code out of reader-only bundles;
- add `oakit xlsx snapshot input.xlsx --output workbook.json`;
- add `oakit xlsx apply workbook.json operations.json --output edited.json`;
- add `oakit xlsx write edited.json --output output.xlsx`;
- make CLI output refuse overwrites, validate hashes, apply conservative
  limits, and emit a machine-readable fidelity report;
- never print opaque package Base64, cell content, formulas, or credentials to
  stderr diagnostics;
- update README round-trip claims only to the achieved support profile/level.

## Delivery plan and atomic commits

Each item below is a delivery boundary. Regression tests, implementation,
infrastructure, and public documentation should remain separate commits when
independently meaningful.

### 1. Freeze fidelity terminology

- add this plan;
- define R0/R1/R2/R3 and “verified output or fail”;
- define the first support-profile schema;
- define snapshot and write diagnostic contracts.

Gate: documentation formatting and public contract review.

### 2. Add independent snapshot fixtures

- build literal round-trip JSON fixtures independent of production code;
- cover valid/invalid Base64, hashes, schema versions, and operations;
- add corpus metadata for exact source SHA-256 expectations.

Gate: fixture self-checks and JSON boundary tests.

### 3. Implement strict snapshot validation

- schema/depth/count/string limits;
- Base64 and source SHA-256 validation;
- canonical semantic hashing;
- immutable input tests;
- Node/browser hash equivalence.

Gate: focused tests, browser tests, mutation targets.

### 4. Implement R0 exact no-op

- strict source parse and base hash verification;
- JSON serialize/parse cycle;
- exact original-byte return when operations are empty;
- Excel/Calc/Sheets corpus SHA equality.

Gate: complete R0 corpus and concurrency tests.

### 5. Build internal package graph

- canonical parts, content types, relationships, hashes, owners;
- opaque, active, extension, signature, and external markers;
- independent graph validity tests;
- conservative conflict classification.

Gate: OPC/XML/security mutation suite and existing reader regression suite.

### 6. Implement operation planner

- stable sheet/object keys;
- ordered atomic operations and `ifMatch`;
- semantic preview and state hash;
- impact graph with explicit block outcomes;
- operation/resource limits.

Gate: public operation tests, property ordering, mutation targets.

### 7. Implement copy and local XML patching

- exact uncompressed copy parts;
- token/span-safe patches;
- patch encoding/escaping/reparse validation;
- part disposition fidelity reports.

Gate: part hash assertions, fuzzed patch boundaries, no undeclared changes.

### 8. Implement basic cell edits

- set/clear literal values and explicit blank cells;
- preserve/apply existing styles;
- shared/inline strings;
- worksheet patch/rebuild policy;
- calculation invalidation for changed precedents.

Gate: R1/R2 cell matrix, Node/browser, Excel/Calc/Sheets core corpus.

### 9. Implement formulas, styles, comments, and hyperlinks

- formula/cache contract and token-aware rewrites;
- append-only styles/number formats;
- comments/authors/persons/VML;
- safe hyperlinks and relationships.

Gate: formula/property/mutation tests, producer recalculation evidence.

### 10. Implement worksheet rules and ranges

- rows/columns/merges;
- names, tables, filters;
- validations and conditional formats;
- print/view/protection metadata;
- full range impact closure.

Gate: structural reference matrix, producer corpus, opaque conflict tests.

### 11. Implement structural sheet operations

- insert/delete rows and columns;
- add/rename/delete/reorder sheets;
- workbook-wide formula/name/chart/pivot/range transformations;
- owned-part add/remove/orphan handling.

Gate: graph-wide public tests, fuzz, mutation, producer repair-log checks.

### 12. Implement drawings, media, and charts

- anchors, IDs, relationships, content types, media lifecycle;
- supported chart edits and formula/cache policy;
- visual producer verification;
- browser cleanup and bundle gates.

Gate: R1/R2 part and semantic tests plus R3 render/open evidence.

### 13. Implement pivots and advanced extensions

- pivot/cache/slicer/timeline impact policy;
- safe metadata edits;
- supported modern extension writers;
- external/query preservation and strict no-refresh behavior.

Gate: advanced Excel/Calc corpus, cache/extension conflicts, security tests.

### 14. Integrate package, CLI, and CI

- `./xlsx/roundtrip` exports and package smoke tests;
- snapshot/apply/write CLI commands;
- Node 20/22/24 and Chromium;
- corpus, scale, producer, mutation, and fidelity report artifacts;
- README and architecture capability truth.

Gate: all checks and versioned support manifest review.

### Recommended commit granularity

The implementation should naturally produce more than thirty small commits
without artificial splitting. A representative sequence is:

```text
docs fidelity contract
test snapshot schema
feat snapshot validator
test canonical hashes
feat canonical hashes
test exact no-op
feat exact no-op
test package graph
feat content-type graph
feat relationship graph
test opaque conflicts
feat capability guard
test operation preconditions
feat operation planner
test dirty closure
feat dirty closure
test part copying
feat part copying
test XML patch spans
feat XML patching
test literal cell writes
feat literal cell writes
test shared strings
feat shared strings
test style allocation
feat style allocation
test formula cache policy
feat formula cache policy
test comments and links
feat comments and links
test structural references
feat structural references
test drawings and charts
feat drawings and charts
test output validation
feat output validation
test browser round-trip
test producer corpus
test mutation coverage
build package exports
feat CLI round-trip
ci fidelity gates
docs public capability
```

Commit subjects use normal intent (`test:`, `feat:`, `fix:`, `refactor:`,
`build:`, `ci:`, `docs:`), not phase labels or score names.

## Definition of done

### R0 exact no-op is complete when

- snapshot JSON is versioned, bounded, and independently validated;
- source bytes, length, and SHA-256 are verified;
- normalized base and state hashes are canonical across Node and browser;
- JSON stringify/parse does not change the snapshot;
- no-operation output equals source bytes exactly for the complete accepted
  Excel/Calc/Sheets corpus, including supported unknown/opaque/signature cases;
- invalid/tampered snapshots fail with typed diagnostics and no output;
- mutation, fuzz, browser, concurrency, and resource gates pass.

### A supported editable domain is complete at R1/R2 when

- its reader contract is already complete;
- every advertised operation has valid, invalid, missing, boundary,
  precondition, conflict, inverse, and concurrency tests;
- the impact graph includes every supported semantic/package dependency;
- every untouched part is byte-identical and every changed part is declared;
- unknown/opaque/signature conflicts block rather than degrade;
- output package graph passes independent strict validation;
- strict reparse equals requested supported semantics and preserves unaffected
  supported semantics;
- formulas/caches/recalculation behavior is explicit;
- exact at/one-over writer limits pass;
- no new survived or uncovered mutants exist.

### An editable domain is complete at R3 when

- R1 and R2 are complete;
- the versioned Excel Windows/macOS, LibreOffice Calc, and applicable Google
  Sheets matrix opens output without repair;
- formulas recalculate according to the declared policy;
- visual and interactive supported features match semantic expectations;
- producer save/reopen does not reveal invalid or missing required structures;
- producer/version evidence is retained as a CI/release artifact.

### General round-trip may be advertised only when

- every advertised reader feature has an explicit preservation/edit level in
  the support manifest;
- documentation distinguishes R0 exact no-op from supported edited fidelity;
- unsupported edits fail before output and have stable diagnostics;
- there is no best-effort path labeled round-trip;
- README, architecture, package exports, declarations, CLI, examples, and
  producer evidence agree;
- existing PPTX and XLSX reader gates remain green.

## Remaining impossibility boundaries

Even after the full plan is implemented, OAKit must not promise:

- byte identity after a real edit;
- preservation of a digital signature after an edit without re-signing;
- current results for external, volatile, proprietary, or unsupported formulas
  without producer recalculation or caller-provided caches;
- safe editing of unknown extensions whose dependencies cannot be proven;
- `.xls`, `.xlsb`, `.xlsm`, encrypted, or rights-managed document support;
- identical visual rendering across producer versions for behavior outside the
  versioned profile;
- recovery of information that the source producer itself did not store.

These boundaries are not unfinished test coverage. They are properties of the
format, cryptography, external services, and the difference between semantic
editing and exact binary preservation.

## Normative references

Use the reader plan's ECMA-376, `MS-XLSX`, and `MS-OE376` references for package,
SpreadsheetML, Markup Compatibility, and extension behavior. Writer-specific
decisions must also check:

- ECMA-376 Part 2 for Open Packaging Conventions, relationships, content types,
  digital signatures, and package conformance:
  <https://ecma-international.org/publications-and-standards/standards/ecma-376/>
- Microsoft Office cryptography structures (`MS-OFFCRYPTO`) when detecting and
  explicitly rejecting encrypted or rights-managed inputs:
  <https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/>

Producer behavior can add a documented compatibility path only when the output
remains package-valid, safe, independently tested, and inside the published
support profile.
