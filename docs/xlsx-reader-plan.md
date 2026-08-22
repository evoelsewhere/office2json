# XLSX Reader Implementation Plan

> Status: proposed architecture and delivery plan. This document does not
> describe an implemented XLSX capability. OAKit must continue to report XLSX
> as planned until the public API, implementation, fixtures, and required gates
> described here exist.

## Goal

Build a browser-neutral, deterministic, bounded XLSX reader with the same
engineering standard as the PPTX reader:

```text
.xlsx package -> normalized, typed, agent-safe workbook JSON
```

“Full feature” in this plan means broad semantic coverage of the workbook
features users can inspect in Excel, LibreOffice Calc, and Google Sheets
exports. It does not mean byte-for-byte preservation, formula execution, macro
execution, query refresh, or writing XLSX files.

The reader is complete only when the required rows in the feature matrix are
implemented through public contracts and independently tested. Returning a
cell grid while ignoring styles, formulas, relationships, or advanced workbook
parts is not considered full-feature support.

## Audit findings incorporated

This plan was audited against the failure patterns already found in PPTX and
the semantics that make spreadsheets different from presentations. The audit
resolved the following design risks before implementation:

| Finding                                                                                                                | Resolution in this plan                                                                                                                 |
| ---------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| A cell could expose both `value` and `formula.cached`, allowing contradictory states.                                  | Use one discriminated `content` union: blank, literal value, or formula with a cached result.                                           |
| Range selection could hide a shared-formula master or style dependency outside the selected rectangle.                 | Selection limits emitted payload, while the parser reads and bounds the minimum supporting metadata needed for correct selected values. |
| A worksheet dimension could trigger rectangular allocation or exclude authored cells when stale.                       | Treat `dimension` and row `spans` as hints only; stream authored cells and return a sparse model.                                       |
| Formula translation by string replacement would corrupt names, strings, structured references, or absolute references. | Require token-aware reference translation without formula evaluation.                                                                   |
| Host locale or time zone could change `displayText` and dates.                                                         | Formatting and date normalization must be deterministic and independent of host locale/time zone.                                       |
| Declared ZIP sizes alone cannot stop expansion attacks.                                                                | Enforce both archive preflight and actual streamed-byte accounting with early abort.                                                    |
| XML chunks can split BOMs, UTF-8 sequences, UTF-16 code units, entities, and tags.                                     | Require an incremental fatal decoder and chunk-boundary adversarial tests for the streaming path.                                       |
| OPC-equivalent or encoded paths could bypass duplicate/traversal checks.                                               | Add one package-wide canonical part-name policy and reject ambiguous normalized identities.                                             |
| Blob URLs created before a later failure could leak because no caller receives them.                                   | Revoke all parser-owned URLs on failure; transfer ownership only after a successful return.                                             |
| Functional tests would not detect quadratic scans or unbounded retained state.                                         | Add scale fixtures, peak-memory/time baselines, early-abort assertions, and bundle/dependency budgets.                                  |

These resolutions are public-contract constraints, not implementation
suggestions. Changing one later requires updated black-box evidence and a
documented migration.

## What the PPTX work teaches us

The XLSX implementation should reuse the proven engineering method, not copy
PowerPoint-specific code.

| PPTX lesson                                                   | XLSX application                                                                                                                              |
| ------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------- |
| The package manifest owns document order.                     | Read sheet order from `workbook.xml`; never infer it from ZIP entry names.                                                                    |
| A relationship is relative to the part that owns its `.rels`. | Resolve worksheet, drawing, table, comments, chart, pivot, and external-link relationships from their owning parts.                           |
| ZIP and XML are untrusted security boundaries.                | Apply compressed, expanded, per-part, XML-depth, XML-node, and format-unit limits before constructing public values.                          |
| Strict and tolerant modes need typed behavior.                | Required workbook failures reject; recoverable optional-part failures emit stable XLSX diagnostics in tolerant mode and throw in strict mode. |
| Parser state must be scoped to one parse.                     | Shared strings, style tables, relationships, formula groups, and media caches belong to one workbook parse only.                              |
| Domain parsers make fidelity changes reviewable.              | Cells, formulas, styles, tables, drawings, charts, pivots, comments, and validation rules each get an owning XLSX module.                     |
| Normalized output must not expose parser internals.           | Do not expose XML nodes, ZIP entries, relationship IDs, shared-string indexes, style XF indexes, or pivot cache internals.                    |
| Independent packages catch self-confirming tests.             | Add a raw-OOXML XLSX fixture builder that never calls production package or normalization helpers.                                            |
| Real producers expose different valid encodings.              | Maintain an Excel, LibreOffice Calc, and Google Sheets export corpus with semantic expectations.                                              |
| Unit coverage alone is insufficient.                          | Require browser, corpus, property/fuzz, mutation, package-export, and CLI gates.                                                              |
| Output growth is also a security boundary.                    | Represent worksheets sparsely and limit returned cells, text, ranges, formulas, styles, drawings, and pivot records.                          |

The history of the PPTX reader also shows the preferred delivery order:
define the public contract, build independent evidence, implement the narrowest
owner, add adversarial boundaries, and only then expand fidelity. Security and
reliability cannot be postponed until after all workbook features are present.

## Scope and explicit non-goals

### Required reader scope

- accept `ArrayBuffer`, `Uint8Array`, and `Blob` without mutating caller data;
- run in Node.js 20, 22, and 24 and in supported browsers;
- expose ESM, CommonJS, declarations, a format subpath, and CLI conversion;
- discover the workbook through OPC content types and package relationships;
- support the applicable Transitional and Strict OOXML namespaces plus Markup
  Compatibility selection without exposing fallback XML;
- preserve workbook sheet order and sheet visibility;
- return sparse cells, semantic values, formulas, cached results, styles, and
  workbook structures through a JSON-compatible model;
- support bounded images, drawings, charts, comments, tables, conditional
  formatting, data validation, pivot metadata, and other rows marked required
  in the feature matrix;
- keep all authored document data untrusted and never execute or fetch it;
- produce deterministic output and stable typed diagnostics;
- support caller selection of sheets and ranges so a large workbook does not
  have to become a huge agent payload.

### Non-goals for the XLSX reader

- evaluating or recalculating formulas;
- claiming a cached formula result is current;
- refreshing Power Query, external data connections, linked data types, or
  pivot caches;
- executing DDE, RTD, `WEBSERVICE`, `HYPERLINK`, add-ins, scripts, ActiveX,
  OLE objects, or any other active content;
- decrypting encrypted Office documents;
- parsing legacy `.xls`, binary `.xlsb`, or macro-enabled `.xlsm` files;
- writing or modifying XLSX packages, which requires the separate preservation
  and writer contract in [xlsx-roundtrip-plan.md](xlsx-roundtrip-plan.md);
- preserving unsupported XML inside the normalized reader model; round-trip
  mode instead carries an opaque source package under the separate contract;
- rendering a pixel-identical Excel worksheet;
- silently repairing malformed packages.

`.xlsm`, `.xlsb`, `.xls`, and XLSX writing require separate contracts and
security reviews. They must not be accepted merely because some package parts
look similar.

## Architectural decisions

### XLSX is a sibling format domain

The reader belongs under `src/formats/xlsx`; it must not branch through the
PPTX parser:

```text
root public API
  -> XLSX public API and types
    -> XLSX workbook orchestrator
      -> XLSX-owned domain modules
        -> shared format-neutral OPC/XML/binary primitives
```

The intended source layout is:

```text
src/formats/xlsx/
├── index.ts                    Public parse functions and public exports
├── types.ts                    Public workbook model and options
├── errors.ts                   XlsxParseError
├── parser.ts                   Package and workbook orchestration
└── internal/
    ├── context.ts              Parse-scoped workbook and worksheet state
    ├── xml-reader.ts           XLSX diagnostic adapter over shared readers
    ├── resource-limits.ts      Workbook-specific limits and diagnostics
    ├── workbook.ts             Workbook properties, sheets, names, calc mode
    ├── worksheet.ts            Sparse row and cell orchestration
    ├── cell-reference.ts       A1 references, ranges, and bounds
    ├── cell-value.ts           Typed cell and cached formula values
    ├── shared-strings.ts       Plain, rich, and phonetic shared strings
    ├── formula.ts              Formula kinds and shared-formula translation
    ├── styles.ts               XF resolution and normalized style table
    ├── number-format.ts        Built-in/custom formats and display semantics
    ├── date-system.ts          1900/1904 serial handling
    ├── dimensions.ts           Rows, columns, outlines, visibility, sizing
    ├── views.ts                Panes, selections, zoom, and sheet views
    ├── merge.ts                Validated non-overlapping merged ranges
    ├── hyperlink.ts            Internal and external links
    ├── comments.ts             Notes and threaded comments
    ├── table.ts                SpreadsheetML tables and totals
    ├── auto-filter.ts          Filters, custom filters, and sort state
    ├── data-validation.ts      Validation rules and prompts
    ├── conditional-format.ts   Rules and differential styles
    ├── drawing.ts              Worksheet anchors and drawing relationships
    ├── chart.ts                XLSX chart ownership and workbook references
    ├── image.ts                Bounded embedded image representations
    ├── sparkline.ts            Sparkline groups and ranges
    ├── pivot.ts                Pivot definitions and bounded cache data
    ├── slicer.ts               Slicer/timeline metadata and relationships
    ├── print.ts                Page setup, margins, breaks, headers/footers
    ├── protection.ts           Workbook/sheet protection metadata
    └── external-data.ts        Safe metadata; never fetch or expose secrets
```

Files should be introduced only when their observable contract is ready. The
tree is an ownership map, not permission to create empty abstractions.

### Worksheets use a sparse model

An Excel worksheet has 1,048,576 rows and 16,384 columns. The reader must never
allocate the rectangular grid implied by a dimension such as `A1:XFD1048576`.
Only authored cells and explicit row/column/range structures are returned.

The following remain separate:

- a missing cell;
- an explicitly empty cell with a style;
- an empty string;
- a formula with no cached result;
- a malformed cell omitted with a diagnostic;
- a cell outside the caller's selected ranges.

That distinction is necessary for formatting, table boundaries, merged cells,
and reliable agent reasoning.

### Formulas are data, never executable instructions

The reader preserves a normalized formula expression and its cached result as
separate values. It does not evaluate the expression or infer that the cached
result is current.

- shared-formula storage is expanded into the expression applicable to each
  returned cell;
- array, data-table, and dynamic-array ranges remain explicit semantic kinds;
- formula text, defined names, and external references remain untrusted data;
- volatile functions, external workbook links, DDE, RTD, and web functions are
  never executed or fetched;
- workbook calculation properties are returned so callers can assess whether
  caches may be stale;
- a missing or invalid cached result is different from an empty result.

Formula expressions are returned without a leading `=` and with XML entities
decoded. Unknown functions, `_xlfn.` functions, future function names, and
valid syntax the reader does not semantically understand are preserved as
data; they are not rejected merely because OAKit cannot calculate them.

Shared-formula translation must use a formula-aware tokenizer. It shifts only
relative A1 references and preserves string literals, escaped quotes, absolute
and mixed references, names, structured references, sheet-qualified and 3D
references, external workbook references, unions, intersections, and locale-
invariant separators. It must never use broad regular-expression replacement.

Implementing a formula engine would be a separate product with its own parser,
dependency graph, compatibility matrix, resource limits, and security review.

### Large XML parts need a streaming path

The current PPTX control parts fit the bounded whole-tree XML reader. XLSX
worksheets, shared strings, and pivot cache records can be much larger. Loading
and then parsing each large part twice would create avoidable memory pressure.

Use two paths:

1. bounded whole-tree parsing for small structural parts such as workbook
   metadata, relationships, styles, tables, and chart definitions;
2. bounded SAX/event parsing for worksheet rows/cells, shared strings, and
   pivot cache records.

The streaming path must enforce the same XML correctness rules as the existing
reader: fatal decoding, forbidden declarations, exact element nesting, valid
qualified names, namespace-aware attributes, depth limits, node limits, byte
limits, and aggregate parse budgets. It must stop decompression and parsing as
soon as a limit is crossed.

Streaming decoding must preserve state across arbitrary chunks. UTF-8 and
UTF-16 BOMs, multibyte characters, surrogate pairs, entities, comments, CDATA,
attributes, and start/end tags may cross chunk boundaries. Skipped XML branches
remain structurally validated and counted; selection must never become a path
that bypasses XML validation.

Markup Compatibility processing uses an explicit versioned set of understood
namespace URIs. `AlternateContent` selects the first `Choice` whose complete
`Requires` set is understood, otherwise its `Fallback`; unknown ignorable
content is skipped or processed only as directed by `ProcessContent`; unknown
non-ignorable content is not silently accepted. Prefix spelling and local-name
similarity never imply namespace support.

### Reader and writer models remain separate

The normalized reader model is optimized for inspection and safe JSON output.
A future writer must not reverse it mechanically. Writing requires explicit
allocation of style IDs, shared strings, relationship IDs, calculation state,
content types, extension parts, and package-valid XML.

### Rejected shortcuts

- A dense `Cell[][]` grid is rejected because it turns a sparse format into an
  allocation proportional to the theoretical worksheet rectangle.
- Returning shared-string, XF, or relationship indexes is rejected because it
  leaks package encoding and forces consumers to rebuild parser state.
- A single generic `OfficeDocument` model is rejected because slide geometry
  and spreadsheet cells/ranges have different invariants.
- Importing PPTX table, chart, text, style, context, or diagnostic modules is
  rejected; only independently proven format-neutral primitives may move to
  `common`.
- Parsing the entire workbook and filtering selection afterward is rejected
  because it does not bound work or output.
- Regular-expression formula shifting, host-locale `Intl` formatting, and
  JavaScript `Date` output are rejected because they are not deterministic or
  semantically safe.
- Adding a broad spreadsheet library as the implementation and test oracle is
  rejected because one dependency cannot independently prove its own formula,
  style, relationship, or security behavior.

## Proposed public API

The exact names must be frozen through contract tests before implementation,
but the initial shape should be close to the existing PPTX API:

```ts
export async function parseXlsx(
  input: XlsxInput,
  options?: XlsxParseOptions,
): Promise<XlsxDocument>;

export async function parseXlsxWithDiagnostics(
  input: XlsxInput,
  options?: XlsxParseOptions,
): Promise<XlsxParseResult>;

export type XlsxInput = ArrayBuffer | Uint8Array | Blob;

export interface XlsxParseOptions {
  errorMode?: 'tolerant' | 'strict';
  displayTextMode?: 'supported' | 'none';
  imageMode?: 'base64' | 'blob' | 'both' | 'none';
  pivotCacheMode?: 'metadata' | 'records' | 'none';
  selection?: XlsxSelection;
  limits?: XlsxResourceLimits;
}
```

Option defaults are part of the public contract:

| Option            | Default              | Contract                                                                                                 |
| ----------------- | -------------------- | -------------------------------------------------------------------------------------------------------- |
| `errorMode`       | `tolerant`           | Recover only documented optional features. Core worksheet, security, and resource failures remain fatal. |
| `displayTextMode` | `supported`          | Emit deterministic display text only for completely supported applied formats.                           |
| `imageMode`       | `none`               | Preserve safe drawing/image metadata without embedding binary representations.                           |
| `pivotCacheMode`  | `metadata`           | Return cache structure without cache records.                                                            |
| `selection`       | all sheets           | Return every sheet payload subject to limits; callers should select ranges for large agent workloads.    |
| `limits`          | safe format defaults | Bound archive, XML, scanned work, output, media, and advanced feature graphs.                            |

Values, formulas, workbook structure, and bounded feature metadata are always
part of the default result. Binary image content and pivot cache records require
explicit modes because they can dominate output size. Intentionally disabled
display/image/pivot payloads are omitted without an unsupported-feature
diagnostic.

### Selection contract

Selection is an output and work bound, not a post-parse filter:

```ts
export interface XlsxSelection {
  sheetNames?: readonly string[];
  ranges?: Readonly<Record<string, readonly string[]>>;
}
```

- an absent `selection` means all sheet payloads; `sheetNames` selects full
  payloads, `ranges` selects worksheet rectangles, and a full-sheet selection
  wins if both target the same worksheet;
- sheet names use a deterministic whole-name, case-insensitive comparison with
  no trimming or locale-dependent casing; returned names preserve authored
  code points and casing;
- range values are unqualified `A1`, `A1:C10`, whole-column, or whole-row
  references because the map key already identifies the sheet; names, 3D
  ranges, external references, and ranges on chart sheets are rejected;
- selections are parsed and validated before worksheet payloads are read;
- workbook metadata and the ordered sheet list remain available even if cell
  payloads are not selected;
- a sheet reports whether its payload was selected, so an unselected sheet is
  not confused with an empty sheet;
- overlapping selections are deduplicated without duplicating cells;
- selection does not change workbook semantics: the parser may read formula
  masters, shared strings, names, styles, row/column defaults, relationships,
  and other bounded dependencies outside the emitted ranges;
- every encountered cell is charged to a scanned-work limit even if it is not
  returned, while `maxReturnedCells` separately bounds public output;
- cells are emitted only inside selected ranges, but a merge, table, validation,
  conditional-format rule, or drawing that intersects a selected range is
  returned with its original unclipped range and an explicit intersection
  relation;
- worksheet metadata states whether the whole sheet or only explicit ranges
  were selected;
- absolute-position drawings have no cell intersection and are returned only
  for a full-sheet payload; range-only selection omits them without treating
  that intentional omission as unsupported content;
- unknown sheet names, conflicting duplicate names, or invalid range/sheet-kind
  combinations produce `invalid-selection` rather than an empty workbook;
- invalid ranges produce typed diagnostics, never a best-effort rectangle;
- limits still apply to the selected result and to package/XML processing.

### Workbook model

The model below is illustrative. It defines the important semantic boundaries,
not final field names:

```ts
export interface XlsxDocument {
  workbook: XlsxWorkbookProperties;
  styles: XlsxStyle[];
  sheets: XlsxSheet[];
}

export interface XlsxWorkbookProperties {
  dateSystem: '1900' | '1904';
  calculation: {
    mode: 'automatic' | 'automatic-except-tables' | 'manual';
    fullCalculationOnLoad: boolean;
    forceFullCalculation: boolean;
  };
  definedNames: XlsxDefinedName[];
}

export type XlsxSheet = XlsxWorksheet | XlsxChartSheet;

export interface XlsxWorksheet {
  kind: 'worksheet';
  index: number;
  name: string;
  state: 'visible' | 'hidden' | 'very-hidden';
  payload: 'not-selected' | 'full-sheet' | 'selected-ranges';
  rows: XlsxRow[];
  columns: XlsxColumnRange[];
  mergedRanges: XlsxRange[];
  tables: XlsxTable[];
  drawings: XlsxDrawing[];
  // Views, filters, validations, comments, print settings, and other
  // worksheet-level structures are normalized alongside these fields.
}

export interface XlsxRow {
  index: number;
  hidden?: boolean;
  height?: number;
  outlineLevel?: number;
  cells: XlsxCell[];
}

export interface XlsxCellBase {
  address: string;
  column: number;
  style?: number;
  displayText?: string;
}

export type XlsxCell = XlsxCellBase &
  (
    | { content: { kind: 'blank' } }
    | { content: { kind: 'value'; value: XlsxCellValue } }
    | {
        content: {
          kind: 'formula';
          formula: XlsxFormula;
          cached: XlsxCellValue | { kind: 'missing' };
        };
      }
  );
```

`style` refers to an OAKit-normalized entry in `document.styles`; it is not a
raw `cellXfs` index. A deterministic normalized table avoids copying large
style objects into every cell while keeping OOXML storage details internal.

### Cell values

Cell values require discriminants so missing, blank, malformed, and cached
formula states do not collapse into `null`:

```ts
export type XlsxCellValue =
  | { kind: 'text'; text: string; runs?: XlsxRichTextRun[] }
  | { kind: 'number'; value: number }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error'; code: string }
  | {
      kind: 'date';
      precision: 'date' | 'time' | 'date-time' | 'duration';
      normalized: string | null;
      source:
        | { kind: 'serial'; value: number; dateSystem: '1900' | '1904' }
        | { kind: 'iso'; value: string };
    };
```

`normalized` is `null` for a serial that cannot represent a real calendar instant,
including the compatibility-only 1900 leap-day serial. Date-only, time-only,
duration, and date-time formats must be distinguished. The original finite
serial remains available so normalization does not invent a date.

`t="d"` ISO date cells retain an ISO source instead of fabricating a numeric
serial. Public values never contain JavaScript `Date` objects because host time
zones would make them non-deterministic. An explicit blank cell is represented
by `content.kind === "blank"`; blank is not a literal cell value.

### Formula model

```ts
export interface XlsxFormula {
  expression: string;
  kind: 'normal' | 'array' | 'data-table' | 'dynamic-array';
  range?: XlsxRange;
}
```

Shared formulas are a package compression mechanism, not a public formula
kind. The parser translates their relative references independently and emits
the normal expression for each returned cell. Translation must be tested for
absolute, relative, mixed, sheet-qualified, 3D, table, and external references.
The shared master may be outside the caller's selected ranges, so the streaming
state machine must retain bounded master metadata until all selected dependents
are resolved.

### Styles and display text

Normalized styles should include:

- fonts, theme colors, tints, and rich-text runs;
- pattern and gradient fills;
- border sides, diagonal borders, colors, and line styles;
- horizontal/vertical alignment, wrapping, indentation, shrink-to-fit, and
  text rotation;
- built-in and custom number formats;
- locked/hidden protection metadata;
- differential styles used by conditional formatting;
- named cell styles where they carry user-visible semantics.

Base style resolution and conditional formatting remain separate. The reader
must not claim a conditional format is active unless its rule can be evaluated
without executing formulas. Unsupported display-format tokens produce a
diagnostic and omit `displayText`; they do not silently fall back to a misleading
string.

`displayText` is a deterministic convenience view, never the source value. It
may be emitted only when OAKit supports the complete applied number-format
section. Formatting must use explicit Excel-compatible rules rather than the
host's default `Intl`, locale, calendar, or time zone. Locale-dependent formats
that cannot be reproduced deterministically retain their normalized format code
and omit `displayText` with a diagnostic.

Repeated unsupported-format diagnostics are deduplicated by normalized style
and format code rather than emitted once per cell. Repeated emitted text still
counts after shared-string expansion because serialized JSON repeats that text.

Semantic cell styles may be deduplicated, but named styles, differential styles,
and distinct user-visible style names must not be collapsed merely because
their current visual fields compare equal. Style references are stable only
within one returned document unless a later public contract defines stronger
identity.

## Feature matrix

“Required” means normalized public support is necessary. “Metadata” means the
safe structure described in the target is the complete reader commitment; no
execution or refresh is implied. “Diagnostic” means recognizing and safely
omitting the content with a stable code is the complete commitment. Every row
must meet its assigned class before the XLSX reader is described as
full-feature.

| Area                     | Class      | Completion target                                                                                                                     | Acceptance evidence                                                                         |
| ------------------------ | ---------- | ------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------- |
| OPC package              | Required   | Content types, root relationships, workbook discovery, canonical part ownership                                                       | Missing/wrong roots, relocated workbook, traversal, ambiguous path, internal/external tests |
| Namespaces/compatibility | Required   | Applicable Transitional/Strict namespaces, `mc:AlternateContent`, Ignorable and ProcessContent rules                                  | Namespace aliases, deterministic choice/fallback, malformed MC tests                        |
| Workbook                 | Required   | Ordered sheets, active sheet, visibility, properties, date system, calculation properties                                             | Excel/Calc/Sheets corpus and exact order/state assertions                                   |
| Calculation metadata     | Metadata   | Calculation mode/version, iteration settings, full-calc flags, calculation-chain metadata                                             | Stale/missing chain cases; no dependency execution or cache-freshness claim                 |
| Sheet kinds              | Required   | Worksheets and chart sheets                                                                                                           | Public discriminated union and independent fixtures                                         |
| Sheet metadata           | Required   | Tab colors, dimensions, defaults, custom views, scenarios, selected/active state                                                      | Producer fixtures and malformed range/value diagnostics                                     |
| Cells                    | Required   | Sparse rows/cells, omitted/inferred references, blank styled cells, booleans, text, numbers, errors, inline strings, direct ISO dates | Type matrix, lexical errors, duplicates, order, omitted refs, grid bounds                   |
| Shared strings           | Required   | Plain strings, rich runs, whitespace preservation, phonetic metadata                                                                  | Index boundaries, rich text, missing/malformed table, output text limits                    |
| Formulas                 | Required   | Normal, shared, array, data-table, dynamic array, cached values, future/unknown functions                                             | Token-aware translation, stale/missing cache, preservation, no execution/fetch              |
| Modern cell metadata     | Required   | Known rich values, cell images, checkboxes, spill metadata; linked-data metadata without refresh                                      | Versioned extension fixtures, bounded payloads, stable unsupported diagnostics              |
| References               | Required   | A1 cells/ranges, quoted sheets, 3D references, external references, structured table references                                       | Seeded properties and minimized invalid references                                          |
| Defined names            | Required   | Workbook- and sheet-scoped names, print area/titles, hidden names                                                                     | Scope/collision tests; formula text remains untrusted                                       |
| Styles                   | Required   | Fonts, fills, borders, alignment, number formats, protection, named/differential styles                                               | Style precedence matrix and mutation-tested normalization                                   |
| Dates/times              | Required   | 1900/1904 systems, direct ISO dates, serial 0/59/60/61, time-only, duration, date-time                                                | Exact boundaries across producers, locales, and time zones                                  |
| Rows/columns             | Required   | Sizes, hidden state, outline levels, collapsed groups, default styles                                                                 | Overlapping column-range precedence and boundaries                                          |
| Views                    | Required   | Freeze/split panes, selections, zoom, right-to-left, gridline/header visibility                                                       | Producer corpus and malformed pane references                                               |
| Merges                   | Required   | Valid, non-overlapping merged ranges                                                                                                  | Reversed, overlapping, duplicate, out-of-bounds, at-limit tests                             |
| Hyperlinks               | Required   | Internal locations and allowlisted external protocols                                                                                 | No navigation/fetch; unsafe protocol and relationship tests                                 |
| Notes/comments           | Required   | Legacy notes, authors, threaded comments, persons, visible state                                                                      | Untrusted text and missing person/comment relationships                                     |
| Tables                   | Required   | Names, ranges, columns, headers, totals, calculated-column formulas, styles                                                           | Range/column cardinality and relationship ownership                                         |
| Filters/sorts            | Required   | AutoFilter, custom/dynamic/top filters, color/icon filters, sort conditions                                                           | Semantic rules without applying destructive filtering                                       |
| Data validation          | Required   | Types, operators, formulas, prompts, error messages, ranges                                                                           | Multi-range, invalid formula/range, list-source tests                                       |
| Conditional formatting   | Required   | Cell, expression, color scale, data bar, icon set, differential style, priority/stop-if-true                                          | Authored order, priority, ranges, extension compatibility                                   |
| Drawings/images          | Required   | One-cell, two-cell, absolute anchors, crop/transform, bounded media                                                                   | Anchor geometry, relationship ownership, browser URL lifecycle                              |
| Charts                   | Required   | Common chart families, series formulas/caches, axes, titles, legends, colors, styles                                                  | XLSX-owned chart tests and cross-producer corpus                                            |
| Sparklines               | Required   | Groups, data ranges, locations, axis/color options                                                                                    | x14 extension fixtures and range validation                                                 |
| Pivot tables             | Required   | Definitions, fields, row/column/data axes, filters, styles                                                                            | Bounded normalized model and producer fixtures                                              |
| Pivot caches             | Metadata   | Cache definition always; records only in explicit mode                                                                                | Record/text limits and no-refresh behavior                                                  |
| Slicers/timelines        | Metadata   | Cache links, pivot/table ownership, safe display metadata                                                                             | Relationship graph and extension namespaces                                                 |
| Print/layout             | Required   | Margins, orientation, paper size, scaling, repeating titles, breaks, headers/footers                                                  | Unit/boundary tests and escaped text                                                        |
| Protection               | Metadata   | Workbook/sheet protected state and algorithm metadata                                                                                 | Never claim password validation or decryption                                               |
| External links           | Metadata   | Safe target and formula metadata only                                                                                                 | Never open linked workbooks; unsafe targets diagnosed                                       |
| Connections/query tables | Metadata   | Safe redacted metadata only                                                                                                           | Never refresh; credentials/connection strings never exposed                                 |
| Active/embedded content  | Diagnostic | Recognize and omit OLE, ActiveX, scripts, executables                                                                                 | Stable security diagnostic and no active payload read                                       |
| Known extensions         | Required   | Normalize known Excel extension namespaces per owning feature                                                                         | Versioned producer fixtures and namespace tests                                             |
| Unknown extensions       | Diagnostic | Safely omit unknown optional extensions                                                                                               | Stable diagnostic without raw XML escape                                                    |
| Document properties      | Required   | Core/app/custom properties using typed safe values                                                                                    | Malformed property types and untrusted text tests                                           |

“Full-feature” is a versioned support profile, not a promise to understand
every Excel feature Microsoft may add later. Before the support claim ships,
publish a machine-reviewable manifest of supported ECMA-376 edition features,
Microsoft extension namespace URIs, feature classes, and tested producer
versions. A newly introduced namespace remains an unknown optional extension
until added with fixtures, public types, corpus evidence, and mutation coverage;
it does not retroactively make an older release unsafe or dishonest.

## Package graph and parse pipeline

### 1. Validate input and archive

- resolve and validate all configured limits before reading input;
- reject oversized compressed input before `JSZip.loadAsync`;
- reject invalid ZIP data with `XlsxParseError` and `invalid-package`;
- validate entry count, declared expanded sizes, per-part size, and total
  expansion before semantic parsing;
- enforce actual expanded-byte budgets while every part streams, even when ZIP
  metadata is missing, false, or inconsistent;
- reject unsafe entry names and duplicate normalized part names;
- canonicalize part identities before lookup and reject traversal, encoded
  separators/dot segments, invalid percent encodings, query/fragment confusion,
  and ambiguous identities while preserving valid case-sensitive part names;
- create one parse-scoped budget tracker for expanded bytes, XML nodes, scanned
  cells, relationships, text, formulas, ranges, and returned objects.

### 2. Discover the workbook through OPC

- read and validate `[Content_Types].xml`;
- read package `_rels/.rels`;
- select exactly one internal `officeDocument` relationship with an XLSX
  workbook content type;
- resolve it through shared OPC URI handling rather than hard-coding
  `xl/workbook.xml`;
- reject duplicate relationship IDs, invalid relationship types, ambiguous main
  parts, and targets that canonicalize outside the package root;
- reject macro-enabled, binary, encrypted, or ambiguous main parts;
- validate the workbook root before reading optional content;
- normalize the applicable Strict/Transitional namespace aliases and process
  Markup Compatibility choices before domain dispatch.

This is intentionally stricter than assuming the conventional path. Valid OPC
packages may relocate parts, while attackers may exploit assumptions about
paths or relationship owners.

### 3. Read workbook-level tables

Parse workbook properties, ordered sheet declarations, defined names,
calculation properties, styles, theme, shared strings, external-link metadata,
and workbook relationships. Required parts and optional parts must have
different failure behavior.

Create immutable normalized workbook tables before parsing cells:

- resolved sheet descriptors;
- normalized style table;
- shared-string accessor with aggregate text accounting;
- date-system and calculation metadata;
- safe external-link descriptors;
- relationship maps scoped to each owning part.

### 4. Select sheets and ranges

Validate selection syntax before worksheet traversal. Preserve every sheet's
ordered metadata, but only read payloads requested by selection. A missing
selected sheet or invalid range is observable; it must not look like an empty
selection.

Build a bounded selection index by row intervals so membership checks do not
scan every selected range for every cell. Read global names, shared strings,
styles, row/column defaults, shared-formula masters, and relationship metadata
needed by selected cells even when those dependencies are outside emitted
ranges. A dependency read is charged to work/resource limits but is not emitted
as a selected cell.

### 5. Stream worksheet cells

For each selected worksheet:

1. validate the worksheet root and namespace;
2. read dimension hints but never allocate from them;
3. stream rows and cells in authored order;
4. infer a spec-valid omitted row as previous row plus one (first row is 1) and
   an omitted cell as previous column plus one within that row (first cell is
   column 1), while validating explicit indexes, forward ordering, worksheet
   bounds, and duplicates;
5. resolve value kind, shared string, formula, cached value, and base style;
6. retain explicit blank styled cells;
7. account separately for scanned cells, returned cells, formula groups,
   formula characters, rich-text runs, and total text characters;
8. emit rows and cells in deterministic coordinate order;
9. parse worksheet-level ranges and relationships through their owning part.

The worksheet `dimension` and row `spans` attributes are optimization hints,
not authorities. They may be stale or wider than authored data and must not
exclude cells or allocate a dense grid. Malformed selected cell coordinates,
values, or formulas are core-data failures, not optional content that tolerant
mode may silently omit. Resource-limit violations always abort the parse.

### 6. Resolve worksheet structures

Parse columns, merges, views, filters, tables, validation, conditional
formatting, hyperlinks, comments, drawings, print settings, protection, and
extension features. Each parser receives a worksheet-scoped context and cannot
read a relationship map owned by another sheet.

### 7. Resolve visual and analytical parts

Follow drawing relationships to images and charts. Follow pivot, slicer, and
timeline graphs from their actual owners. Deduplicate media bytes within one
parse, but never share caches across workbooks or concurrent calls.

Chart series may contain both formula references and cached values. Preserve
both and never fetch an external workbook to fill a missing cache.

### 8. Finalize normalized output

- ensure no `NaN`, `Infinity`, unsafe range index, raw XML node, relationship
  ID, or internal exception escapes;
- ensure all public arrays have deterministic order;
- deduplicate normalized styles deterministically;
- retain diagnostics in discovery order with duplicate failures suppressed;
- verify returned counts against output limits;
- revoke all object URLs created during a parse that later fails;
- transfer ownership of object URLs to the caller only after a successful
  return and document cleanup.

## What can move to `common`

XLSX becomes the second real consumer of several package primitives. That is
the point at which extraction is justified, but each extraction should be a
separate behavior-preserving change with PPTX and shared tests.

| Candidate                                       | Decision                                                                                                                                                                                     |
| ----------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| ZIP entry streaming and byte budgets            | Already format-neutral; reuse directly.                                                                                                                                                      |
| OPC part URI and relationship target resolution | Reuse the format-neutral base only after adding package-root ownership, canonical percent/path handling, and ambiguous-name tests required by a second format.                               |
| Content-type parsing                            | Extract a shared parser after independent PPTX/XLSX contract tests define duplicate, missing, and malformed behavior.                                                                        |
| Relationship XML parsing                        | Extract owner-scoped relationships with internal/external discriminants, duplicate-ID checks, type validation, and canonical targets; format adapters map failures to their own diagnostics. |
| Archive preflight and aggregate budgets         | Extract generic mechanics; keep `PptxResourceLimits` and `XlsxResourceLimits` as public format contracts.                                                                                    |
| XML whole-tree reader                           | Reuse for bounded small parts.                                                                                                                                                               |
| XML streaming reader                            | Add format-neutral event and budget mechanics; keep worksheet state machines under XLSX.                                                                                                     |
| Binary/base64 and media MIME helpers            | Reuse directly.                                                                                                                                                                              |
| Safe hyperlink protocol handling                | Reuse policy and serialization; keep hyperlink semantics format-owned.                                                                                                                       |
| Numeric lexical validation                      | Reuse generic integer/finite/range primitives; keep Excel serial and cell rules under XLSX.                                                                                                  |
| DrawingML colors and transforms                 | Extract only after XLSX fixtures prove the same DrawingML contract. SpreadsheetML style colors still belong to XLSX.                                                                         |
| DrawingML chart primitives                      | Do not import `pptx/internal/chart.ts`. First build the XLSX owner; extract only truly identical chart normalization with two-format tests.                                                  |
| PowerPoint tables                               | Never reuse. A slide table and a SpreadsheetML table have different models and inheritance.                                                                                                  |
| PowerPoint text HTML                            | Never reuse as the XLSX cell model. Reuse escaping primitives only where an output format actually needs escaped text.                                                                       |
| PPTX parser context or diagnostics              | Never reuse directly. Mirror the lifecycle pattern with XLSX-owned types and errors.                                                                                                         |

The migration rule is: no PPTX behavior may change accidentally while making a
helper generic. Run focused PPTX tests before and after every extraction and do
not combine extraction with a new XLSX feature in one commit.

## Security and resource limits

XLSX inherits every OPC, ZIP, XML, relationship, numeric, and media invariant
from the PPTX reader. It adds limits specific to large worksheets and workbook
graphs.

Candidate public limit names are:

```ts
export interface XlsxResourceLimits {
  maxInputBytes?: number;
  maxEntries?: number;
  maxTotalUncompressedBytes?: number;
  maxPartBytes?: number;
  maxXmlBytes?: number;
  maxXmlDepth?: number;
  maxXmlNodes?: number;
  maxTotalXmlNodes?: number;
  maxMediaBytes?: number;
  maxWorksheets?: number;
  maxRelationships?: number;
  maxDefinedNames?: number;
  maxScannedCells?: number;
  maxReturnedCells?: number;
  maxRowsPerWorksheet?: number;
  maxColumnsPerWorksheet?: number;
  maxSharedStrings?: number;
  maxRichTextRuns?: number;
  maxTextCharacters?: number;
  maxFormulaCharacters?: number;
  maxTotalFormulaCharacters?: number;
  maxFormulaGroups?: number;
  maxStyles?: number;
  maxMergedRanges?: number;
  maxRangeAreas?: number;
  maxTables?: number;
  maxHyperlinks?: number;
  maxValidationRules?: number;
  maxConditionalFormattingRules?: number;
  maxComments?: number;
  maxDrawings?: number;
  maxCharts?: number;
  maxPivotRecords?: number;
}
```

Defaults are not frozen by this plan. Start with the existing package/XML
budgets, then benchmark candidate worksheet, cell, text, style, and pivot
limits against the producer corpus and realistic agent output sizes. A larger
valid Excel grid is not automatically a safe default payload. Callers can
raise positive safe-integer limits explicitly when their isolation and memory
budgets allow it.

Limit relationships are part of validation. `maxXmlBytes` and `maxMediaBytes`
cannot exceed `maxPartBytes`; `maxReturnedCells` cannot exceed
`maxScannedCells`; per-formula characters cannot exceed the aggregate formula
budget; configured row and column bounds cannot exceed XLSX's fixed grid; and
all count/byte arithmetic must remain within safe integers. Format hard maxima
cannot be raised by options.

Accounting definitions must be frozen in tests. `maxScannedCells` counts every
worksheet cell encountered before selection; `maxReturnedCells` counts public
cell objects; text counts the characters that will appear after shared-string
expansion, including repeated references; formula totals count normalized
expressions after shared-formula expansion; and relationship/range/style counts
include entries that are parsed but later found unreferenced. This prevents a
selection or deduplication path from hiding attacker-controlled work.

Every limit requires:

- invalid configuration tests;
- below-limit, exactly-at-limit, and one-over-limit tests;
- strict and tolerant assertions proving the violation is fatal in both;
- the precise limit name, actual value, limit, and offending part when known;
- mutation coverage for boundary comparisons and accounting updates.

### Active and external content

- external relationships are preserved only when `TargetMode="External"` is
  explicit;
- no external relationship is fetched;
- external URI user information and connection secrets are redacted before
  public output or diagnostics; diagnostic messages never echo full formulas,
  connection strings, comments, or cell text;
- embedded packages, ActiveX, OLE, scripts, and executable content are never
  loaded as active objects;
- formula expressions are never executed;
- external connection strings and credentials are not returned; safe metadata
  can state that a connection exists and that sensitive content was omitted;
- hyperlink protocols use an allowlist;
- comments, names, formulas, alt text, properties, filenames, and sheet values
  remain untrusted data, never agent instructions;
- a macro-enabled, binary, or encrypted main workbook is rejected in both
  modes; optional embedded OLE/ActiveX content is omitted with a security
  diagnostic in tolerant mode and rejected in strict mode;
- object URLs are created only in explicit blob modes, revoked internally if
  parsing fails, and transferred to the caller only on success.

## Diagnostics and recovery

XLSX owns `XlsxDiagnostic`, `XlsxDiagnosticCode`, and `XlsxParseError`. Initial
codes should cover the proven PPTX categories and XLSX-specific outcomes:

```text
invalid-package
invalid-document-structure
invalid-document-value
invalid-cell-reference
invalid-formula
invalid-selection
invalid-relationship-target
missing-required-part
resource-limit-exceeded
security-rejected-content
unsupported-feature
xml-parse-failed
xml-read-failed
```

Diagnostic messages are bounded explanatory text, not a machine contract.
Callers branch on code and structured fields such as `part`, `sheet`, `cell`,
`range`, `relationshipType`, and resource-limit metadata. Untrusted document
content is omitted, redacted, or length-bounded in messages so one formula or
comment cannot become a log-injection or output-growth path.

Recovery rules:

| Condition                                                                       | Tolerant                                                                    | Strict                                          |
| ------------------------------------------------------------------------------- | --------------------------------------------------------------------------- | ----------------------------------------------- |
| Invalid ZIP, missing workbook, unsafe root relationship, incompatible main part | Throw                                                                       | Throw                                           |
| Resource limit or unsafe OPC/XML structure                                      | Throw                                                                       | Throw                                           |
| Malformed required workbook or selected worksheet root/cell stream              | Throw                                                                       | Throw                                           |
| Malformed selected cell coordinate, value, formula, or required dependency      | Throw                                                                       | Throw                                           |
| Missing/malformed optional comments, drawing, chart, or pivot part              | Omit that optional feature and diagnose                                     | Throw                                           |
| Unsupported valid optional feature                                              | Return documented safe metadata when available; otherwise omit and diagnose | Throw because authored content would be omitted |
| Optional embedded OLE/ActiveX content                                           | Omit and emit `security-rejected-content`                                   | Throw                                           |
| Intentionally disabled image/pivot output mode                                  | Omit without warning                                                        | Omit without warning                            |

Tolerant mode is for documented optional-feature recovery. It does not permit a
partial selected worksheet core because silently dropping a malformed financial
cell can be materially misleading. XML validity and resource/security boundaries
remain fatal regardless of whether a malformed node would have been selected.

## Producer compatibility

The corpus should cover at least:

- Microsoft Excel for Windows;
- Microsoft Excel for macOS;
- LibreOffice Calc;
- Google Sheets exported as XLSX.

Optional compatibility coverage can include Apple Numbers exports and files
created by common server-side generators, but they do not replace the required
producer set.

Each corpus case records:

- source and producer/version;
- license or provenance;
- whole-file SHA-256 when exports are stable;
- bounded file and expanded sizes;
- ordered sheet names and visibility;
- minimum non-empty cell, formula, style, table, chart, comment, or pivot counts
  relevant to that fixture;
- selected representative values and formula expressions;
- expected diagnostics;
- a documented fingerprint strategy for non-deterministic exports.

Corpus assertions must be semantic. A parser that returns ordered empty sheets
must not pass a workbook containing values, formulas, charts, or tables.

## Independent test architecture

### Raw black-box fixture builder

Add `test/black-box/xlsx-package.ts` with an independently authored minimal OPC
package containing:

- `[Content_Types].xml`;
- package relationships;
- a workbook and workbook relationships;
- one worksheet;
- optional styles and shared strings;
- helpers that replace or remove exact parts.

The builder uses raw XML and JSZip only. It must not call production URI,
relationship, cell-reference, style, formula, or normalization helpers.
It needs variants for a relocated workbook part, Strict namespaces, Markup
Compatibility choices, omitted row/cell references, direct ISO dates, stale
dimensions, shared-formula masters outside selection, and chunked large parts.

Public tests import `parseXlsx` from `src`, never private parsers. Expected
values are literal semantic values, not values calculated with production
algorithms.

### Focused test suites

```text
test/black-box/xlsx-public-api.test.ts
test/black-box/xlsx-package-validity.test.ts
test/black-box/xlsx-adversarial.test.ts
test/xlsx/workbook.test.ts
test/xlsx/worksheet.test.ts
test/xlsx/cell-value.test.ts
test/xlsx/cell-reference.test.ts
test/xlsx/formula.test.ts
test/xlsx/shared-strings.test.ts
test/xlsx/styles.test.ts
test/xlsx/number-format.test.ts
test/xlsx/date-system.test.ts
test/xlsx/dimensions.test.ts
test/xlsx/merge.test.ts
test/xlsx/table.test.ts
test/xlsx/filter.test.ts
test/xlsx/data-validation.test.ts
test/xlsx/conditional-format.test.ts
test/xlsx/comments.test.ts
test/xlsx/drawing.test.ts
test/xlsx/chart.test.ts
test/xlsx/pivot.test.ts
test/xlsx/resource-limits.test.ts
test/browser/xlsx-browser.test.ts
test/corpus/xlsx-corpus.test.ts
```

Test data remains independent from production behavior. Snapshot-only tests are
not sufficient for formulas, relationships, ranges, styles, diagnostics,
ordering, or limits.

### Conformance oracles

Literal public semantic assertions are the primary oracle. Secondary
differential checks may compare representative fixtures with Microsoft Excel,
LibreOffice Calc, the Open XML SDK validator, or another independent reader,
but no one producer/library becomes the definition of correctness. Differential
disagreement must be reduced to the normative package structure and a literal
expected public value before a regression is accepted.

Tests must never compute expected formula shifts, date serials, style
precedence, or range intersections with the production helper being tested.
Golden JSON can aid review but cannot be the only evidence for a semantic
contract.

### Required public invariants

- identical document output for equivalent `Uint8Array`, `ArrayBuffer`, and
  `Blob` inputs;
- caller input and options are not mutated;
- sequential and concurrent parses are deterministic and isolated;
- sheet order follows the workbook manifest;
- cells are sparse, ordered, and unique by coordinate;
- every public number is finite and valid for its semantic range;
- cell content has exactly one valid discriminated state;
- formula expressions and cached values remain distinct, and unknown functions
  are preserved without evaluation;
- selection returns correct values when required styles, names, shared strings,
  or shared-formula masters live outside emitted ranges;
- stale dimensions and row spans neither allocate a grid nor hide cells;
- date and display output is identical under different host locales and time
  zones and contains no JavaScript `Date` objects;
- no external relationship is fetched;
- no raw XML, relationship ID, ZIP object, cache, or internal error escapes;
- strict and tolerant outcomes match the documented matrix;
- diagnostics remain bounded and do not echo sensitive/unbounded document text;
- browser tests revoke successful object URLs, and failed parses prove all
  parser-owned object URLs were already revoked.

### Property and fuzz coverage

Use recorded seeds and bounded run counts for:

- ZIP byte mutations, percent/path canonicalization, and duplicate/unsafe entry
  names;
- XML element, attribute, encoding, namespace, Markup Compatibility, and nesting
  mutations;
- streaming chunk splits across BOMs, multibyte characters, surrogate pairs,
  entities, comments, CDATA, attributes, and tags;
- relationship owner/target/type combinations;
- A1 cell and range parsing at `A1`, `XFD1048576`, and each one-over boundary;
- quoted sheet names, escaped apostrophes, 3D ranges, structured references,
  and malformed reference tokens;
- numeric, boolean, error, date serial, row, column, style, and shared-string
  lexical attributes;
- shared formula translation across absolute, relative, mixed, string-literal,
  named, structured, sheet-qualified, 3D, and external references;
- overlapping merge, selection, validation, and conditional-format ranges;
- style precedence and number-format sections;
- output accounting for cells, text, formulas, ranges, and pivot records.

Every minimized counterexample becomes a deterministic regression fixture.

### Mutation testing

Add XLSX targets incrementally to `scripts/mutation-scope.mjs`. A changed target
must add no `Survived` or `NoCoverage` mutants and must not lower the current
100% thresholds.

High-value mutation targets include:

- cell and range bounds;
- relationship ownership and external-target checks;
- shared formula translation;
- date-system boundaries and serial 60;
- style and column-range precedence;
- formula/cached-value discriminants;
- token-aware formula shifting and unknown-function preservation;
- limit accounting and fatal behavior;
- safe hyperlink and external-data redaction;
- worksheet selection and sparse output;
- duplicate cell, range, and relationship detection.
- failed-parse media cleanup and ownership transfer.

Kill mutants through public semantic assertions. Do not exclude meaningful
XLSX code merely to preserve the score.

### Scale and performance gates

Correctness limits need evidence that the implementation does not retain or
scan data quadratically. Add generated, ignored scale fixtures for:

- a very sparse worksheet whose used cells are near the grid boundaries;
- a dense worksheet close to the default returned-cell limit;
- many rows with only one selected range near the end;
- a large shared-string table with repeated and unique rich text;
- many styles, formulas, merges, validations, and conditional-format ranges;
- drawings/media near the byte budget;
- pivot cache records near the explicit record limit;
- an intentional one-over-limit stream that proves decompression and XML event
  delivery stop immediately.

Record elapsed time, peak heap/RSS, parsed events, scanned cells, returned
cells, and expanded bytes on a pinned CI runtime. Freeze regression thresholds
after the first representative implementation baseline; do not invent them from
the plan. The required asymptotic properties are fixed now:

- memory is proportional to bounded returned state plus explicitly retained
  dependency tables, never to the theoretical worksheet rectangle;
- range membership is indexed rather than `cells × selectedRanges`;
- each XML event and relationship is processed a bounded number of times;
- limit failure stops further decompression, XML parsing, media creation, and
  public-model construction;
- repeated and concurrent parses do not grow process-global caches.

Run the smaller scale guard in pull-request CI and the full memory/time matrix
in the reliability workflow. An outer worker/process timeout and memory cap
remain required for public uploads even when internal limits pass.

### Runtime dependency policy

Prefer the existing ZIP, SAX/XML, color, and runtime-neutral primitives. A new
formula tokenizer, number-format engine, or spreadsheet dependency is not
accepted merely to accelerate feature count. Before adding any runtime package,
document and test:

- why focused implementation or an existing dependency is insufficient;
- browser compatibility and absence of hidden Node.js APIs;
- ESM/CommonJS behavior and tree-shaking from the `./xlsx` subpath;
- minified/gzipped bundle-size change and parse-start cost;
- maintenance activity, known security posture, license, and transitive graph;
- whether it evaluates formulas, loads external data, or introduces mutable
  process-global state;
- independent tests that prevent the dependency from becoming the correctness
  oracle.

No dependency may fetch network data, execute formulas, or weaken XML/OPC
validation inside the format core.

## Build, package, CLI, browser, and CI

When the core reader contract is stable:

- export `parseXlsx`, `parseXlsxWithDiagnostics`, `XlsxParseError`, and public
  types from `src/formats/xlsx/index.ts` and the root API;
- add `@evoelsewhere/oakit/xlsx` to `package.json` exports;
- add `xlsx/index` to `tsup.config.ts`;
- extend package smoke tests for ESM, CommonJS, declarations, and subpath
  imports;
- extend CLI format inference and `--format` to `pptx | xlsx` through a format
  dispatch table rather than XLSX conditionals scattered through the PPTX
  adapter;
- retain a single JSON envelope with `format: "xlsx"`, document, and
  diagnostics;
- add CLI `--sheet`/`--range` selection without changing the programmatic
  selection semantics; stdin still requires `--format xlsx`;
- keep conservative CLI defaults (`imageMode: "none"`, no pivot records, and a
  bounded returned-cell count) and reject serialized output beyond a documented
  CLI byte budget rather than exhausting stdout/agent context;
- keep the format core browser-neutral and test public `Blob` input in
  Chromium;
- prove the `./xlsx` subpath does not pull the PPTX orchestrator or Node-only CLI
  adapters into the browser chunk, and enforce a reviewed bundle-size baseline;
- add Excel/Calc/Sheets corpus jobs without weakening the existing PPTX job;
- include XLSX mutation targets in the reliability workflow and upload the
  combined report;
- run the small scale guard in pull-request CI and the full scale/memory suite
  in the scheduled reliability workflow.

README support tables and examples change from “Planned” only after all public
exports and required gates pass.

## Delivery sequence and commit slicing

Each numbered item is a review boundary, not one large commit. Tests and
implementation remain separate commits when the regression evidence is useful
on its own.

### 1. Freeze the baseline contract

- add this plan and proposed decisions;
- add compile-only public type sketches only after review;
- freeze the discriminated cell-content and selected-sheet failure behavior;
- define the first diagnostic and limit contracts;
- add the minimal independent XLSX package fixture.

Gate: documentation format plus fixture self-inspection.

### 2. Extract proven common package primitives

- add shared content-type parsing tests;
- add shared owner-scoped relationship parsing tests;
- add canonical OPC part identity and duplicate-path tests;
- extract archive budget mechanics from PPTX-specific errors;
- retain PPTX adapters and behavior exactly;
- add shared streaming XML primitives with adversarial tests.

Gates: focused common tests, full PPTX tests, browser tests, mutation tests for
changed shared targets.

### 3. Establish the XLSX public skeleton

- add `XlsxParseError`, diagnostics, options, and limits;
- open packages and discover the workbook through package relationships;
- return ordered sheet metadata with no cell payload;
- prove input equivalence, immutability, concurrency, and strict/tolerant
  behavior.

Gate: XLSX black-box public API and package-validity suites.

### 4. Implement sparse cells and shared strings

- stream rows and cells;
- add A1 reference/range parsing;
- implement all cell value discriminants;
- support spec-valid omitted row/cell references and direct ISO date cells;
- preserve explicit blank styled cells;
- resolve plain, inline, and rich shared strings;
- enforce cell, text, row, column, and shared-string limits.

Gates: focused cell/reference/string tests, seeded property tests, browser
parse, mutation targets.

### 5. Implement styles, dates, and display semantics

- normalize style tables and precedence;
- add built-in/custom number formats;
- implement 1900/1904 systems and serial boundaries;
- add row/column defaults and explicit cell overrides;
- expose display text only when formatting is supported exactly and prove host
  locale/time-zone independence.

Gates: style/date matrix, Excel/Calc/Sheets core corpus, mutation targets.

### 6. Implement formulas and defined names

- preserve expressions and cached results;
- expand shared formulas with token-aware reference translation, including
  masters outside selected ranges;
- normalize array, data-table, and dynamic-array metadata;
- expose calculation properties and scoped names;
- prove no evaluation, fetching, or stale-cache claim occurs.

Gates: formula translation/property suite, adversarial external-reference tests,
mutation targets.

### 7. Implement worksheet structure

- rows, columns, hidden state, outline, views, panes, selections, merges;
- hyperlinks, notes, threaded comments, print settings, and protection;
- selection-aware output and range accounting.

Gates: focused structure tests and producer corpus.

### 8. Implement rules and table semantics

- SpreadsheetML tables;
- filters and sort state;
- data validation;
- conditional formatting and differential styles;
- sparklines and known extension namespaces.

Gates: public semantic tests, invalid/unsupported matrix, mutation targets,
cross-producer corpus.

### 9. Implement drawings, images, and charts

- worksheet anchors and transforms;
- explicit media modes and object URL ownership;
- revoke parser-owned media URLs on every later failure path;
- chart relationships, series formulas/caches, axes, titles, legends, and
  styles;
- extract shared DrawingML only where two-format evidence proves equivalence.

Gates: Node and browser media tests, chart corpus, PPTX regression suite,
mutation targets for shared changes.

### 10. Implement pivots and advanced metadata

- pivot definitions, fields, axes, filters, and styles;
- bounded cache metadata and opt-in records;
- slicers and timelines;
- external links, connections, and query-table metadata with redaction;
- security diagnostics for active and unsupported content.

Gates: advanced producer corpus, record-limit boundaries, no-fetch/redaction
tests, mutation targets.

### 11. Publish the integration surface

- add root and `./xlsx` package exports;
- update build and package smoke tests;
- extend CLI dispatch, selection, limits, help, and bounded output behavior;
- update README status, examples, limits, and security model;
- update `docs/architecture.md` with the implemented XLSX path;
- add Node 20/22/24, Chromium, corpus, and mutation CI coverage;
- establish bundle-size plus small/full performance baselines.

Gates: `pnpm check`, `pnpm test:browser`, all PPTX and XLSX corpus tiers,
complete mutation audit, and package consumer smoke tests.

## Definition of done

The XLSX reader may be called full-feature only when:

- every feature-matrix row marked required has a public normalized contract;
- every metadata and diagnostic row meets the explicitly weaker but complete
  commitment assigned in the matrix;
- all public types, implementation, README examples, architecture, CLI, and
  package exports agree;
- independent black-box packages cover required, optional, malformed, missing,
  unsupported, and security-rejected states;
- Excel, LibreOffice Calc, and Google Sheets corpus cases assert meaningful
  workbook semantics;
- exact numeric, cell, range, XML, ZIP, media, text, style, and pivot limits
  pass at-limit and reject one-over inputs;
- selected ranges remain semantically correct when dependencies are outside the
  emitted rectangle, and scanned/output work is independently bounded;
- strict and tolerant recovery are tested at the public boundary;
- Node 20/22/24 and Chromium pass;
- XLSX corpus, seeded fuzz/property tests, and mutation testing pass without
  new survivors or uncovered targets;
- no external content is fetched or executed;
- returned data is deterministic, sparse, bounded, JSON-compatible, and free
  of non-finite numbers, host-dependent dates, JavaScript `Date` objects, or
  internal parser state;
- streaming, selection indexing, early abort, peak memory, parse time, and
  browser bundle size remain within reviewed regression baselines;
- any new runtime dependency has passed the browser, bundle, maintenance,
  security, licensing, and necessity review;
- existing PPTX behavior and gates remain green;
- remaining unsupported valid extensions are listed explicitly with stable
  diagnostics rather than hidden behind a broad support claim.

## Main risks and mitigations

| Risk                                                      | Mitigation                                                                                              |
| --------------------------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| Huge worksheets exhaust memory even when ZIP limits pass. | Stream hot parts, sparse output, caller selection, cell/text/formula limits, outer process limits.      |
| Formula caches are stale or absent.                       | Separate formula and cached value; expose calculation properties; never recalculate or claim freshness. |
| Formula reference shifting corrupts names or literals.    | Token-aware translation with literal tests for every reference class; no broad replacement.             |
| Selection omits a dependency needed by a returned cell.   | Dependency-aware bounded reads and fixtures with masters/styles outside selected ranges.                |
| Style resolution becomes a second spreadsheet renderer.   | Normalize authored semantics, use explicit supported formatting, diagnose unsupported display tokens.   |
| Locale/time zone changes dates or display text.           | No JavaScript `Date`; explicit date systems and deterministic supported formatting only.                |
| Encoded/ambiguous OPC paths bypass ownership checks.      | Canonical part identities, post-decoding traversal checks, and duplicate normalized-name rejection.     |
| Shared code changes regress PPTX.                         | Extract one primitive at a time with unchanged PPTX public assertions and mutation coverage.            |
| “Full feature” expands without a measurable endpoint.     | Treat this matrix and definition of done as the support claim boundary.                                 |
| Advanced Excel extensions vary by version and producer.   | Namespace-specific fixtures, producer corpus, stable unsupported diagnostics, no raw XML escape.        |
| Connections leak secrets or trigger network activity.     | Never refresh; redact connection strings and credentials; no network capability in format core.         |
| Partial tolerant output misleads financial consumers.     | Reject malformed selected worksheet core data; recover only documented optional parts with diagnostics. |
| Browser object URLs leak.                                 | Opt-in blob mode, explicit caller ownership, test cleanup on success and failure.                       |
| Pivot caches dominate output.                             | Metadata default, records opt-in, strict record/text/byte limits.                                       |
| A broad dependency hides behavior or inflates browsers.   | Necessity/security/license review, independent tests, and package-subpath bundle budgets.               |

## Normative references

Implementation decisions should be checked against:

- ECMA-376 Office Open XML, especially Part 1 SpreadsheetML, Part 2 OPC, and
  Part 3 Markup Compatibility:
  <https://ecma-international.org/publications-and-standards/standards/ecma-376/>
- Microsoft SpreadsheetML extensions (`MS-XLSX`):
  <https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/>
- Microsoft Office implementation information for ECMA-376 (`MS-OE376`):
  <https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/>

Producer behavior may justify a documented compatibility path, but it must not
override package safety, XML correctness, relationship ownership, or public
value invariants.
