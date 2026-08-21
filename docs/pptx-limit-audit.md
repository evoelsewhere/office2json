# PowerPoint remaining-limit audit

This document is the authoritative inventory of meaningful PPTX capability
limits after the native text, shape, image, table, group, and common-chart
profiles. It separates product gaps from deliberate security boundaries and
from claims that require producer evidence rather than more code.

Audit baseline: 2026-08-21. A row is complete only when its public API,
independent OOXML test, strict reparse, Office-free render, package-preservation
assertions, browser coverage where applicable, mutation coverage, CI, and
documentation all agree.

## Status vocabulary

| Status                | Meaning                                                                         |
| --------------------- | ------------------------------------------------------------------------------- |
| Complete              | Implemented and certified by the current required gates                         |
| Certification pending | Implemented locally; full CI/Reliability evidence is not yet attached           |
| Planned native work   | A bounded native operation can be added without changing the security model     |
| Preservation boundary | Source content is copied exactly but semantic mutation is intentionally blocked |
| Explicit non-goal     | The project must not imply this capability                                      |

## Current native baseline

| Domain       | Source-free creation                                      | Source-preserving edit                                      | Current status |
| ------------ | --------------------------------------------------------- | ----------------------------------------------------------- | -------------- |
| Text         | Structured bodies, paragraphs, runs, fields, and breaks   | Plain-run replacement and frame transform                   | Complete       |
| Shapes       | Rect, roundRect, ellipse, fill, line, and transform       | Empty-shape transform                                       | Complete       |
| Images       | Signature-checked PNG/JPEG media and native `p:pic`       | Picture-frame transform; media bytes remain exact           | Complete       |
| Tables       | Grid, cells, text, fills, borders, and rectangular merges | Frame/grid transform and single plain cell-run replacement  | Complete       |
| Groups       | Recursive native children and explicit child space        | Nested group/descendant transforms and nested plain text    | Complete       |
| Charts       | Bar, line, pie, and doughnut with literal ChartML caches  | Chart-frame transform; ChartML remains byte-exact           | Complete       |
| Preservation | N/A                                                       | No-op byte-exact `R0`; untouched dirty-neighbor parts exact | Complete       |

The table-cell operation reuses the integrity-bound `replace-text` contract.
Targets use keys such as
`slide-1-element-2-row-3-cell-1-run-1`. Text nested inside native groups uses
the complete owner path, for example `slide-1-element-2-element-1-run-1`.
Both operations patch only the owning slide XML and reject fields, breaks,
multiple source text nodes, stale preconditions, unsafe indexes, extensions,
and Markup Compatibility branches.

Certification evidence for this capability row:

- [CI run 32450108185](https://github.com/evoelsewhere/oakit/actions/runs/32450108185):
  29/29 jobs passed across Node 20/22/24, Chromium/Firefox/WebKit, package
  smoke, and focused mutation modules.
- [Reliability run 32450510762](https://github.com/evoelsewhere/oakit/actions/runs/32450510762):
  60/60 jobs passed across fuzzing, corpus, render/memory budgets, full-file,
  shape, and static-inclusive mutation.
- The merged mutation artifact covers 108 files and 24,318 mutants: 18,774
  killed, 5,544 compile errors, and zero missed.
- The local full gate passed 169 test files / 3,702 tests, three seeded fuzz
  tests, ESM/CJS/declaration builds, and packed-package smoke.

## Remaining native feature gaps

### Near-term bounded operations

| Gap                   | Current evidence                                                                            | Required native result                                                                                                      | Completion evidence                                                                                                    |
| --------------------- | ------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| Shape-owned rich text | A shape with non-empty content becomes preservation-only in the round-trip preview          | Represent shape geometry and structured text together; replace bounded text without losing shape semantics                  | Independent shape-with-text fixture, style/geometry preservation, nested-group case, browser test, mutation audit      |
| Image crop            | Crop is parsed for rendering but is not part of the editable scene contract                 | Add normalized crop metadata and `setPptxRoundTripImageCrop`; patch only `a:srcRect` in the owning `p:pic`                  | Percent boundary tests, absent-vs-zero semantics, media byte equality, SVG crop verification                           |
| Image replacement     | Existing image media is copied and frame transforms are editable                            | Replace one supported PNG/JPEG relationship target with bounded bytes while preserving or explicitly allocating media parts | Signature/hash tests, relationship/content-type graph, shared-media policy, orphan prevention                          |
| Table structure       | Cell plain text and frame/grid sizing are editable; merges are creation-only                | Edit cell body/properties and merge/unmerge a rectangular region                                                            | Literal table XML assertions, continuation-cell validation, exact row/column invariants, producer reopen               |
| Chart data            | Native ChartML caches are created but preserved during edits                                | Replace bounded categories/values for literal-cache charts without formulas or external data                                | Cache cardinality/format assertions, chart-part-only dirty closure, visual verification, unsupported-formula rejection |
| Chart formatting      | Series color and selected options are creation-only                                         | Patch only explicitly modeled series/plot options and preserve unknown ChartML                                              | Exact ChartML owner rules, opaque-extension conflict tests, chart corpus                                               |
| Advanced charts       | Scatter, bubble, area, radar, stock, surface, and 3D variants are read or preservation-only | Add one chart family at a time with a versioned creation/edit capability row                                                | Independent ChartML fixtures, numeric/property fuzzing, PowerPoint/LibreOffice/Google Slides evidence                  |

### Presentation structure and authored owners

| Gap                                | Required policy before implementation                                                                                                                      |
| ---------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Add/duplicate/delete/reorder slide | Update presentation manifest and relationships; define notes, comments, chart, media, hyperlink, custom-show, section, zoom, and opaque-reference behavior |
| Slide background/hidden/layout     | Preserve authored owner and placeholder inheritance; reject ambiguous layout rebinding                                                                     |
| Speaker notes                      | Model notes slides separately from notes masters; preserve backlinks and notes-owned media                                                                 |
| Transitions                        | Model transition kind, duration, advance policy, and optional sound relationship                                                                           |
| Masters/layouts/themes             | Preserve master → layout → slide ownership and scheme-authored colors; never flatten resolved values into local RGB overrides                              |
| Accessibility metadata             | Model title, description, decorative state, language, and reading order in literal nonvisual properties                                                    |

These operations are not safe as blind slide-XML rewrites. Each needs a
reference-graph impact rule and an opaque-dependency conflict policy before its
API can be exposed.

### Advanced semantic domains

| Domain                         | Current policy                                              | Why it remains bounded                                                                                     |
| ------------------------------ | ----------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------- |
| Audio/video                    | Read/render metadata; semantic creation/edit not advertised | Media, poster, timing, and producer playback relationships form one representation group                   |
| Animations                     | Preserve timing trees outside dirty conflicts               | Targets, triggers, conditions, and media timing cross-reference stable shape IDs                           |
| SmartArt                       | Parse/preview and preserve its multi-part graph             | Layout algorithms and diagram data/drawing/style/color parts cannot be reconstructed from flattened shapes |
| Office Math                    | Parse supported content and preserve source                 | LaTeX preview is not a reversible OMML source                                                              |
| Comments/sections/custom shows | Preservation-only                                           | Cross-part identities and slide references require dedicated models                                        |
| Embedded fonts                 | Read metadata and preserve source                           | Embedding rights, obfuscation, theme use, and fallback behavior need a separate policy                     |

## Fidelity limits that are not API gaps

### Producer elevation

Native shape/image/table/group/chart creation currently reports runtime `C2`,
and native transforms/text-owner edits report runtime `R2`. Elevating a profile
to effective `C3` or `R3` requires controlled PowerPoint, LibreOffice, and
Google Slides save/reopen evidence for that exact capability row. A green
runtime verifier alone cannot make the producer claim.

### Preview approximation

Office-free SVG/PNG rendering is deterministic and warning-bearing, but it is
not pixel-identical to every Office producer. Remaining fidelity variables
include unavailable fonts, complete table-cell typography, advanced ChartML,
SmartArt algorithms, media effects, math fallbacks, and producer-specific
extensions. The correct result is an explicit warning, not a silent fidelity
claim.

### Rich text provenance

The current portable scene collapses some parsed rich text to a single plain
run for bounded edits. Broader range editing requires preserving paragraph/run
ownership, formatting, fields, links, breaks, and authored absence. It must not
derive a reverse mapping from rendered HTML.

## Deliberate security and preservation boundaries

The following are not defects to “work around”:

- encrypted or password-protected packages reject before semantic editing;
- signatures are preserved only while no edit invalidates them; no edit may
  silently remove a signature;
- `.pptm`, VBA, ActiveX, OLE, and executable relationships are not created or
  mutated by the PPTX profile;
- external relationships are never fetched;
- unknown Markup Compatibility branches and extensions intersecting a dirty
  subtree block the operation;
- malformed, unsafe, over-budget, or ambiguous ownership fails closed;
- portable snapshot hashes and source Base64 are immutable edit boundaries.

## Explicit non-goals

- Reconstructing an arbitrary original deck byte-for-byte from normalized JSON.
- Claiming arbitrary PPTX creation or editing.
- Claiming pixel-identical rendering across machines and producer versions.
- Guessing unsupported OOXML, producer extensions, chart formulas, SmartArt
  algorithms, animation graphs, or active-content behavior.

## Execution order

The remaining native work should proceed in this order:

1. Certify table-cell and nested-group text editing on all gates.
2. Add image crop editing because it has one owner part, no media-byte change,
   and a narrow semantic surface.
3. Add literal-cache chart data editing with strict formula/external-data
   rejection.
4. Add bounded table structure/style operations.
5. Represent shape-owned rich text without losing geometry/style provenance.
6. Add image replacement with explicit shared-media allocation policy.
7. Add slide manifest operations and speaker notes after the complete
   reference graph is available.
8. Add advanced chart, transition, accessibility, and producer-elevated rows
   independently.

Every completed row must keep the strict mutation threshold at 100%, add its
source file to mutation scope, pass Node 20/22/24 and Chromium/Firefox/WebKit,
strictly reparse and render its output, and prove untouched package payloads
remain byte-exact.
