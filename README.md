<p align="center">
  <img src="docs/assets/oakit-logo.png" alt="OAKit logo" width="220" />
</p>

<h1 align="center">Office Agent Kit</h1>

<p align="center">
  <strong> Office documents become agent-ready knowledge. </strong>
</p>

OAKit gives agents and automation systems a reliable way to read, understand,
preserve, preview, and create supported PowerPoint content through a consistent
structured model, with Excel and Word on the same product path. It owns the
difficult OOXML work—ZIP packages, relationships, inheritance, units, media,
malformed input, and producer differences—so agent workflows can operate on
meaningful document data instead of raw XML.

> **Project status:** pre-stable (`0.0.x`). Implemented PowerPoint capabilities
> include bounded reading, producer-verified text-profile `C3` creation,
> byte-exact `R0` portable hand-offs, producer-verified `R3` plain-text and text
> transform editing, runtime-verified native shape/image/table/group/chart `C2`
> creation, native transform/image-crop plus table/nested-group plain-text `R2`
> editing, and Office-free SVG/PNG previews. Excel and Word remain product
> direction rather than completed APIs.

## Real-world PowerPoint evidence

OAKit is exercised against a transient corpus of 30 complex templates selected
from [SlidesMania](https://slidesmania.com/). The controlled gate downloads each
source from its SlidesMania page, strict-parses and renders every slide without
Office, imports and exports it through Google Slides, deletes the temporary
presentation, then repeats strict parsing and rendering on the exported PPTX.

<p align="center">
  <img src="docs/evidence/0.0.1/slidesmania/producer-audit.png" alt="Audit summary for 30 SlidesMania templates round-tripped through Google Slides" width="100%" />
</p>

Run [32045412714](https://github.com/evoelsewhere/oakit/actions/runs/32045412714)
verified 30 decks, 733 slides, and 9,285 source elements with 100% minimum text
and element retention. Attribution was preserved in every export and all 30
temporary Google presentations were deleted. The tracked
[evidence manifest](docs/evidence/0.0.1/slidesmania/evidence.json) contains only
source-page links, hashes, and metrics—no SlidesMania templates or rendered
reproductions, in accordance with the
[SlidesMania template license](https://slidesmania.com/copyright-and-legal-information/#license).

The remaining warnings are explicit and narrow: 6,733 output text runs may use
a substitute when their authored font is unavailable, while 199 tables retain
their structure and layout but lack complete cell typography metadata in the
portable model. The same audit reports zero approximation warnings for shapes,
fills, and media. This gate proves robust opening, semantic retention, producer
compatibility, and Office-free rendering; it does not claim pixel-identical
rendering across machines with different installed fonts.

## Documentation

- [PowerPoint usage guide](docs/pptx-usage-guide.md): detailed Node.js,
  browser, CLI, portable JSON, editing, creation, rendering, security, and agent
  workflow examples.
- [Architecture](docs/architecture.md): package boundaries, data flow,
  invariants, extension points, and reliability model.
- [PowerPoint round-trip plan](docs/pptx-roundtrip-plan.md): fidelity profiles,
  supported edit scope, and progression toward broader native support.
- [PowerPoint remaining-limit audit](docs/pptx-limit-audit.md): implemented
  boundaries, native feature gaps, deliberate security policies, and the
  evidence required to close each row.

## Why OAKit

Office files are not single documents internally. They are ZIP packages made
of interconnected XML parts, relationships, themes, layouts, media, charts,
and vendor-specific extensions. That representation is a poor tool boundary
for an AI agent.

OAKit turns those internals into bounded, deterministic application data that is
easier to:

- summarize, classify, index, and search;
- inspect slide structure and extract semantic content;
- build document-aware tools and agent actions;
- validate generated changes before writing a file;
- run consistently in Node.js and modern browsers;
- process untrusted uploads with explicit diagnostics and resource limits.

OAKit is model- and framework-neutral. It does not require a particular LLM,
agent runtime, tool-calling protocol, or vector database.

## Format support

| Format               | Read | Create                                            | Edit                                                 | Preserve      | Preview |
| -------------------- | ---- | ------------------------------------------------- | ---------------------------------------------------- | ------------- | ------- |
| PowerPoint (`.pptx`) | Yes  | Text C3 + native shape/image/table/group/chart C2 | Text R3 + native transforms/crop/table/group text R2 | Byte-exact R0 | SVG/PNG |
| Excel (`.xlsx`)      | No   | No                                                | No                                                   | No            | No      |
| Word (`.docx`)       | No   | No                                                | No                                                   | No            | No      |

The runtime reports `C2` after deterministic package construction, strict
reparse, semantic comparison, and Office-free rendering. The declared
`pptx-create-text-v1` capability is certified at effective `C3` by controlled
PowerPoint, LibreOffice, and Google Slides save/reopen evidence. Likewise, the
runtime's `R2` verification is certified at effective `R3` for the
`pptx-roundtrip-text-v1` plain-text and text-transform operations. `R0` means an
unchanged source package is restored byte for byte through runtime or portable
JSON. The `pptx-create-native-v1` and `pptx-roundtrip-native-v1` profiles add
runtime-verified non-text shape/image/table/group/chart creation and slide-owned,
including nested, shape/image/table/group/chart transforms plus bounded native
image crop, table-cell, and nested-group text edits; they do not yet claim
producer-elevated C3/R3. These profile levels
do not claim arbitrary PPTX editing, full
reconstruction from normalized JSON, or pixel-identical rendering.

The PowerPoint reader currently handles:

- slide order, size, backgrounds, layouts, masters, themes, and notes;
- rich text, paragraphs, lists, links, fonts, and text fitting;
- preset and custom shapes, connectors, fills, borders, shadows, and groups;
- images, cropping, filters, embedded audio, and embedded video;
- tables, charts, SmartArt diagrams, and Office Math;
- transitions, relationship resolution, diagnostics, and resource limits.

OOXML has a very large extension surface. Unsupported optional structures may
be omitted with a diagnostic rather than represented inaccurately.

## Installation

The target npm package is `@evoelsewhere/oakit`:

```bash
pnpm add @evoelsewhere/oakit
```

```bash
npm install @evoelsewhere/oakit
```

### Homebrew CLI

On macOS or Linux, install the command-line interface from the EvoElsewhere
tap after the first formula release:

```bash
brew install evoelsewhere/tap/oakit
```

The Homebrew formula installs the `oakit` executable. Use npm or pnpm when the
programmatic JavaScript API is required.

The npm release and Homebrew formula are distributed independently. When a
registry release is not available for the desired version, use the repository
directly for development.

## Command-line interface

The package installs the `oakit` executable for deterministic Office-to-JSON
workflows in terminals, scripts, CI jobs, and agent sandboxes. The current CLI
accepts PowerPoint (`.pptx`) input.

Install the published CLI globally from npm:

```bash
npm install --global @evoelsewhere/oakit
oakit --version
```

It can also be run without a global installation:

```bash
npx --package @evoelsewhere/oakit oakit deck.pptx --pretty
```

### Convert a file

JSON is written to stdout by default, making the command suitable for pipes:

```bash
oakit deck.pptx > deck.json
```

Use the explicit `convert` command and `--output` when writing a file directly:

```bash
oakit convert deck.pptx --output deck.json --pretty
```

Both command forms are equivalent. OAKit refuses to overwrite the input document
with JSON output.

### Render slides for an agent

Render one or more slides as PNG files plus a structured manifest:

```bash
oakit render deck.pptx \
  --output previews \
  --render-format png \
  --slides 1,3 \
  --scale 1
```

Use `--render-format svg` for self-contained vector output. The output directory
contains deterministic `slide-N.png` or `slide-N.svg` files and
`manifest.json`. The manifest records each filename, byte length, MIME type,
dimensions, source slide number, and approximation warnings so an agent can
inspect the preview without inferring metadata from filenames.

Rendering runs in-process and does not require Microsoft Office, LibreOffice,
Google Slides, a headless browser, or a conversion service. PNG rendering is a
Node.js CLI capability; SVG rendering uses the same browser-neutral renderer
exposed by the public API.

### Preserve a byte-exact agent hand-off

Create portable JSON that carries the strict semantic preview, integrity
metadata, and complete source package:

```bash
oakit snapshot deck.pptx --output deck.oakit.json --pretty
```

Restore it only after its JSON shape, canonical Base64, source hash, semantic
preview, and consistency metadata have been verified:

```bash
oakit restore deck.oakit.json --output restored.pptx
```

A successful no-op restore is an `R0` hand-off: `restored.pptx` is byte-for-byte
identical to `deck.pptx`, including package parts OAKit does not interpret.
The portable document preview is integrity-bound to the source. Editing that
preview directly is rejected; authorized changes are represented as bound,
hash-protected operations.

### Replace supported text through portable JSON

Inspect the snapshot's `document` for a stable run key, schedule an edit, then
restore the edited portable snapshot:

```bash
oakit edit-text deck.oakit.json \
  --target slide-1-element-1-run-1 \
  --value "Updated by an agent" \
  --output edited.oakit.json \
  --pretty

oakit restore edited.oakit.json --output edited.pptx
```

The portable JSON continues to carry the original PPTX bytes plus an operation
log; it never hides an already-edited package behind stale operations. Restore
checks the source hash and exact text precondition, patches only the owning
slide part, verifies all untouched payloads byte-for-byte, strict-parses the
result, and compares the full semantic preview. The current edit profile is
deliberately narrow: one plain DrawingML text node in a slide-owned text shape.
Fields, line breaks, multiple runs, compatibility extensions, signed packages,
and macro-enabled packages fail closed with a typed error.

Move, resize, rotate, or flip that text element with a partial transform. Fields
not provided are copied from the bound preview:

```bash
oakit transform-text edited.oakit.json \
  --target slide-1-element-1 \
  --x=-10 \
  --width 500 \
  --rotation 15 \
  --flip-horizontal true \
  --output transformed.oakit.json
```

Transform editing requires one simple DrawingML `a:xfrm` owned by the target
text shape. Group coordinate children and compatibility extensions fail closed.
PowerPoint preserves the declared values exactly in the native producer gate;
LibreOffice's controlled save/reopen gate allows at most 0.2 point of geometry
quantization while requiring rotation and flip state to remain exact.

All three commands run in-process without Office software or a conversion
service. `snapshot` and `edit-text` can write JSON to stdout. `restore` requires
an output file so binary data is never mixed with terminal text.

### Read from stdin

Use `-` as the input path and provide the format explicitly:

```bash
cat deck.pptx | oakit - --format pptx --document-only > deck.json
cat deck.pptx | oakit snapshot - --format pptx > deck.oakit.json
cat deck.oakit.json | oakit edit-text - --target slide-1-element-1-run-1 --value "Updated" > edited.oakit.json
cat edited.oakit.json | oakit restore - --output edited.pptx
```

`--format pptx` is required for stdin because there is no filename extension
from which to infer the format.

### CLI options

```text
Usage: oakit [convert] <input.pptx|-> [options]
       oakit render <input.pptx|-> --output <directory> [options]
       oakit snapshot <input.pptx|-> [--output <file>]
       oakit edit-text <input.json|-> --target <run-key> --value <text> [options]
       oakit transform-text <input.json|-> --target <element-key> [options]
       oakit restore <input.json|-> --output <file.pptx>

Convert options:
  -o, --output <file>          Write JSON to a file instead of stdout
      --strict                 Reject malformed optional OOXML content
      --pretty                 Format JSON with two-space indentation
      --document-only          Omit format metadata and diagnostics
      --image-mode <mode>      Image output: none (default) or base64

Render options:
  -o, --output <directory>     Write slide files and manifest.json
      --render-format <format> png (default) or svg
      --slides <list>          One-based comma-separated slide numbers
      --scale <number>         Positive decimal output scale (default: 1)

Snapshot options:
  -o, --output <file>          Write portable JSON instead of stdout
      --pretty                 Format portable JSON with two-space indentation

Edit text options:
  -o, --output <file>          Write edited portable JSON instead of stdout
      --target <run-key>       Stable text run key from the portable document
      --value <text>           Replacement text; use --value=-5 for leading -
      --pretty                 Format portable JSON with two-space indentation

Transform text options:
  -o, --output <file>          Write edited portable JSON instead of stdout
      --target <element-key>   Stable text element key from the portable document
      --x <number>             Set horizontal position; use --x=-10 when negative
      --y <number>             Set vertical position; use --y=-10 when negative
      --width <number>         Set positive width
      --height <number>        Set positive height
      --rotation <number>      Set rotation in degrees
      --flip-horizontal <bool> Set horizontal flip: true or false
      --flip-vertical <bool>   Set vertical flip: true or false
      --pretty                 Format portable JSON with two-space indentation

Restore options:
  -o, --output <file>          Required PowerPoint output path

PPTX input options:
      --format <pptx>          Input format; required when reading stdin
  -h, --help                   Show help
  -v, --version                Show the installed OAKit version
```

### Convert output

The default `convert` output is an envelope that keeps format and recovery
information available to automation:

```json
{
  "format": "pptx",
  "document": {
    "slides": []
  },
  "diagnostics": []
}
```

Use `--document-only` when a downstream tool accepts only the normalized
document model. Use `--strict` when partial recovery is not acceptable. Images
are omitted by default to keep agent context and pipeline output bounded;
enable `--image-mode base64` only when the binary representation is required.
Audio and video payloads are never emitted by the CLI.

### Errors and exit codes

Errors are written as single-line JSON to stderr without a stack trace:

```json
{
  "error": {
    "code": "unsupported-format",
    "message": "Unsupported Office format: docx"
  }
}
```

| Exit code | Meaning                                                |
| --------- | ------------------------------------------------------ |
| `0`       | Requested conversion, hand-off, render, or help passed |
| `1`       | Input, validation, rendering, or output failed         |
| `2`       | Invalid command-line usage                             |

The CLI processes one document per invocation. Resource-limit failures remain
fatal in tolerant mode, matching the programmatic API's security boundary.

## Quick start

### Node.js

```ts
import { readFile } from 'node:fs/promises';
import { parsePptxWithDiagnostics } from '@evoelsewhere/oakit';

const input = await readFile('./quarterly-review.pptx');
const { document, diagnostics } = await parsePptxWithDiagnostics(input, {
  imageMode: 'none',
  errorMode: 'tolerant',
});

console.log({
  slideCount: document.slides.length,
  size: document.size,
  fonts: document.usedFonts,
  diagnostics,
});
```

Node.js `Buffer` extends `Uint8Array`, so bytes returned by `readFile` can be
passed directly to OAKit.

### Browser

```ts
import { parsePptx } from '@evoelsewhere/oakit/pptx';

const picker = document.querySelector<HTMLInputElement>('#presentation');
const file = picker?.files?.[0];

if (file) {
  const presentation = await parsePptx(file, {
    imageMode: 'both',
    videoMode: 'blob',
    audioMode: 'blob',
  });

  console.log(presentation.slides);
}
```

### Agent tool boundary

Use the diagnostic API when an agent must distinguish usable partial output
from a clean parse:

```ts
import { parsePptxWithDiagnostics } from '@evoelsewhere/oakit';

export async function inspectPresentation(bytes: Uint8Array) {
  const result = await parsePptxWithDiagnostics(bytes, {
    imageMode: 'none',
    videoMode: 'none',
    audioMode: 'none',
    errorMode: 'tolerant',
  });

  return {
    kind: 'presentation' as const,
    document: result.document,
    diagnostics: result.diagnostics,
  };
}
```

Document text is untrusted content. An agent host must keep it in the data
portion of its prompt or tool result and must not treat instructions embedded
in a document as trusted system or developer instructions.

## Public PowerPoint API

Both entry points expose the same reader:

```ts
import {
  parsePptx,
  parsePptxWithDiagnostics,
  PptxParseError,
} from '@evoelsewhere/oakit';

import { parsePptx as parsePptxFormat } from '@evoelsewhere/oakit/pptx';
```

The format-specific entry point is preferred when an application only needs
PowerPoint support.

### Input

```ts
type PptxInput = ArrayBuffer | Uint8Array | Blob;
```

### Options

```ts
interface PptxParseOptions {
  imageMode?: 'base64' | 'blob' | 'both' | 'none';
  videoMode?: 'blob' | 'none';
  audioMode?: 'blob' | 'none';
  errorMode?: 'tolerant' | 'strict';
  limits?: PptxResourceLimits;
}
```

| Option      | Default       | Behavior                                            |
| ----------- | ------------- | --------------------------------------------------- |
| `imageMode` | `base64`      | Return data URLs, object URLs, both, or neither.    |
| `videoMode` | `none`        | Create object URLs for supported embedded video.    |
| `audioMode` | `none`        | Create object URLs for supported embedded audio.    |
| `errorMode` | `tolerant`    | Recover with diagnostics or reject malformed OOXML. |
| `limits`    | Safe defaults | Bound archive, XML, media, and slide processing.    |

### Output

```ts
interface PptxDocument {
  size: {
    width: number;
    height: number;
  };
  themeColors: string[];
  usedFonts: string[];
  slides: PptxSlide[];
}
```

Positions and dimensions use points. Each slide separates authored elements
from inherited layout and master content:

```ts
interface PptxSlide {
  fill: Fill;
  elements: PptxElement[];
  layoutElements: PptxElement[];
  note: string;
  transition?: SlideTransition | null;
}
```

Elements use a discriminated `type` field:

| `type`    | Content                                                       |
| --------- | ------------------------------------------------------------- |
| `text`    | Positioned rich-text HTML and text layout                     |
| `shape`   | Shape metadata and an SVG-compatible path when available      |
| `image`   | Package reference, selected representation, crop, and filters |
| `video`   | Embedded or external reference and optional object URL        |
| `audio`   | Embedded reference and optional object URL                    |
| `table`   | Cells, merges, dimensions, fills, and borders                 |
| `chart`   | Normalized series, labels, colors, and options                |
| `diagram` | SmartArt drawing elements and logical text                    |
| `math`    | Parsed LaTeX and an optional fallback image                   |
| `group`   | Nested elements in the group coordinate space                 |

Text is returned as an escaped HTML fragment. Applications that inject
document HTML into a page should still apply their own sanitizer as defense in
depth.

### Create bounded text, shape, image, table, and group presentations

`createPptx` accepts the versioned scene model, validates it strictly, and
returns deterministic package bytes with a `C2` runtime fidelity report. The
creation profiles support source-free slides containing structured text,
native rect/roundRect/ellipse shapes, signature-checked PNG/JPEG images, and
structured native tables with exact grids, cell text, fills, borders, and merges,
plus nested groups with explicit outer and child coordinate spaces; they reject
unsupported scene elements rather than emitting a guessed package.
The release capability matrix certifies the exact
`pptx-create-text-v1` profile at effective `C3`.

```ts
import { writeFile } from 'node:fs/promises';
import { createPptx, type PptxSceneDocument } from '@evoelsewhere/oakit';

const scene: PptxSceneDocument = {
  layouts: [],
  masters: [],
  media: [],
  schemaVersion: 2,
  size: { height: 540, width: 960 },
  slides: [
    {
      elements: [
        {
          authored: {
            transform: { height: 80, width: 600, x: 40, y: 40 },
          },
          key: 'title',
          resolved: { hidden: false },
          text: {
            body: { anchor: 'center', wrap: true },
            paragraphs: [
              {
                children: [
                  {
                    key: 'title-run',
                    properties: { bold: true, fontSize: 28 },
                    text: 'Agent-ready presentation',
                    type: 'run',
                  },
                ],
                key: 'title-paragraph',
              },
            ],
          },
          type: 'text',
        },
      ],
      key: 'slide-1',
    },
  ],
  themes: [],
};

const created = await createPptx(scene);
await writeFile('created.pptx', created.data);
console.log(created.report.level); // C2
```

### Agent-ready slide previews without Office

OAKit can turn parsed slides into self-contained SVG in Node.js or a browser.
Node.js callers can also create PNG bytes through the dedicated Node entry
point. Neither path launches or requires Microsoft Office, LibreOffice, Google
Slides, a headless browser, or a network service.

```ts
import { readFile } from 'node:fs/promises';
import { renderPptxToSvg } from '@evoelsewhere/oakit';
import { renderPptxToPng } from '@evoelsewhere/oakit/pptx/node';

const input = await readFile('./deck.pptx');
const svg = await renderPptxToSvg(input, {
  scale: 1,
  slideNumbers: [1],
});
const png = await renderPptxToPng(input, {
  scale: 1,
  slideNumbers: [1],
});

console.log(svg.slides[0]?.data, png.slides[0]?.data);
```

Every rendered slide includes bytes, MIME type, dimensions, its one-based slide
number, and structured warnings for visual approximations. SVG output escapes
document text, embeds only validated raster data URLs, and never follows an
external reference. PNG rasterization uses the generated self-contained SVG.

For a JSON-safe, byte-exact agent hand-off, use the separate round-trip API:

```ts
import {
  parsePptxRoundTripJson,
  readPptxRoundTrip,
  replacePptxRoundTripText,
  setPptxRoundTripGroupTransform,
  setPptxRoundTripImageTransform,
  setPptxRoundTripShapeTransform,
  setPptxRoundTripTableTransform,
  setPptxRoundTripTextTransform,
  serializePptxRoundTripJson,
  writePptxRoundTrip,
} from '@evoelsewhere/oakit';

const runtime = await readPptxRoundTrip(input);
const textEdited = await replacePptxRoundTripText(runtime, {
  targetKey: 'slide-1-element-1-run-1',
  value: 'Updated by an agent',
});
const edited = await setPptxRoundTripTextTransform(textEdited, {
  targetKey: 'slide-1-element-1',
  value: {
    flipHorizontal: false,
    flipVertical: false,
    height: 80,
    rotation: 15,
    width: 500,
    x: 40,
    y: 40,
  },
});
const portable = await serializePptxRoundTripJson(edited);
const wireValue: unknown = JSON.parse(JSON.stringify(portable));
const restored = await parsePptxRoundTripJson(wireValue);
const output = await writePptxRoundTrip(restored);

console.log(output.report.level); // R2
```

The portable envelope validates its exact shape, canonical Base64, source
digest, package inventory, semantic preview, operation log, and consistency
hashes. Changing preview fields or operation values directly is detected as
tampering rather than treated as an edit.

Rendering is a portable visual aid for agents, not a claim of pixel-identical
PowerPoint layout. Exact `R0` package preservation and preview rendering are
separate operations: a preview never changes the source package used by the
portable JSON round trip.

Render limits are validated before work begins or immediately after encoding
the bounded result:

| Render limit        |    Default |
| ------------------- | ---------: |
| Slides per request  |      1,000 |
| Elements per slide  |     10,000 |
| Scale               |          8 |
| Pixels per slide    | 32 MiPixel |
| SVG bytes per slide |    128 MiB |
| PNG bytes per slide |    256 MiB |

Callers processing public uploads should lower these limits to their own
latency and memory budgets.

## Diagnostics and strict mode

`parsePptx` returns the document directly. `parsePptxWithDiagnostics` returns:

```ts
interface PptxParseResult {
  document: PptxDocument;
  diagnostics: PptxDiagnostic[];
}
```

Tolerant mode may omit malformed optional content while recording a structured
diagnostic. Strict mode rejects malformed XML, unsafe relationships, invalid
values, and missing required parts with `PptxParseError`.

Resource-limit violations always reject, including in tolerant mode. They
represent a security boundary, not a recoverable fidelity problem.

## Security model

OAKit treats every uploaded package as untrusted input. The reader:

- rejects unsafe package and relationship paths;
- rejects malformed XML structures and forbidden declarations;
- validates numeric values before conversion;
- does not execute macros, scripts, media, or hyperlinks;
- does not fetch external relationships;
- filters supported hyperlink protocols;
- bounds compressed input, ZIP entries, expanded bytes, XML complexity,
  embedded media, and slide count.

Default limits:

| Limit                     |   Default |
| ------------------------- | --------: |
| Compressed input          |   100 MiB |
| Non-directory ZIP entries |    10,000 |
| Total declared expansion  |   256 MiB |
| One expanded package part |    64 MiB |
| One expanded XML part     |    16 MiB |
| XML nesting depth         |       128 |
| XML elements per part     |   250,000 |
| XML elements per package  | 1,000,000 |
| One expanded media part   |    64 MiB |
| Slides                    |     1,000 |

For public uploads, also isolate parsing in a worker or process and enforce an
outer timeout and memory limit.

## Media lifecycle

When a blob mode is enabled, OAKit calls `URL.createObjectURL`. The caller owns
the returned URLs and must release them:

```ts
URL.revokeObjectURL(element.blob);
```

Remember to traverse nested group and diagram elements when releasing media.

## Runtime support

- Node.js 20, 22, and 24;
- modern browsers with `Blob` and `URL.createObjectURL` support;
- ESM and CommonJS;
- declarations and source maps.

## Development

```bash
pnpm install
pnpm check
```

| Command                                      | Purpose                                                 |
| -------------------------------------------- | ------------------------------------------------------- |
| `pnpm dev`                                   | Rebuild in watch mode                                   |
| `pnpm test`                                  | Run deterministic unit, integration, and property tests |
| `pnpm test:browser`                          | Run the public API suite in Chromium                    |
| `pnpm test:corpus`                           | Verify PowerPoint and LibreOffice documents             |
| `pnpm test:corpus:large`                     | Include the large Google Slides corpus                  |
| `pnpm test:fuzz`                             | Fuzz SVG/PNG safety using reproducible recorded seeds   |
| `pnpm test:mutation`                         | Measure whether tests detect behavioral mutations       |
| `pnpm test:mutation:module -- relationships` | Run one incremental focused mutation module             |
| `pnpm test:package`                          | Smoke-test package exports and the bundled CLI          |
| `pnpm test:render:e2e`                       | Prove Office-free portable edit and SVG/PNG rendering   |
| `pnpm test:render:memory`                    | Measure SVG/PNG memory at 1, 25, and 100 slides         |
| `pnpm test:producer:powerpoint`              | Run macOS PowerPoint save/reopen evidence               |
| `pnpm test:producer:google-slides`           | Run controlled Google Slides import/export evidence     |
| `pnpm typecheck`                             | Run strict type checking                                |
| `pnpm lint`                                  | Run ESLint                                              |
| `pnpm format:check`                          | Verify formatting                                       |
| `pnpm build`                                 | Build ESM, CommonJS, declarations, and source maps      |
| `pnpm check`                                 | Run the required pull-request quality gates             |

The fast CI matrix runs on Node.js 20, 22, and 24 plus Chromium. Producer
corpus and mutation suites run in the reliability workflow. Pull requests run
the seven patch modules independently with focused tests and incremental
caches. Release Reliability forces every mutant, including static mutants, and
balances general file shards from recorded elapsed Stryker time.

Read [docs/architecture.md](docs/architecture.md) before changing parser
ownership, public models, resource handling, or format boundaries. Development
rules for coding agents and contributors live in [AGENTS.md](AGENTS.md).

## Roadmap

- stabilize the normalized PowerPoint model;
- expand real-producer fidelity and adversarial corpus coverage;
- expand mutation-tested, part-preserving PowerPoint operations;
- introduce Excel and Word as isolated format domains;
- expand higher-level document operations suitable for agent tools;
- keep the core independent of model vendors and agent frameworks.

Capabilities are documented only after their public API, implementation, and
tests exist.

## License

[MIT](LICENSE) © 2026 EvoElsewhere.
