<div align="center">

# officelens

> Check DOCX and PPTX. Gate the build.

**Offline accessibility audit for DOCX and PPTX. Deterministic rules, JSON and SARIF 2.1.0 output, and CI exit codes.**

[![CI](https://github.com/srivtx/officelens/actions/workflows/ci.yml/badge.svg)](https://github.com/srivtx/officelens/actions/workflows/ci.yml)
[![release](https://img.shields.io/github/v/release/srivtx/officelens?sort=semver&color=4f46e5)](https://github.com/srivtx/officelens/releases)
[![license](https://img.shields.io/badge/license-MIT-0f766e)](LICENSE)
[![runtime](https://img.shields.io/badge/runtime-Bun-14151A?logo=bun&logoColor=white)](https://bun.sh)
[![types](https://img.shields.io/badge/TypeScript-strict-3178C6?logo=typescript&logoColor=white)](tsconfig.json)
[![tests](https://img.shields.io/badge/tests-17-0f766e)](#testing)
[![network](https://img.shields.io/badge/network-none-0f766e)](#privacy)

</div>

---

**Live site:** [officelens](https://officelens-srivtx.vercel.app)  ·  **Playground:** [https://officelens-srivtx.vercel.app/#playground](https://officelens-srivtx.vercel.app/#playground)  ·  **Source:** [github.com/srivtx/officelens](https://github.com/srivtx/officelens)

## The problem

Most accessibility checkers target formats other than Office:

- PDF has veraPDF and PAC.
- EPUB has DAISY ACE.
- HTML has axe-core and Pa11y.

For `.docx` and `.pptx`, the main option is the **Microsoft Accessibility
Checker**, a GUI tool for Windows, macOS, and the web. It has no CLI and no
JSON, and it does not run on a Linux CI runner. Microsoft's own documentation
tells teams to run it alongside a manual review because it has blind spots.

The other alternatives fall short:

| Tool | Why it doesn't solve it |
|---|---|
| Microsoft Accessibility Checker | GUI only; not scriptable; not headless |
| [Accessr](https://www.accessr.net/) | Browser-only, `.docx` only, no CLI, no JSON |
| `ooxml-cli`, `xarsh/ooxml-validator`, `openxml-audit` | Schema/well-formedness only, with no WCAG semantics |
| `ez-a` | ~0 stars, Python GUI, PPTX alt text only |

There is no open-source, offline, cross-platform, scriptable auditor for OOXML
documents, so `officelens` fills that gap.

## Install

`officelens` is not published to npm. Install it from GitHub with the one-line script (requires [Bun](https://bun.sh)):

```bash
# One-line install (installs the `officelens` binary)
curl -fsSL https://raw.githubusercontent.com/srivtx/officelens/main/install.sh | sh

# Or run once, without installing
bunx github:srivtx/officelens report.docx

# Install globally
bun add -g github:srivtx/officelens
officelens report.docx

# Add to a project as a dev dependency
bun add -d github:srivtx/officelens
```

## Usage

```bash
# Audit one or more documents (human-readable)
officelens report.docx deck.pptx

# Machine-readable JSON for a pipeline
officelens report.docx --json

# Just the summary line
officelens report.docx --quiet
```

Exit code is `1` when any error-severity issue is found, `0` otherwise, and `2`
when called with no arguments.

### Library

```ts
import { audit, formatJson, openOoxml, auditDocx, auditPptx } from "officelens";

const result = audit(bytes, "report.docx");
// { file, kind: "docx", issues, counts: { error, warning, info } }
```

## Rules

### DOCX

| Code | Severity | WCAG | Check |
|---|---|---|---|
| DOCX-ALT-001 | error | 1.1.1 | Drawing (`wp:docPr`) without `descr` or `title` |
| DOCX-HEAD-002 | warning | 1.3.1 | Body text but no headings at all |
| DOCX-HEAD-003 | warning | 1.3.1 | Heading levels skip (e.g. H1 → H3) |
| DOCX-LANG-004 | warning | 3.1.1 | No document language (`w:lang`) |
| DOCX-TBL-005 | error | 1.3.1 | Table whose first row is not a header (`w:tblHeader`) |
| DOCX-LINK-006 | warning | 2.4.4 | Hyperlink whose visible text is a raw URL |

### PPTX

| Code | Severity | WCAG | Check |
|---|---|---|---|
| PPTX-ALT-001 | error | 1.1.1 | Picture/shape (`a:cNvPr`) without `descr` |
| PPTX-TITLE-002 | error | 1.3.1 | Slide without a title placeholder |
| PPTX-TBL-003 | warning | 1.3.1 | Table not marked with `firstRow="1"` |
| PPTX-LANG-004 | warning | 3.1.1 | No run language (`a:rPr@lang`) |

`OOXML-000` (info) is reported when a file cannot be opened as an OOXML
package.

## How it works

```
.docx/.pptx ──unzip (fflate)──▶ parts + _rels
                                  ├── resolve main document / slide rels
                                  └── parse XML (fast-xml-parser, attributes on)
                                        └── rule functions → Issue[]
                                              └── JSON / text report
```

No Office, no Java, no browser, no network. The package is opened directly and
the XML is read.

## CI

```yaml
- name: Accessibility gate
  run: officelens docs/*.docx docs/*.pptx
```

### SARIF and code scanning

Emit a [SARIF 2.1.0](https://docs.oasis-open.org/sarif/sarif/v2.1.0/sarif-v2.1.0.html)
report and upload it to GitHub code scanning:

```bash
officelens report.docx --sarif officelens.sarif
```

```yaml
permissions:
  security-events: write

steps:
  - name: Accessibility audit (SARIF)
    run: officelens docs/*.docx docs/*.pptx --sarif officelens.sarif

  - name: Upload SARIF
    if: always()
    uses: github/codeql-action/upload-sarif@v3
    with:
      sarif_file: officelens.sarif
```

`--fail-on` sets the severity that fails the build: `error` (default),
`warning`, `info`, or `none`. The report is written even when the audit exits
nonzero, so the upload step still runs.

## Development

Clone the repository and install its dependencies:

```bash
git clone https://github.com/srivtx/officelens
cd officelens
bun install
```

| Gate | Command |
|---|---|
| Tests | `bun test` — 17 tests across docx, pptx, sarif, and CLI |
| Types | `bunx tsc --noEmit` — clean under strict mode |
| Fixtures | `bun run make-fixtures` — writes good and bad DOCX/PPTX |
| Site bundle | `bun run build:site` — bundles `src/index.ts` into `site/assets/demo.js` |
| Site check | `bun run check:site` — verifies internal links, classes, and page structure |

Fixtures are built in-repo as real OOXML zips, so the tests exercise the same
package parsing path as production files.

Preview the landing page and playground:

```bash
python3 -m http.server 8000 --directory site
```

Then open <http://localhost:8000>. The playground audits dropped `.docx` and
`.pptx` files entirely client-side; nothing is uploaded and no network request
is made.

## Privacy

No network code, no telemetry. Documents never leave the machine.

## Limitations

- Static XML analysis: it reads the parts directly rather than rendering.
- Not every WCAG success criterion is machine-checkable; this covers the
  structural, high-signal failures that block assistive technology.
- Spreadsheets (`.xlsx`) are not covered yet.

## The suite

- **booklens** — EPUB accessibility audit and fix
- **officelens** — DOCX/PPTX accessibility audit *(this repo)*
- **odflens** — ODT/ODS/ODP accessibility audit
- **iconlens** — standalone SVG accessibility lint
- **waxseal** — detached Ed25519 seal for WACZ web archives

## License

[MIT](LICENSE).
