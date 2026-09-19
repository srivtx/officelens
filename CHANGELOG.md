# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-09-19

### Added

- Initial release of `officelens`, an offline DOCX/PPTX accessibility auditor.
- `officelens` CLI with `--json`, `--quiet`, `--sarif <path>`, and
  `--fail-on <level>` options plus CI-friendly exit codes.
- DOCX rules: `DOCX-ALT-001`, `DOCX-HEAD-002`, `DOCX-HEAD-003`,
  `DOCX-TBL-005`, `DOCX-LINK-006`, `DOCX-LANG-004`.
- PPTX rules: `PPTX-ALT-001`, `PPTX-TITLE-002`, `PPTX-TBL-003`,
  `PPTX-LANG-004`.
- SARIF 2.1.0 output via `toSarif` and `writeSarif` for GitHub code scanning.
- Fixture generator (`scripts/make-fixtures.ts`) producing accessible and
  inaccessible DOCX and PPTX samples.
- Library entry point re-exporting `audit`, `formatText`, `formatJson`,
  `openOoxml`, `auditDocx`, and `auditPptx`.

[0.1.0]: https://github.com/srivtx/officelens/releases/tag/v0.1.0
