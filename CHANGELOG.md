# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-09-19

### Added

- Initial release of `ooxml-a11y`, an offline DOCX/PPTX accessibility auditor.
- `ooxml-a11y` CLI with `--json` and `--quiet` output modes and CI-friendly
  exit codes.
- DOCX rules: `DOCX-ALT-001`, `DOCX-HEAD-002`, `DOCX-HEAD-003`,
  `DOCX-TBL-005`, `DOCX-LINK-006`, `DOCX-LANG-004`.
- PPTX rules: `PPTX-ALT-001`, `PPTX-TITLE-002`, `PPTX-TBL-003`,
  `PPTX-LANG-004`.
- Fixture generator (`scripts/make-fixtures.ts`) producing accessible and
  inaccessible DOCX and PPTX samples.
- Library entry point re-exporting `audit`, `formatText`, `formatJson`,
  `openOoxml`, `auditDocx`, and `auditPptx`.

[0.1.0]: https://github.com/ooxml-a11y/ooxml-a11y/releases/tag/v0.1.0
