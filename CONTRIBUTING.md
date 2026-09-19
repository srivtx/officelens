# Contributing to officelens

Thanks for helping improve OOXML accessibility tooling. This document covers
what you need to build, test, and submit a change.

## Development setup

officelens targets [Bun](https://bun.sh) and TypeScript in strict mode.

```bash
git clone https://github.com/srivtx/officelens.git
cd officelens
bun install
```

`fixtures/` is generated and gitignored, so build it before running anything
that reads a sample file:

```bash
bun run make-fixtures
```

Run the CLI from source while you work:

```bash
bun run src/cli.ts fixtures/bad.docx
```

## The gate

Every pull request must pass the same gate CI runs:

```bash
bunx tsc --noEmit && bun test
```

Do not open a PR with a red typecheck or a failing test. Fix the cause rather
than disabling a rule or test.

## Fixtures

`fixtures/` holds the OOXML packages the tests run against. They are generated,
not hand-edited:

```bash
bun run make-fixtures
```

`good.docx`, `good.pptx`, `bad.docx`, and `bad.pptx` cover the DOCX and PPTX
paths. When you add a rule, add a fixture part that exercises it and assert on
the emitted issue codes in `tests/docx.test.ts` or `tests/pptx.test.ts`. CLI
behavior belongs in `tests/cli.test.ts`.

## Code style

- Strict TypeScript. Prefer precise types over `any`, and avoid non-null
  assertions to dodge null checks. The recursive XML walkers (`findAll`,
  `collect`) operate on untyped parser output, so `any` and a rare `!` are
  acceptable there and only there.
- No new runtime dependencies without discussion in an issue first. The offline
  and dependency-light posture is a feature.
- No network access, ever. Unzipping and parsing are local operations.
- Keep the DOCX and PPTX walkers separate and explicit; do not share state
  between document modes.
- Match the surrounding style; keep modules small and focused.
- No comments unless they explain something non-obvious.

## Commit messages

Use [Conventional Commits](https://www.conventionalcommits.org/):

```
<type>(<scope>): <summary>

fix(docx): flag images without alt text in headers
feat(pptx): report reading order gaps per slide
test(cli): cover --json exit codes
docs: document the JSON report schema
```

Common types: `feat`, `fix`, `test`, `docs`, `refactor`, `chore`. Useful scopes:
`docx`, `pptx`, `cli`, `ooxml`.

## Pull request checklist

- [ ] Tests added or updated for the change.
- [ ] `bunx tsc --noEmit` is clean.
- [ ] `bun test` passes.
- [ ] Docs (`README.md`) updated when behavior or flags change.
- [ ] Commit messages follow Conventional Commits.
