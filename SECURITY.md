# Security Policy

## officelens

officelens audits the accessibility of OOXML documents — Word (`.docx`) and
PowerPoint (`.pptx`). It opens an untrusted package, parses its XML parts, and
reports failures as JSON or human-readable output.

## Supported versions

The latest commit on `main` is the only supported version. Security fixes land
on `main` and ship in the next tagged release. Older tags do not receive
backports.

| Version | Supported |
| --- | --- |
| Latest on `main` | Yes |
| Older tags | No |

## Threat model

- **Offline by design.** officelens contains no network code. It never opens a
  socket, resolves external relationships, or checks for updates.
- **No telemetry.** Nothing about your documents, your usage, or your machine
  is collected or transmitted.
- **Documents never leave the machine.** Unzipping, XML parsing, and auditing
  all run in-process and locally.
- **Untrusted input.** A DOCX or PPTX is treated as hostile: ZIP members and
  XML parts are parsed defensively, and a malformed or oversized package fails
  safely rather than escaping the working directory or exhausting the process.
  The ZIP container is rejected before extraction when it declares more than
  65,535 members, more than 256 MiB in a single member, or more than 1 GiB
  uncompressed in total. A part that is present but not well-formed XML is
  reported as an `OOXML-000` error rather than silently skipped.
- **No code execution from input.** Macros, embedded objects, and external
  relationships in a document are never evaluated or followed.
- **Deterministic output.** The same document yields the same findings and exit
  code, with no dependency on external services.

## Reporting a vulnerability

Report privately through GitHub Security Advisories on the repository:

https://github.com/srivtx/officelens/security/advisories/new

Do not open a public issue for a suspected vulnerability. Include a
description, the affected revision, a minimal reproducer (a DOCX or PPTX
fixture where possible), and any suggested fix. Expect an acknowledgement
within a few days.

## Verifying a build

```bash
bun install
bunx tsc --noEmit
bun test
```

This installs the locked dependency set, typechecks in strict mode, and runs
the test suite against the generated fixtures. In CI the same gate runs on
every push and pull request.
