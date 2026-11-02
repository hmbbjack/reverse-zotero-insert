# Changelog

[Project home](../README.en.md) · [简体中文](CHANGELOG.md)

- **Unreleased** — Share Word field parsing across detection, conversion, and auditing.
  Preserve internal bookmarks, comment markers, and proofing markers during repeated conversion,
  including fields spanning paragraphs. Reject nested Zotero citations in verification and CLI audits;
  add regression coverage for simple fields, nested fields, and text boxes.

- **v3.1** — Fix lost tabs/line breaks, repeated-citation omission checks, audit exit codes,
  DOI truncation during ingestion, renumbering selection/collisions, and MCP request IDs/error handling.
  Consolidate shared logic; add pip packaging, ZIP downloads, and release automation.
- **v3** — Scan-first workflow, automatic citation ID offsets, legacy parenthesized DOI support,
  cross-paragraph field tracking, explicit library indexing failures, group reconciliation,
  table auditing, and pre-delivery checks.
- **v2** — Mixed-document support, preservation of existing fields, table DOI citations,
  and library metadata refresh.
- **v1** — Basic field generation, reference resolution, and confirmed Zotero imports.
