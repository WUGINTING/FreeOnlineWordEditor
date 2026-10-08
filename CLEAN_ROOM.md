# Clean-room provenance record

This project was written from scratch. It is **not** derived from the source code
of SuperDoc (AGPL-3.0) or any other copyleft DOCX editor.

## How it was written
The code was written by the author with an AI coding assistant (Anthropic's Claude), which did
much of the writing, working from the materials listed below. What an AI model learned from
before this project cannot be audited by the author; this record covers what was consulted
during the work.

## What was consulted
- The public README / product description of SuperDoc, only to learn its
  *user-facing feature list* (open, edit, save DOCX in the browser; paginated view).
  No SuperDoc source files were opened, read, or copied.
- ECMA-376 / ISO/IEC 29500 (Office Open XML) — the public file-format standard.
- Public documentation of the permissively licensed dependencies listed below.

## Dependencies and their licenses
| Package | License |
|---|---|
| vue | MIT |
| prosemirror-* (model, state, view, transform, commands, keymap, history, tables, dropcursor, gapcursor) | MIT |
| jszip | dual MIT / GPL-3.0 — used under **MIT** |
| vite, vitest, typescript, vue-tsc, jsdom (dev only) | MIT / Apache-2.0 |

## Log
- 2026-09-23: project started; architecture chosen independently
  (ProseMirror document model + own OOXML reader/writer + measured pagination).
- 2026-09-23: header/footer editing; lossless round-trip redesign (original property XML
  patched in place, wrappers kept as layers). Designed from ECMA-376 and verified against
  Microsoft Word; no third-party editor source consulted.
