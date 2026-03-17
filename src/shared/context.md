# Shared Action Layer Context

`src/shared/` contains the main behavior layer.

## Files

- `shapeTools.ts`: large Office.js action module for layout, formatting, builders, and slide ops.
- `presentationTools.ts`: presentation-level QA/export helpers for comments, speaker notes, and selected-slide deck generation.

## Core Contract

- User-facing actions return `Promise<ActionResult>`.
- `ActionResult` is `{ type, message }`.
- `type` is one of `success`, `info`, `warning`, or `error`.

This keeps ribbon commands and taskpane status handling consistent.

## Main Domains

- position and size copy/paste
- alignment and grouped-alignment helpers
- style transfer and text transforms
- grid / gantt builders
- slide organization and presentation helpers
- PPTX export/scrub workflows that operate on the compressed file package when Office.js lacks a direct API

## State and Assumptions

- Position/size copy-paste uses in-memory module state.
- Several layout/build flows still assume a `960x540` slide model.
- PowerPoint API gaps are handled with warnings instead of silent failures where practical.
- Presentation QA/export flows use `Office.context.document.getFileAsync(Office.FileType.Compressed)` and mutate the PPTX package in-browser.
- Local mode can optionally hand generated `.pptx` files to the localhost bridge for native save/open/email actions.

## Editing Guidance

1. Preserve `ActionResult` shape.
2. Keep manifest-bound action names stable.
3. Be deliberate with `load(...)` / `context.sync()` placement.
4. Re-test selection-order-sensitive flows after edits.
5. For presentation-package edits, prefer removing package relationships first; orphaned unused parts are safer than dangling references.
