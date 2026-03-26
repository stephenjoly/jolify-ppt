# Dialogs Context

This folder contains static Office dialog pages used for structured user input.

## Current Dialogs

- `grid-builder.html`: grid input.
- `symbol-picker.html`: bounded symbol picker for text insertion.

## Integration Pattern

1. Shared code opens a dialog with `openDialog(...)`.
2. Dialog sends JSON back through `Office.context.ui.messageParent(...)`.
3. Parent parses the payload and returns an `ActionResult`.

## Contract Notes

- Cancel should send `JSON.stringify(null)`.
- Payload shape must stay aligned with the parsing code in `shapeTools.ts`.
- These pages are intentionally static and unbundled; keep them lightweight and self-contained.
