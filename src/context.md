# Source Context (`src/`)

This folder contains the add-in runtime code and its static dialog UIs.

## Structure

- `shared/`: Office.js action implementations.
- `commands/`: ribbon/function-file entrypoint.
- `taskpane/`: optional action launcher UI for testing and discovery.
- `dialogs/`: static HTML dialog pages for the remaining builders.

## Runtime Flow

1. Ribbon click maps from `manifest.xml` to a function name.
2. `src/commands/commands.ts` receives the Office command event.
3. Command wrapper calls into `src/shared/shapeTools.ts`.
4. Shared action runs in `PowerPoint.run(...)` or through Office dialog APIs and returns `ActionResult`.

Taskpane clicks bypass the Office command-event wrapper and call the same shared actions directly.

## Current Feature Families

- position/size copy-paste
- style transfer
- alignment and grouped alignment
- text box transforms
- layout and chart builders
- slide operations and presentation helpers

## Editing Guidance

- Keep command names stable once they are wired into `manifest.xml`.
- Prefer adding focused helper functions in `shared/` instead of expanding already-large blocks without structure.
- If a feature needs structured input, add a dialog in `dialogs/` instead of overloading ribbon actions with hard-coded assumptions.
