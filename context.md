# Jolify Project Context

Jolify is a PowerPoint add-in for slide production work. It combines a custom ribbon, an optional taskpane, static dialog pages, and a small AI helper layer for leadline generation.

## Core Runtime Shape

1. `manifest.xml` defines the Jolify tab, ribbon controls, native Office controls, and command bindings.
2. `src/commands/commands.ts` exposes ribbon handlers on `window`.
3. `src/shared/shapeTools.ts` implements the main Office.js actions.
4. `src/taskpane/taskpane.ts` wires taskpane buttons to the same shared actions.
5. Static HTML dialogs in `src/dialogs/` send JSON payloads back through `Office.context.ui.messageParent(...)`.

## Repository Map

- `src/shared/`: main behavior layer. `shapeTools.ts` is the main action file.
- `src/commands/`: Office function-file entrypoint for ribbon commands.
- `src/taskpane/`: taskpane UI and status handling for smoke-testing and discoverability.
- `src/dialogs/`: static dialog pages for builders and AI setup/briefing.
- `assets/`: app icons, ribbon icons, sticker assets, and site branding imagery.
- `scripts/`: icon generation, local HTTPS file server, and local-runtime packaging.
- `deploy/`: optional Docker/Nginx deployment path.
- `.github/workflows/`: GitHub Pages and container-publish automation.

## Build and Install Paths

- `npm start`: local webpack dev server plus Office sideload flow.
- `npm stop`: stop the local debug/sideload flow.
- `npm run build`: production webpack build plus local-runtime bundle packaging.
- `install/install.sh` / `install/uninstall.sh`: hosted-mode installer/uninstaller sources.
- `install/install-local.sh` / `install/uninstall-local.sh`: macOS local-stable runtime installer/uninstaller sources.

## Feature Areas In The Repo

- Shape positioning and size copy/paste.
- Alignment, grouped alignment, and distribution helpers.
- Style transfer and connector normalization.
- Text utilities such as split/merge/auto-flow/font equalization.
- Builder flows for grids.
- Diagnostics and slide operations.
- AI leadline generation and AI settings storage.

## Documentation Map

- `README.md`: end-user install and development overview.
- `docs/FUTURE_FEATURES.md`: backlog of intentionally removed or deferred features.
- `docs/TROUBLESHOOTING.md`: local runtime and sideload failure playbook.
- `src/context.md`: source-tree structure and feature routing.
- `src/shared/context.md`: shared action layer and AI helper notes.
- `src/commands/context.md`: command binding rules.
- `src/taskpane/context.md`: taskpane action wiring.
- `src/dialogs/context.md`: dialog inventory and messaging contracts.
- `assets/context.md`: asset and icon guidance.
- `scripts/context.md`: build/support scripts.
- `deploy/context.md`: self-hosted deployment notes.
- `.github/workflows/context.md`: CI/CD behavior.

## Current Caveats

- Several layout/build actions assume a 16:9 slide model (`960x540` points).
- The manifest is actively tuned for ribbon layout, so control regrouping changes can alter PowerPoint button rendering.
- `npm run validate` currently reports a pre-existing package/product-ID problem unrelated to normal local development.
