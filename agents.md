# Jolify Agent Notes

This file is a lightweight working agreement for anyone helping on the repo.

## Project Priorities

- Keep Jolify practical for real slide work. Ribbon ergonomics and PowerPoint behavior matter as much as code cleanliness.
- Prefer shared action logic over duplicated command/taskpane behavior.
- Treat manifest changes carefully because they can change ribbon rendering in non-obvious ways.

## How To Work Safely Here

- Read existing `context.md` files before making broad structural changes.
- Avoid deleting files unless they are clearly unused and not part of active feature work.
- Do not assume a clean git worktree; the repo often contains in-progress local changes.
- Keep edits scoped and reversible.

## Cleanup Rules

- When removing a feature from the supported product surface, remove it end-to-end in the same pass:
  - `manifest.xml`
  - `src/commands/commands.ts`
  - `src/taskpane/taskpane.ts`
  - `src/taskpane/taskpane.html`
  - any dialog HTML in `src/dialogs/`
  - any webpack outputs tied to that dialog or command
  - any icon assets and icon-generation mappings
  - any related docs, context files, and backlog references
- Do not leave a feature half-retired. A command that is removed from the ribbon but still lives in the taskpane, webpack build, or shared code is considered repo bloat and a future debugging risk.
- After a cleanup pass, search the repo for the removed feature name and command IDs to confirm no stale bindings remain.
- Treat checked-in generated files skeptically. If webpack or another build step generates an artifact, prefer deleting the source copy unless the build truly depends on it.
- Keep repo layout by intent:
  - `install/` for user-facing install/uninstall entrypoints
  - `scripts/` for developer/build utilities
  - `docs/` for planning and troubleshooting material that is not part of the runtime

## Ribbon And Manifest Rules

- `manifest.xml` is the source of truth for ribbon layout and command bindings.
- Dense ribbon groups tend to render as smaller buttons in PowerPoint; too many thin groups often produce oversized buttons.
- After changing `manifest.xml`, reload with `npm stop`, `npm start`, and fully restart PowerPoint.
- Keep `<FunctionName>` values aligned with the names exposed in `src/commands/commands.ts`.
- If the ribbon renders but custom Jolify buttons are inert while native PowerPoint controls still work, assume the hidden command runtime is broken before changing the ribbon layout.
- In that case, check `https://localhost:3300/commands.html`, inspect port `3300`, and follow `docs/TROUBLESHOOTING.md` before making more code changes.

## Shared Code Rules

- Put behavior in `src/shared/shapeTools.ts`, not in ribbon/taskpane wrappers.
- Preserve the `ActionResult` contract for user-facing actions.
- Reuse dialogs for structured inputs instead of hard-coding temporary prompts.

## Current Repo Reality

- There is also a local-stable runtime path (`install/install-local.sh`, `install/uninstall-local.sh`, `scripts/local_server.py`, `scripts/package-local-runtime.js`).
- `npm run validate` currently reports a known manifest package/product-ID issue; treat new validation errors separately from that baseline.

## Verification Expectations

- For code changes, prefer at least one concrete verification step (`npm run build`, `xmllint --noout manifest.xml`, or local PowerPoint retest guidance).
- If something cannot be fully verified locally, state that explicitly.
- When local Office behavior is inconsistent, include runtime troubleshooting as part of normal verification:
  - verify what `commands.html` serves
  - check for stale `webpack` processes on port `3300`
  - check the installed WEF manifest
  - prefer fixing stale runtime state before expanding the command surface
- After feature removal or ribbon simplification, also verify that the taskpane and build output no longer surface the removed feature.
