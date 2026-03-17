# Taskpane Context

The taskpane is a secondary launcher for many Jolify actions. It is mainly useful for development, smoke-testing, and discoverability.

## Files

- `taskpane.html`: button layout and status region.
- `taskpane.ts`: `ACTIONS` map, busy-state handling, and result rendering.

## Current Wiring

- `ACTIONS` is the source of truth from button ID to action runner.
- All current taskpane actions come from `shapeTools.ts`.
- `setBusy(...)` disables the full taskpane while an action runs.
- `setStatus(...)` renders `ActionResult` consistently.

## Editing Guidance

- Keep button IDs aligned between HTML and the `ACTIONS` map.
- Add new actions here only if they are useful for testing or discovery; the ribbon is still the primary UX.
- Do not strand new user-facing workflows in the taskpane; if a feature matters for production use, wire it to the ribbon.
