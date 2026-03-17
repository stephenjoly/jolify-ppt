# Jolify Troubleshooting Playbook

This document captures the common local-debugging failure modes for Jolify and the steps to take before changing more code.

## 1. Ribbon Tab Shows But Jolify Buttons Do Nothing

Typical symptom:
- The `Jolify` tab is visible.
- Native PowerPoint controls such as `Shape Fill`, `Shape Outline`, or `Font Color` still work.
- Custom Jolify buttons do nothing.

What this usually means:
- The manifest loaded, so PowerPoint can render the ribbon.
- The hidden function-file runtime (`commands.html` / `commands.js`) did not load correctly, so `ExecuteFunction` handlers were never bound.

### Debug Steps

1. Check what the manifest points to.
   - Local dev should point to `https://localhost:3300/taskpane.html` and `https://localhost:3300/commands.html`.
   - `npm start` uses `dev/manifest.xml`, not `manifest.xml`, so the WEF folder should show a separate `Jolify Dev` add-in identity during local testing.

2. Check the installed WEF manifest.
   - Common location on this machine:
   - `$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

3. Check whether port `3300` is already occupied.
   ```bash
   lsof -nP -iTCP:3300 -sTCP:LISTEN
   ```

4. If something is already listening on `3300`, inspect it.
   ```bash
   ps -p <PID> -o pid,ppid,command=
   ```

5. Check what `commands.html` is actually serving.
   ```bash
   curl -k https://localhost:3300/commands.html | sed -n '1,40p'
   ```

6. If `commands.html` shows an `HtmlWebpackPlugin` error page instead of a real HTML function file, the ribbon buttons will be inert even though the tab renders.

### Fix

Kill the stale dev server and restart cleanly:
```bash
kill <PID>
npm stop
pkill -f "webpack serve" || true
npm start
```

Then fully quit and reopen PowerPoint.

## 2. Add-in Error Appears And Jolify Won't Open

Typical symptom:
- PowerPoint reports an add-in error.
- The Jolify tab or add-in surface fails to open.

### Debug Steps

1. Validate the current local build:
   ```bash
   npm run build:dev
   xmllint --noout manifest.xml
   xmllint --noout dev/manifest.xml
   ```

2. Treat `npm run validate` carefully.
   - This repo has a known baseline manifest validation issue around package/product ID.
   - Do not confuse those baseline errors with a new local regression.

3. If the failure started right after adding ribbon controls or command imports:
   - back out the last ribbon expansion first
   - restore the smaller known-good command surface

## 3. `npm start` Says The Dev Server Is Already Running

Typical symptom:
- `npm start` prints `The dev server is already running on port 3300.`

Risk:
- PowerPoint may reconnect to an old in-memory webpack build instead of the current code.

### Fix

Force a real restart:
```bash
npm stop
pkill -f "webpack serve" || true
npm start
```

## 4. Uninstall Scripts Fail With `permission denied`

Typical symptom:
- `./install/uninstall.sh` or `./install/uninstall-local.sh` fails with `permission denied`

Cause:
- The script is not executable.

### Fix

Run it with `bash`:
```bash
bash ./install/uninstall.sh
bash ./install/uninstall-local.sh
```

## 5. WEF Folder Audit

Useful check when Jolify seems duplicated or stale:

```bash
for d in \
  "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef" \
  "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Wef" \
  "$HOME/Library/Application Support/Microsoft/Office/16.0/Wef"; do
  printf '=== %s ===\n' "$d"
  if [ -d "$d" ]; then
    ls -al "$d"
  else
    echo "(missing)"
  fi
done
```

Interpretation:
- One Jolify manifest in WEF is expected during local dev.
- If stable and local dev are both installed, seeing two Jolify manifests is expected as long as one is the `Jolify Dev` manifest ID.
- The repo having `manifest.xml`, `dev/manifest.xml`, and `dist/manifest.xml` is not itself the problem.

## 6. If Custom Ribbon Buttons Still Fail After A Clean Restart

Do not immediately rewrite the ribbon layout.

Instead:
1. Confirm `commands.html` serves a real function file, not an error page.
2. Confirm the command bundle builds cleanly.
3. Reduce to one minimal custom command if needed to isolate `ExecuteFunction`.
4. Only after the command runtime is healthy should you add more ribbon controls or feature groups.

## 7. Working Rule

When native PowerPoint ribbon controls work but Jolify buttons do not, assume a command-runtime problem first, not a ribbon-layout problem.
