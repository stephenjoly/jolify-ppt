# Deployment Checklist

This is the repeatable Jolify release path once `Jolify Dev` looks correct in PowerPoint.

## Release Model

- `Jolify Dev` is the localhost sideloaded add-in from `dev/manifest.xml`.
- `Jolify` hosted mode uses the stable `manifest.xml`, but that manifest is copied into the user's WEF folder during install.
- `Jolify` local stable mode downloads `jolify-local-bundle.tar.gz`, installs a localhost runtime, and copies a local manifest into the user's WEF folder.
- Pushing to `main` rebuilds `dist/` and deploys GitHub Pages automatically through [deploy-pages.yml](/Users/stephenjoly/Documents/Coding/ppt-addin/.github/workflows/deploy-pages.yml).

## What Updates Automatically

- GitHub Pages website: updates automatically after the Actions deploy finishes.
- Hosted runtime assets (`commands.js`, `taskpane.html`, icons, dialogs): update automatically for users whose installed stable manifest already points at the hosted site.
- Hosted stable manifest in an existing user's WEF folder: does **not** auto-update. Users need to rerun `install.sh` if the shipped `manifest.xml` changed.
- Local stable runtime: does **not** auto-update. Users need to rerun `install-local.sh` to pull the new bundle and manifest.

## Standard Release Checklist

1. Finish the change in `Jolify Dev`.
2. In PowerPoint, fully quit and relaunch after manifest changes:
   - `npm stop`
   - `npm start`
   - fully restart PowerPoint
3. Confirm the stable-facing behavior in `Jolify Dev`.
   - ribbon renders correctly
   - custom Jolify buttons still execute
   - any native injected controls still render
4. Check repo state:
   - `git status --short`
   - commit the working increment before release
5. Run the release verification commands:
   ```bash
   xmllint --noout manifest.xml
   xmllint --noout dev/manifest.xml
   npm run build
   ```
6. Optionally run:
   ```bash
   npm run validate
   ```
   Note: this repo has a known baseline package/product-ID validation issue. Treat new validation errors separately from that baseline.
7. Re-read the generated release surface in `dist/` if the change touched installers, icons, or site content.
8. Push to `main`.
9. Wait for the GitHub Pages workflow to finish successfully.
10. Verify production outputs:
   - `https://stephenjoly.github.io/jolify-ppt/`
   - `https://stephenjoly.github.io/jolify-ppt/manifest.xml`
   - `https://stephenjoly.github.io/jolify-ppt/install.sh`
   - `https://stephenjoly.github.io/jolify-ppt/install-local.sh`
11. If the stable manifest changed, rerun hosted install on machines that should pick up the new ribbon:
   ```bash
   curl -fsSL https://stephenjoly.github.io/jolify-ppt/install.sh | bash
   ```
12. If the local stable runtime changed, rerun the local stable installer on machines that should pick up the new bundle:
   ```bash
   curl -fsSL https://stephenjoly.github.io/jolify-ppt/install-local.sh | bash
   ```

## Fast Decision Rules

- Changed only website copy or hosted JS/HTML assets without changing `manifest.xml`:
  existing hosted installs should pick up the update after the Pages deploy.
- Changed `manifest.xml`:
  hosted users need to reinstall to get the new stable manifest.
- Changed anything used by the packaged local runtime:
  local stable users need to rerun `install-local.sh`.
- Changed only `dev/manifest.xml` or local dev behavior:
  no production rollout is needed until the stable `manifest.xml` and production build are updated.

## Quick Rollout Routine

Use this when you want the shortest safe path:

```bash
git status --short
xmllint --noout manifest.xml
xmllint --noout dev/manifest.xml
npm run build
git add -A
git commit -m "Release: describe change"
git push origin main
```

Then:

1. Watch the Pages deployment succeed.
2. Test the live site.
3. Reinstall hosted mode if the manifest changed.
4. Reinstall local stable mode if the local bundle changed.

## Local Verification After Release

Hosted mode:

```bash
curl -fsSL https://stephenjoly.github.io/jolify-ppt/install.sh | bash
```

Local stable mode:

```bash
curl -fsSL https://stephenjoly.github.io/jolify-ppt/install-local.sh | bash
```

Use hosted mode when you want the simplest update path. Use local stable mode when you want the runtime hosted on the machine with launchd keeping it alive.
