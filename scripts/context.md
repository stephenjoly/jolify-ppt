# Scripts Context

Utility scripts for asset generation and local-runtime packaging.

## Files

- `generate-icons.js`: generates ribbon PNGs from Fluent UI SVG sources into `assets/icons/`.
- `package-local-runtime.js`: packages the production build into `dist/jolify-local-bundle.tar.gz` and writes `dist/manifest.local.xml`.
- `local_server.py`: local-stable HTTPS runtime. It serves built assets, exposes `/healthz`, and handles native save/open/email/picture-deck bridge actions on macOS.

## Typical Usage

- Regenerate icons manually:

```bash
node scripts/generate-icons.js
```

- Package the local runtime:

```bash
npm run build
```

That build runs webpack first and then `package-local-runtime.js`.

## Update Guidance

- When adding a command, update icon generation only if it needs a dedicated ribbon icon.
- Keep the local-runtime packager aligned with the actual production build outputs in `dist/`.
- Keep the local HTTPS server intentionally narrow: serve built assets and only the native endpoints required for macOS save/open/email/picture-deck workflows.
