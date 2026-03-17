# Workflow Context

This folder holds the CI/CD automation for static hosting and optional container publishing.

## `deploy-pages.yml`

Primary release path.

Trigger:

- push to `main`

Flow:

1. checkout repository
2. set up Node 20
3. `npm ci`
4. `node scripts/generate-icons.js`
5. `npm run build`
6. publish `dist/` to GitHub Pages

Because `npm run build` also runs `scripts/package-local-runtime.js`, the Pages build publishes both the normal static site and the downloadable local-runtime bundle.

## `docker-publish.yml`

Optional container path.

- manual trigger only (`workflow_dispatch`)
- builds and pushes the Docker image to GHCR
- requires the `CR_PAT` secret

Use this only if you are actively using the self-hosted Docker deployment path.
