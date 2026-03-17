# Deployment Context

This folder contains the optional self-hosted deployment path. The default release path for the repo is still GitHub Pages.

## Files

- `Dockerfile`: multi-stage Node build to Nginx image.
- `docker-compose.yml`: local/server compose wrapper for that image.
- `nginx/default.conf`: static hosting plus `/healthz`.
- `traefik/ppt-addin.yml`: sample reverse-proxy config.
- `README.md`: container usage notes.

## When `deploy/` Matters

Use this folder when you need:

- self-hosting instead of GitHub Pages
- custom infrastructure or private hosting
- explicit reverse-proxy/TLS control

## Related Repo Paths

- `.github/workflows/docker-publish.yml` builds and pushes the image manually.
- `.github/workflows/deploy-pages.yml` remains the primary static-site release path.

## Operational Notes

- Office add-ins must be served over trusted HTTPS.
- If you self-host behind Nginx or Traefik, make sure the manifest and static assets are reachable at stable HTTPS URLs that Office trusts.
