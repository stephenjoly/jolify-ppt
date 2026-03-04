# ppt-addin Docker Deployment

This folder bundles everything required to build and run a production container for the PowerPoint add-in. The image builds the static files, copies them into an Nginx container, and serves them over HTTP. Place this folder in a public repository and use it from your home server or any container host.

## Build & Publish the Image

1. From the repository root (the folder containing `package.json`) run:

   ```bash
   docker compose -f deploy/docker-compose.yml build
   ```

   The multi-stage Dockerfile installs dependencies, runs `npm run build`, and copies the resulting `dist/` assets into the final image.

2. Tag the image for your registry, for example GitHub Container Registry:

   ```bash
   docker tag ppt-addin:latest ghcr.io/<your-github-username>/ppt-addin:latest
   docker push ghcr.io/<your-github-username>/ppt-addin:latest
   ```

3. Update the manifest (`manifest.xml` and `dist/manifest.xml`) to point to the final HTTPS endpoint that will host the add-in, then distribute the manifest to users (side-load or publish via the Microsoft 365 admin center).

## Run with Docker Compose

On your home server (or any host with Docker Compose), pull the image and start the service:

```bash
docker compose -f deploy/docker-compose.yml up -d
```

By default the container listens on port `80` internally and the compose file maps it to `8080`. Adjust `ports` if you need a different host port.

Office add-ins require HTTPS with a certificate trusted by the client. Place this container behind a reverse proxy (Traefik, Caddy, Nginx, etc.) that terminates TLS, or extend the Dockerfile to add certificates directly if you prefer end-to-end TLS inside the container.

A health endpoint is available at `/healthz` for simple uptime checks.

## Updating the Add-in

Whenever you make changes to the add-in code:

1. Rebuild the image (`docker compose build`).
2. Push it to your registry.
3. Restart the service (`docker compose up -d` on the host).

Users only need the manifest pointing to your HTTPS endpoint; they do not need to run `npm start` locally.
