# Jolify — PowerPoint Add-in

A custom PowerPoint task pane add-in with shape positioning, alignment, formatting, and text utilities.

## Features

### Position & Size
- Copy/paste position and size between shapes (clipboard-style)
- Swap positions between two shapes

### Match (first selected shape → all others)
- Copy outline style (color, dash, weight, transparency)
- Copy fill style (solid color or no-fill)
- Match height, width, or both

### Align
- Align left, center, right, top, middle, bottom
  - 1 shape selected: aligns to the slide
  - Multiple shapes: aligns relative to each other
- Distribute horizontally or vertically (3+ shapes)

### Text
- Split a text box into one text box per line

### Branding
- Add/remove a "DRAFT" sticker to the slide master

---

## Development

### Prerequisites
- Node.js 20+
- PowerPoint for Mac (Microsoft 365)

### Run locally

```bash
npm install
npm start
```

This starts the webpack dev server on `https://localhost:3300` and sideloads the add-in into PowerPoint automatically.

---

## Deployment

The add-in is a static site (HTML + JS) served by Nginx in a Docker container, sitting behind Traefik for HTTPS.

### CI/CD

Pushing to `main` triggers a GitHub Action that builds the Docker image and publishes it to GHCR:

```
ghcr.io/stephenjoly/jolify-ppt/ppt-addin:latest
ghcr.io/stephenjoly/jolify-ppt/ppt-addin:<commit-sha>
```

Required GitHub secret: `CR_PAT` — a personal access token with `write:packages` scope.

### Server setup

Copy `deploy/` to your server. Add `deploy/traefik/ppt-addin.yml` to your Traefik dynamic config directory, then:

```bash
docker compose -f deploy/docker-compose.yml up -d
```

### Upgrading

```bash
docker compose -f deploy/docker-compose.yml pull
docker compose -f deploy/docker-compose.yml up -d
```

### Sideloading the manifest (one-time)

Build locally to generate the production manifest:

```bash
npm run build
# → dist/manifest.xml now contains https://ppt-addon.stephenjoly.net URLs
```

Copy it to PowerPoint's sideload folder:

```bash
# App Store version of Office:
cp dist/manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/

# Retail/volume license version:
cp dist/manifest.xml ~/Library/Application\ Support/Microsoft/Office/16.0/Wef/
```

Restart PowerPoint. The manifest only needs to be re-sideloaded if the manifest itself changes (e.g. a new ribbon button or a URL change). All code changes are picked up automatically from the server.
