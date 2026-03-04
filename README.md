# Jolify — PowerPoint Add-in

A native PowerPoint ribbon add-in with shape positioning, alignment, formatting, and text utilities. Built for teams that work intensively with slide decks.

**Live site:** https://stephenjoly.github.io/jolify-ppt/

---

## Features

| Group | Buttons |
|-------|---------|
| **Position** | Copy/Paste Pos+Size, Copy/Paste Position, Copy/Paste Size, Swap Positions |
| **Match** | Copy Outline, Copy Fill, Match Height, Match Width, Match H+W |
| **Align** | Align Left/Center/Right/Top/Middle/Bottom, Distribute H/V |
| **Text** | Split Text Box |
| **Branding** | Add Draft, Remove Draft |

---

## Install (end users)

**Option A — Terminal:**
```bash
curl -fsSL https://stephenjoly.github.io/jolify-ppt/install.sh | bash
```

**Option B — Manual:**
1. Download `manifest.xml` from https://stephenjoly.github.io/jolify-ppt/
2. Open PowerPoint → Insert → Add-ins → Upload My Add-in → select the file
3. Restart PowerPoint

The *Jolify* tab appears in the ribbon. Re-installation is not needed for future updates.

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

Starts the webpack dev server on `https://localhost:3300` and sideloads the add-in into PowerPoint automatically.

### Stop

```bash
npm stop
```

---

## Deployment

Pushing to `main` triggers the GitHub Actions workflow which builds the project and deploys it to GitHub Pages automatically. No server or Docker required.

**Production URL:** `https://stephenjoly.github.io/jolify-ppt/`

The manifest only needs to be re-sideloaded by users if it structurally changes (e.g. new ribbon buttons). All JS/HTML updates are picked up automatically from the live site.
