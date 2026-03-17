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

### Hosted mode

Best for users who want the lightest install and do not want a local background service.

```bash
curl -fsSL https://stephenjoly.github.io/jolify-ppt/install.sh | bash
```

Manual alternative:
1. Download `manifest.xml` from https://stephenjoly.github.io/jolify-ppt/
2. Open PowerPoint → Insert → Add-ins → Upload My Add-in → select the file
3. Restart PowerPoint

Uninstall:

```bash
curl -fsSL https://stephenjoly.github.io/jolify-ppt/uninstall.sh | bash
```

### Local stable mode

Best for users who want Jolify hosted locally with a restart-safe macOS `launchd` agent.

```bash
curl -fsSL https://stephenjoly.github.io/jolify-ppt/install-local.sh | bash
```

This installs a localhost-served Jolify runtime, a user-level `launchd` agent, and a trusted localhost certificate. PowerPoint will use the local runtime instead of the GitHub-hosted one.

Uninstall:

```bash
curl -fsSL https://stephenjoly.github.io/jolify-ppt/uninstall-local.sh | bash
```

Notes:
- Hosted mode is the simplest update path.
- Local stable mode is more resilient after restarts or accidental process termination.
- AI features still require internet access in either mode.

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

Starts the webpack dev server on `https://localhost:3300` and sideloads `dev/manifest.xml` into PowerPoint automatically.
The dev add-in uses a separate add-in ID and `Jolify Dev` tab, so it can coexist with the stable installed add-in.

### Stop

```bash
npm stop
```

This stops the `Jolify Dev` sideload only. It does not remove a stable hosted or local-stable install.

## Developer docs map

- `agents.md` - project-specific working norms for contributors and coding agents
- `context.md` - project architecture and contributor entry point
- `docs/FUTURE_FEATURES.md` - feature backlog intentionally kept out of the stable ribbon
- `docs/TROUBLESHOOTING.md` - local runtime and sideload troubleshooting playbook
- `manifest.xml` - stable hosted/local-stable manifest
- `dev/manifest.xml` - local sideload manifest for side-by-side development
- `install/install.sh` - hosted installer source
- `install/uninstall.sh` - hosted uninstaller source
- `install/install-local.sh` - local-mode installer source
- `install/uninstall-local.sh` - local-mode uninstaller source
- `src/context.md` - source tree overview
- `src/shared/context.md` - action layer conventions and caveats
- `src/commands/context.md` - ribbon function binding rules
- `src/taskpane/context.md` - taskpane wiring and status model
- `src/dialogs/context.md` - dialog contracts and messaging pattern
- `assets/context.md` - static assets and icon generation notes
- `scripts/context.md` - utility script usage
- `deploy/context.md` - container deployment context
- `.github/workflows/context.md` - CI/CD workflow behavior

---

## Deployment

Pushing to `main` triggers the GitHub Actions workflow which builds the project and deploys it to GitHub Pages automatically. No server or Docker required.

**Production URL:** `https://stephenjoly.github.io/jolify-ppt/`

Hosted mode picks up JS/HTML updates automatically from the live site. Local stable mode is installed from the hosted site but then serves Jolify locally through its launch agent.
