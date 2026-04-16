# Personal Tools Repo

This repository is now organized as a **collection of small personal tools**.

## Current tools

- `tools/enex-onenote/` – ENEX → OneNote converter with a modern browser UI.

## Repo structure

- `index.html` – tools home page / launcher.
- `styles.css` – styles for the tools home page.
- `tools/enex-onenote/index.html` – converter interface.
- `tools/enex-onenote/styles.css` – converter styles.
- `tools/enex-onenote/app.js` – ENEX parsing + DOCX generation + download flow.

## ENEX → OneNote converter highlights

- Drag & drop or browse multiple ENEX files.
- Progress bar and status updates.
- Error panel with partial-failure reporting.
- Structured ZIP output grouped by ENEX filename.
- Optional single-DOCX download when exactly one note is converted.

## Cloudflare Pages deployment

This repo is static and deploys directly with Cloudflare Pages.

1. Push this repository to GitHub/GitLab.
2. In Cloudflare Dashboard, create a new **Pages** project and connect the repo.
3. Build settings:
   - **Framework preset:** None
   - **Build command:** *(leave empty)*
   - **Build output directory:** `/`
4. Deploy.

Your site root (`/`) serves the tool launcher, and the first tool is available at:

- `/tools/enex-onenote/`
