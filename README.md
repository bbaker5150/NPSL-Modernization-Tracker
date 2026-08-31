# NPSL Modernization Project Tracker

A single-file, SharePoint-backed web application for visualizing the NPSL modernization portfolio from initial need through sustainment and closeout.

The included baseline was normalized from `POAM_NPSL_Mod_Projects.xlsx`:

- 14 measurement-area projects
- 34 WBS tasks per project (476 task records)
- 7 visual pipeline stages
- Historical dates and workbook notes preserved
- Invalid source date orderings flagged for review rather than silently changed

## What the app provides

- Portfolio dashboard with project, progress, attention, and risk metrics
- Seven-stage cross-project pipeline visualizer
- Kanban-style pipeline board
- Searchable project register
- Engineer-focused **My work** view using the signed-in SharePoint email
- Project detail workspace with WBS tasks, updates, risks, ownership, health, and milestones
- Project and task editing, project claiming, status updates, risk capture, and CSV export
- Light and dark themes
- Browser-local demo mode outside SharePoint
- Automatic SharePoint mode when hosted from a `sharepoint.com` site
- Forge-compatible single HTML build with no external runtime requests

## Local development

```bash
npm install
npm run dev
```

The local build uses browser storage and starts with the workbook baseline. Use this mode for design review and testing without touching SharePoint.

## Build

Standard multi-file build:

```bash
npm run build
```

Forge single-file build:

```bash
npm run build:singlefile
npm run verify:singlefile
```

Deploy `build-singlefile/modernization-project-tracker.html` through the same Forge/App Page flow used by the existing portal tools.

## Main-branch build pipeline

`.github/workflows/singlefile.yml` runs on every push to `main` and on manual dispatch. It:

1. installs from `package-lock.json`;
2. audits and tests the application;
3. stamps and builds the Forge-ready single HTML;
4. verifies the manifest and zero-external-subresource contract;
5. boots the finished file inside an `iframe srcdoc` and exercises the portfolio, board, and project WBS;
6. uploads the HTML and SHA-256 file as a 90-day workflow artifact; and
7. creates an immutable per-commit GitHub release so the newest build is always available from:

`https://github.com/<owner>/<repo>/releases/latest/download/modernization-project-tracker.html`

## SharePoint configuration

The app discovers its SharePoint web from `_spPageContextInfo`, the same-origin parent frame, or the `/sites/<name>` URL. A deployment can override defaults before the app bundle runs:

```html
<script>
  window.MOD_TRACKER_CONFIG = {
    listPrefix: 'Modernization',
    webUrl: 'https://tenant.sharepoint.com/sites/NPSL',
    managerEmails: [
      'portfolio.manager@navy.mil'
    ]
  };
</script>
```

Configuration options:

| Option | Default | Purpose |
| --- | --- | --- |
| `listPrefix` | `Modernization` | Prefix for the four SharePoint List titles. |
| `webUrl` | Auto-detected | Target SharePoint web URL. |
| `managerEmails` | `[]` | Users allowed to see the full portfolio navigation. Empty means all site users can see portfolio views. |
| `forceLocal` | `false` | Force browser-local demo storage for troubleshooting. |

On first SharePoint load, the setup screen creates the lists and can import the workbook baseline. The signed-in user needs Edit or Full Control.

## Data model

See [SHAREPOINT_DEPLOYMENT.md](SHAREPOINT_DEPLOYMENT.md) for list fields, permissions, configuration, and deployment behavior.

## Important baseline note

The workbook contains historical planning data. The app labels imported projects **Needs Review** and displays a persistent validation notice. Project owners, current health, and target dates should be confirmed before current portfolio reporting begins.
