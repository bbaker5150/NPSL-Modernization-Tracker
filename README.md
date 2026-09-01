# Modernization Project Tracker

A single-file, SharePoint-backed web application for visualizing modernization work from initial need through sustainment and closeout.

## What the app provides

- Portfolio dashboard, five-phase pipeline, Kanban board, and searchable project register
- Engineer-focused **My work** queues tied to the signed-in SharePoint login identity
- Project WBS tasks, updates, risks, ownership, health, and milestones
- Full portfolio access for every user who has permission to the hosting SharePoint site and lists
- An editable sample portfolio for onboarding and visualization; sample changes stay in memory and are never written to SharePoint
- Searchable acronym glossary and styled multi-sheet Excel export with the same reference data
- Light and dark themes with the NAVAIR seal in the tracker identity
- A Forge-compatible single HTML build with no external runtime requests

The bundled mock portfolio was normalized from the August 2026 modernization review workbook. Its 28 projects and 161 phase/action tasks preserve the source's discipline mix, pipeline distribution, action coverage, due-date coverage, and risk states; project details and personnel are replaced with neutral scenarios and consistent role aliases. The live SharePoint workspace always starts empty.

## Local development

```bash
npm install
npm run dev
```

Local development uses clean-slate browser storage. Use **Open sample portfolio** in the app to practice with the example pipeline without persisting those changes.

## Build

```bash
npm run build:singlefile
npm run verify:singlefile
```

Deploy `build-singlefile/modernization-project-tracker.html` through the established Forge/App Page flow used by the existing SharePoint tools.

## Main-branch pipeline

`.github/workflows/singlefile.yml` runs on every push to `main` and on manual dispatch. It installs locked dependencies, audits, runs all tests, builds the Forge-ready HTML, verifies the manifest and zero-external-resource contract, exercises the finished file inside an `iframe srcdoc`, uploads a 90-day workflow artifact, and publishes an immutable GitHub release.

The newest successful build is always available at:

`https://github.com/<owner>/<repo>/releases/latest/download/modernization-project-tracker.html`

## SharePoint configuration

The app discovers its SharePoint web from `_spPageContextInfo`, the same-origin parent frame, or the `/sites/<name>` URL. A deployment can override defaults before the app bundle runs:

```html
<script>
  window.MOD_TRACKER_CONFIG = {
    listPrefix: 'Modernization',
    webUrl: 'https://tenant.sharepoint.com/sites/Modernization',
    hideLists: true
  };
</script>
```

| Option | Default | Purpose |
| --- | --- | --- |
| `listPrefix` | `Modernization` | Prefix for the four SharePoint List titles. |
| `webUrl` | Auto-detected | Target SharePoint web URL. |
| `forceLocal` | `false` | Force clean-slate browser storage for local troubleshooting. |
| `forceSharePoint` | `false` | Force REST mode for an on-premises or custom SharePoint host. |
| `hideLists` | `true` | Hide the four backing lists from the normal Site Contents view without affecting API access. |

On first SharePoint load, the app silently creates empty lists and missing fields while the loading screen is displayed. Newly created backing lists are hidden from the normal Site Contents view in their initial create request, so startup does not issue follow-up MERGE operations. The signed-in user needs permission to create lists and fields for that first run; normal use needs whatever read/edit rights the site owner grants. SharePoint permissions are the access boundary—there is no separate manager allowlist in the app.

See [SHAREPOINT_DEPLOYMENT.md](SHAREPOINT_DEPLOYMENT.md) for the list schema and deployment behavior.
