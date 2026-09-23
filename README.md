# Modernization Project Tracker

A single-file, SharePoint-backed web application for visualizing modernization work from initial need through sustainment and closeout.

## What the app provides

- Portfolio dashboard, five-phase pipeline, Kanban board, and searchable project register
- Engineer-focused **My work** queues tied to the signed-in SharePoint login identity
- Project pipeline tasks, updates, risks, ownership, health, and milestones
- Manager portfolio access and ownership-scoped standard-user views
- The original 37-entry acronym reference, with manager add/remove controls backed by SharePoint
- Styled multi-sheet Excel export including the current acronym reference
- Light and dark themes with the NAVAIR seal in the tracker identity
- A Forge-compatible single HTML build with no external runtime requests

## Tasking and roles

- Portfolio projects now uses the full searchable row register; Needs attention opens every unfinished project with pending tasks, owners, overdue flags, deferrals and justifications.
- Managers maintain users/roles, assign projects and deadlines, and add/delete tasks within each project phase.
- Standard users see their assigned work and may change only task status, deferred dates, and justifications.
- Deferrals require a date after the original deadline and a justification. Not Required requires justification. The original due date is retained.
- Task codes have been removed from the UI, templates, exports, and active SharePoint schema.

## Local development

```bash
npm install
npm run dev
```

Local development uses clean-slate browser storage and the same empty portfolio behavior as SharePoint.

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
| `listPrefix` | `Modernization` | Prefix for the six SharePoint List titles. |
| `webUrl` | Auto-detected | Target SharePoint web URL. |
| `forceLocal` | `false` | Force clean-slate browser storage for local troubleshooting. |
| `forceSharePoint` | `false` | Force REST mode for an on-premises or custom SharePoint host. |
| `hideLists` | `true` | Hide the six backing lists from the normal Site Contents view without affecting API access. |

On first SharePoint load, the app silently creates empty lists and missing fields while the loading screen is displayed. Newly created backing lists are hidden from the normal Site Contents view in their initial create request, so startup does not issue follow-up MERGE operations. The signed-in user needs permission to create lists and fields for that first run; normal use needs whatever read/edit rights the site owner grants. The saved Users directory determines application roles; SharePoint site administrators bootstrap manager access. Unknown identities default to standard users. SharePoint permissions remain the server-side access boundary; see the deployment guide before granting list access.

See [SHAREPOINT_DEPLOYMENT.md](SHAREPOINT_DEPLOYMENT.md) for the list schema and deployment behavior.
