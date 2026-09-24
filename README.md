# Modernization Project Tracker

A single-file, SharePoint-backed web application for visualizing modernization work from initial need through sustainment and closeout.

## What the app provides

- Portfolio dashboard, four-phase pipeline, Kanban board, and searchable project register
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

## Testing manager access

New users land on **My work** and can also open **Acronym glossary** and **Users and managers**. Sign-ins automatically register the SharePoint name/email in the shared directory. On **Users and managers**, enter `admin123` under **Testing manager access** to promote the current signed-in identity; no separate username is needed. Existing saved manager roles remain in place.

This is a testing-only client-side convenience, not secure authentication. Set `window.MOD_TRACKER_CONFIG.testingManagerPassword` to a different test password, or `false` to disable promotion. Disabling the flow does not demote previously promoted users; a manager must change those roles separately.

Portfolio search lives inside the full-width register. Click a stage to filter; Ctrl-click (Cmd-click on Mac) adds/removes stages. Clearing stage selections shows all projects, subject to the search and health filters. Empty portfolios omit the register and duplicate create prompt.

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
| `testingManagerPassword` | `admin123` | Testing-only self-promotion password; use `false` to disable. |
| `forceLocal` | `false` | Force clean-slate browser storage for local troubleshooting. |
| `forceSharePoint` | `false` | Force REST mode for an on-premises or custom SharePoint host. |
| `hideLists` | `true` | Hide the six backing lists from the normal Site Contents view without affecting API access. |

On first SharePoint load, the app silently creates empty lists and missing fields while the loading screen is displayed. Newly created backing lists are hidden from the normal Site Contents view in their initial create request, so startup does not issue follow-up MERGE operations. The signed-in user needs permission to create lists and fields for that first run; normal use needs whatever read/edit rights the site owner grants. The saved Users directory determines application roles; Every new identity, including site administrators, defaults to standard-user app access. The testing password flow can grant the signed-in user a persisted Manager role. SharePoint permissions remain the server-side access boundary; see the deployment guide before granting list access.

See [SHAREPOINT_DEPLOYMENT.md](SHAREPOINT_DEPLOYMENT.md) for the list schema and deployment behavior.


## September phase and progress update

Pipeline order is Requirement (MSA), Acquisition (TMRR), Procurement (EMD), Deployment (P&D). Both legacy and DAWIA terms appear together. Existing Development and Production/Procurement tasks map to Procurement (EMD); Operation and Sustainment maps to Deployment (P&D). Records, titles, owners and deadlines are retained; managers can move individual tasks if needed.

Project owners and managers can choose **Progress calculation** in the project drawer. Phase mode (default) counts phases whose existing tasks are all resolved; empty phases are not resolved. Task mode is exactly Complete tasks divided by **all existing tasks**, independent of their phase; Not Required is not counted as Complete. A project with no tasks reports 0%. Owners can add tasks to their own projects early; assignment, due dates, project creation, and changes to existing task definitions remain manager-controlled.

Needs attention shows only unfinished projects with outstanding tasks. Complete, Not Required, and legacy Not Applicable tasks are excluded even when they retain old deferral information.
