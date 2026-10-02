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

- Portfolio projects uses a searchable register. Needs Attention counts outstanding tasks by CHENG Team, NPSL, and NPSL Metrology Engineering. Organization is a task field; older tasks inherit their project's previous organization until edited.
- Managers have full project/task access and maintain users, roles, assignments, and original deadlines.
- Project Engineers read and edit only their assigned projects and related tasks, including documents, estimates, status, deferrals, and justifications. Project target completion and original task deadlines remain manager-only. The previous owner editing checkbox is retired.
- Viewers read all projects and documents without editing. Task assignment alone does not grant Project Engineer project access.
- Existing project owners with legacy User/SME roles become Project Engineers automatically, matched by login/email rather than display name. Existing Managers stay Managers; other legacy accounts become Viewers. A manager's next app startup persists directory migration; an owner's own startup persists their migrated role. Explicit later Viewer assignments are respected. Owners missing from the directory register as Project Engineers when they first open the app.
- Deferrals require a date after the original deadline and a justification. Not Required requires justification. The original due date is retained.
- Task codes have been removed from the UI, templates, exports, and active SharePoint schema.

## Testing manager access

New users land on **My work** and can also open **Acronym glossary** and **Users and managers**. Sign-ins automatically register the SharePoint name/email in the shared directory. Testing manager access is hidden by default. On **Users and managers**, press `Ctrl+M`, then enter `admin123` to promote the current signed-in identity; no separate username is needed. Existing saved manager roles remain in place.

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

### Status-focused review and progress views

- `In Progress – At Program Office` identifies work awaiting action outside NPSL. It does not override project health; managers set health independently. Project status is no longer exposed in the editor or drawer, and legacy status fields are retained only for storage compatibility.
- Needs attention groups outstanding tasks by status in independently collapsible sections. Resolved tasks remain excluded.
- The project progress chevron switches both the calculation and the markers between phases and tasks. Task mode labels the next outstanding task and shows one numbered marker per existing task; only Complete tasks receive a completed marker. Owners and managers save this preference; other assigned users can switch their own current view.
- The Portfolio projects Progress header has a separate view selector that compares every row using the same phase or task calculation without rewriting project preferences.
- Managers edit or delete a project through the vertical settings menu beside the drawer close button. Task creation uses the plus icon in each phase heading.

The brand returns to Portfolio home, and Active projects opens the pipeline board. Standard users continue to see only repository-authorized projects and tasks in these views. Needs attention uses horizontal, collapsible status columns with matching status pills. Task markers are distributed across a connected rail; long task lists scroll horizontally. Project menus dismiss on outside interaction or Escape. Managers can promote a directory entry with Make manager; the form uses name/email/role and retains SharePoint login identifiers internally.

Excel exports follow the current view and portfolio filters, preserve access scope, and include a compact Projects register, grouped Needs Attention, supporting detail sheets, and summary navigation links. SharePoint display names are shortened for reading without changing stored account identities. The pipeline selector centers each phase name above a neutral track and shows its project count in a circle below.

Needs Attention offers collapsible status or project groups, searchable task cards, and sorting by effective deadline, assignee, or hours. Notes and deferral justifications are collapsed by default. Complete and Not Required tasks are excluded. The task editor includes Assigned / creation date: managers and assigned project engineers may set it, new tasks default to today's UTC date, and legacy tasks stay blank until entered. This date is separate from Start Date and does not reset on reassignment. SharePoint provisioning adds AssignedDate and task Organization fields; schema-edit permissions are required for these additive updates.

### Task documents and estimated hours

- **Documents:** Select files while creating a task; saving creates the task and uploads queued files without reopening. Failed files remain queued for retry against the same task, without re-uploading successful files. Existing tasks upload immediately and refresh automatically. The shared limit is 50 MB per nonempty file. Native SharePoint task-item attachments preserve duplicate-name protection. Managers and assigned project engineers can upload, rename, download, and delete; Viewers can read/download. Local preview stores bytes in IndexedDB.
- **Task ownership and stage:** The inline saved-user picker stores name, email, and login identity together. Unresolved typed names cannot be saved. The stage comes from the phase where Add task was selected.
- **Est. Hours:** An optional nonnegative decimal; blank differs from zero. Managers and assigned project engineers can edit it. Attention cards and the Organization Workload, Needs Attention, and Tasks workbook sheets include task estimates, organizations, and dates. Export rechecks access and applies the current attention search/organization to the attention and workload sheets; supporting detail sheets cover the included projects.
- App roles govern tracker behavior. SharePoint ACLs remain the server-side security boundary; this change does not isolate site permissions or alter invitations.
