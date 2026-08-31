# SharePoint and Forge deployment

## Deployment artifact

Run:

```bash
npm ci
npm test
npm run build:singlefile
npm run verify:singlefile
```

Upload or ship:

`build-singlefile/modernization-project-tracker.html`

The `singlefile.yml` GitHub Actions workflow performs the same sequence automatically on every push to `main`, smoke-tests the actual `iframe srcdoc` runtime, uploads a workflow artifact, and publishes the HTML plus SHA-256 checksum as a GitHub release.

The output contains the application JavaScript, CSS, workbook seed data, and vendored Forge runtime in one HTML file. It has no external scripts, stylesheets, fonts, or images.

The page must be opened from the SharePoint site that owns its lists, or from a same-origin embedded SharePoint page. Opening the file from a local disk cannot use Microsoft 365 authentication.

## Authentication

The app does not collect a username or password. It uses the current Microsoft 365 session and resolves the signed-in user through:

`/_api/web/currentuser`

Writes use a time-limited form digest from:

`/_api/contextinfo`

## Lists

With the default prefix, setup creates:

| List | Purpose |
| --- | --- |
| `ModernizationProjects` | One portfolio row per modernization project. |
| `ModernizationTasks` | WBS tasks linked by `ProjectKey`. |
| `ModernizationUpdates` | Status updates and decisions. |
| `ModernizationRisks` | Risks, issues, ownership, and mitigation. |

Provisioning is idempotent and additive. It creates missing lists and fields but does not delete, rename, or retype existing fields.

### Project fields

`RecordId`, `ProjectKey`, `MeasurementArea`, `Description`, `OwnerName`, `OwnerEmail`, `ManagerName`, `ManagerEmail`, `Priority`, `Health`, `ProjectStatus`, `CurrentStageKey`, `PercentComplete`, `TargetFinish`, `NextMilestone`, `NextMilestoneDate`, `SourceNotes`, `ImportedBaseline`, and `TagsJson`.

### Task fields

`RecordId`, `ProjectKey`, `WBS`, `TaskTitle`, `PhaseKey`, `SortOrder`, `TaskStatus`, `StartDate`, `FinishDate`, `DueDate`, `OwnerName`, `OwnerEmail`, `Notes`, `BlockedReason`, `SourceStartLabel`, and `DataIssue`.

### Update fields

`RecordId`, `ProjectKey`, `UpdateType`, `Summary`, `EntryDate`, `AuthorName`, and `AuthorEmail`.

### Risk fields

`RecordId`, `ProjectKey`, `RiskTitle`, `Severity`, `Probability`, `Mitigation`, `OwnerName`, `RiskStatus`, and `DueDate`.

## Access model

The app uses two complementary controls:

1. SharePoint permissions are the security boundary. Users can only read or write lists allowed by the site.
2. `managerEmails` controls whether the full portfolio navigation is shown. When it is empty, every site user has portfolio navigation so a first deployment cannot lock out administrators.

Engineers use **My work**, which filters projects and tasks by the current user's SharePoint email. They can claim an unassigned project or task. Managers use portfolio, pipeline, and project-register views.

If the portfolio requires strict separation, configure list permissions and item-level policies in SharePoint. Hiding navigation is not a security boundary.

## Workbook baseline import

The first-time setup can import 14 projects and 476 tasks. Writes run in small concurrent groups to avoid an unbounded burst against SharePoint.

Import is skipped when project rows already exist. To re-import in a test site, use a different `listPrefix` or remove the test lists through normal SharePoint administration.

The source workbook is historical. Imported projects are intentionally marked **Needs Review**. The app preserves source values, including fiscal-year starts, notes, and impossible start/finish orderings; those orderings display as data-quality warnings.

## Forge compatibility

The single-file build follows the same constraints as the existing Metrology Workbench deployment:

- Vite bundles React and all app code.
- `vite-plugin-singlefile` inlines JavaScript and CSS.
- the sanitiser-hardening build step escapes tag-like sequences inside JavaScript literals without changing their values;
- the Forge manifest, development console, and test recorder runtimes are injected during the build;
- runtime hashes are verified;
- the finished file has zero external subresource references.
