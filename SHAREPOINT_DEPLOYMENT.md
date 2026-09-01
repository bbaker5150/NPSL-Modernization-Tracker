# SharePoint and Forge deployment

## Deployment artifact

Run:

```bash
npm ci
npm test
npm run build:singlefile
npm run verify:singlefile
```

Upload `build-singlefile/modernization-project-tracker.html` to the SharePoint site that will own the tracker lists. The GitHub Actions workflow performs the same build on every push to `main`, smoke-tests the actual Forge `iframe srcdoc` runtime, uploads a workflow artifact, and publishes the HTML plus its SHA-256 checksum as a release.

The output contains React, Excel export support, styles, the editable sample dataset, the NAVAIR seal, and the vendored Forge runtime in one HTML file. It has no external scripts, stylesheets, fonts, or images.

## Authentication and access

The app has no separate sign-in or role database. It uses the current Microsoft 365 session and resolves the user through `/_api/web/currentuser`. Writes use a time-limited form digest from `/_api/contextinfo`. Existing-item updates use SharePoint's validation endpoint, avoiding REST method-override MERGE requests and their repetitive host notifications.

Every user who can open the HTML and access the lists receives the complete portfolio UI. SharePoint site/list permissions are the security boundary and determine who can read or edit data. The **My work** view matches the signed-in SharePoint `LoginName` stored in `OwnerKey`; email is used only as a compatibility fallback for older records.

## Lists

With the default prefix, the automatic clean-slate provisioning creates:

| List | Purpose |
| --- | --- |
| `ModernizationProjects` | One portfolio row per modernization project. |
| `ModernizationTasks` | WBS tasks linked by `ProjectKey`. |
| `ModernizationUpdates` | Status updates and decisions. |
| `ModernizationRisks` | Risks, issues, ownership, and mitigation. |

Provisioning is idempotent and additive. It creates missing lists and fields but does not delete, rename, retype, or import example records.

New backing lists are marked `Hidden` in the same request that creates them. This keeps each tracker page from adding four more entries to the normal **Site Contents** view while preserving full REST access for the application and site administrators. Existing lists retain their current visibility so normal startup never performs a separate visibility MERGE. A site owner can still reach a hidden list by its direct URL or set `hideLists: false` before a new workspace is provisioned.

### Project fields

`RecordId`, `ProjectKey`, `MeasurementArea`, `Description`, `OwnerName`, `OwnerEmail`, `OwnerKey`, `ManagerName`, `ManagerEmail`, `Priority`, `Health`, `ProjectStatus`, `CurrentStageKey`, `PercentComplete`, `TargetFinish`, `NextMilestone`, `NextMilestoneDate`, `SourceNotes`, `ImportedBaseline`, and `TagsJson`.

### Task fields

`RecordId`, `ProjectKey`, `WBS`, `TaskTitle`, `PhaseKey`, `SortOrder`, `TaskStatus`, `StartDate`, `FinishDate`, `DueDate`, `OwnerName`, `OwnerEmail`, `OwnerKey`, `Notes`, `BlockedReason`, `SourceStartLabel`, and `DataIssue`.

### Update fields

`RecordId`, `ProjectKey`, `UpdateType`, `Summary`, `EntryDate`, `AuthorName`, `AuthorEmail`, and `AuthorKey`.

### Risk fields

`RecordId`, `ProjectKey`, `RiskTitle`, `Severity`, `Probability`, `Mitigation`, `OwnerName`, `OwnerKey`, `RiskStatus`, and `DueDate`.

## Sample portfolio

The workbook-derived records are bundled only for **Open sample portfolio** mode. Users can edit projects and tasks, post updates, add risks, and practice deleting records; those changes remain in memory for the current page session and never call the SharePoint write API or browser persistence. Return to live data to work with actual projects.

## Forge compatibility

- Vite and `vite-plugin-singlefile` inline all application code and assets.
- Sanitizer hardening preserves JavaScript literals inside `iframe srcdoc`.
- Forge manifest, developer console, and test-recorder runtimes remain installed and hash-verified.
- The Forge developer and recorder buttons are hidden with the runtime's actual DOM selectors.
- The finished file has zero external subresource references.
