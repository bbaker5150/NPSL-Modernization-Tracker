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

The output contains React, Excel export support, styles, the read-only mock dataset, the NAVAIR seal, and the vendored Forge runtime in one HTML file. It has no external scripts, stylesheets, fonts, or images.

## Authentication and access

The app has no separate sign-in or role database. It uses the current Microsoft 365 session and resolves the user through `/_api/web/currentuser`. Writes use a time-limited form digest from `/_api/contextinfo`.

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

### Project fields

`RecordId`, `ProjectKey`, `MeasurementArea`, `Description`, `OwnerName`, `OwnerEmail`, `OwnerKey`, `ManagerName`, `ManagerEmail`, `Priority`, `Health`, `ProjectStatus`, `CurrentStageKey`, `PercentComplete`, `TargetFinish`, `NextMilestone`, `NextMilestoneDate`, `SourceNotes`, `ImportedBaseline`, and `TagsJson`.

### Task fields

`RecordId`, `ProjectKey`, `WBS`, `TaskTitle`, `PhaseKey`, `SortOrder`, `TaskStatus`, `StartDate`, `FinishDate`, `DueDate`, `OwnerName`, `OwnerEmail`, `OwnerKey`, `Notes`, `BlockedReason`, `SourceStartLabel`, and `DataIssue`.

### Update fields

`RecordId`, `ProjectKey`, `UpdateType`, `Summary`, `EntryDate`, `AuthorName`, `AuthorEmail`, and `AuthorKey`.

### Risk fields

`RecordId`, `ProjectKey`, `RiskTitle`, `Severity`, `Probability`, `Mitigation`, `OwnerName`, `OwnerKey`, `RiskStatus`, and `DueDate`.

## Mock preview

The workbook-derived records are bundled only for the read-only **Preview mock portfolio** mode. Toggling the preview swaps the visible in-memory dataset and never calls the SharePoint write API or browser persistence. Return to live data to create and edit actual projects.

## Forge compatibility

- Vite and `vite-plugin-singlefile` inline all application code and assets.
- Sanitizer hardening preserves JavaScript literals inside `iframe srcdoc`.
- Forge manifest, developer console, and test-recorder runtimes remain installed and hash-verified.
- The Forge developer and recorder buttons are hidden with the runtime's actual DOM selectors.
- The finished file has zero external subresource references.
