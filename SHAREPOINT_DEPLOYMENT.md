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

The output contains React, Excel export support, styles, the NAVAIR seal, and the vendored Forge runtime in one HTML file. It has no external scripts, stylesheets, fonts, or images.

## Authentication and access

The app uses a saved `ModernizationUsers` directory for application roles, without a separate sign-in. It uses the current Microsoft 365 session and resolves the user through `/_api/web/currentuser`. Writes use a time-limited form digest from `/_api/contextinfo`. Existing-item updates use SharePoint's validation endpoint, avoiding REST method-override MERGE requests and their repetitive host notifications.

All new identities default to User, including SharePoint site administrators and the local development identity. A site administrator still needs to perform the first upgraded load to provision missing lists/fields, but this no longer automatically grants the Manager app role. Existing saved Manager roles are preserved.

Every sign-in upserts the current SharePoint name, email, and login into `ModernizationUsers`, preserving its saved role. Standard users can open only My work, Acronym glossary, and Users and managers. The directory is visible to all users; only managers may edit other users or create assignments.

For the requested testing workflow, open **Users and managers**, enter `admin123`, and choose **Enable manager access**. The password applies only to the currently signed-in identity, and the Manager role is persisted in the directory. Configure `testingManagerPassword` before the bundle runs to override this default or set it to `false` to disable the flow. A disabled flow leaves existing roles intact. This password is bundled client-side and is not a production security control.

Testing self-registration and promotion require permission to add/update directory items under the current SharePoint session. If list permissions reject the write, the app displays the error and does not grant access. This change does not elevate SharePoint permissions. A production deployment must replace this convenience workflow with trusted registration/promotion and apply the list controls below.

Managers see the full application, create/assign projects and deadlines, maintain the directory, and add/delete phase tasks. Standard users see projects they own or have assigned tasks in; only their task rows are exposed when they do not own the project. Their repository writes are restricted to status and exception fields. Exports use the same scoped data. Project-owner changes also move tasks that still match the previous owner, while preserving separately assigned tasks.

### Required server-side security configuration

These roles enforce the application workflow, **not an independent server authorization boundary**. The single HTML application calls SharePoint directly under the signed-in user's existing permissions. It does not install a privileged backend or modify SharePoint ACLs. Anyone with direct list access retains that access outside the app; application filtering must not be used as a confidentiality guarantee.

Restrict writes to `ModernizationUsers` to trusted managers/site administrators so users cannot change their own role through SharePoint. Project, update, risk, and glossary list writes should likewise be manager-only. For confidential project separation, configure item-level SharePoint read permissions for the intended owners and managers. The existing Author-based "read own items" setting is insufficient because managers create projects on behalf of others.

SharePoint list permissions do not provide column-level write authorization: granting task edit access also permits direct edits to deadline/title columns outside this app. Strict enforcement of "status/deferral only" against direct API use requires a server-side write service or a separate submission-list workflow with automation, which this GitHub-only HTML change does not deploy. Validate tenant permissions before using these app roles as an organizational security policy.


## Lists

With the default prefix, the automatic clean-slate provisioning creates:

| List | Purpose |
| --- | --- |
| `ModernizationProjects` | One portfolio row per modernization project. |
| `ModernizationTasks` | Pipeline tasks linked by `ProjectKey`. |
| `ModernizationUpdates` | Status updates and decisions. |
| `ModernizationRisks` | Risks, issues, ownership, and mitigation. |
| `ModernizationAcronyms` | Shared acronym glossary. |
| `ModernizationUsers` | Saved identities and Manager/User application roles. |

Provisioning is idempotent and additive. It creates missing lists and fields but does not delete, rename, retype, or import example records.

New backing lists are marked `Hidden` in the same request that creates them. This keeps each tracker page from adding six more entries to the normal **Site Contents** view while preserving full REST access for the application and site administrators. Existing lists retain their current visibility so normal startup never performs a separate visibility MERGE. A site owner can still reach a hidden list by its direct URL or set `hideLists: false` before a new workspace is provisioned.

### Project fields

`RecordId`, `ProjectKey`, `MeasurementArea`, `Description`, `OwnerName`, `OwnerEmail`, `OwnerKey`, `ManagerName`, `ManagerEmail`, `Priority`, `Health`, `ProjectStatus`, `CurrentStageKey`, `PercentComplete`, `ProgressMode`, `TargetFinish`, `NextMilestone`, `NextMilestoneDate`, `SourceNotes`, `ImportedBaseline`, and `TagsJson`.

### Task fields

`RecordId`, `ProjectKey`, `TaskTitle`, `PhaseKey`, `SortOrder`, `TaskStatus`, `StartDate`, `FinishDate`, `DueDate`, `OwnerName`, `OwnerEmail`, `OwnerKey`, `Notes`, `BlockedReason`, `SourceStartLabel`, `DataIssue`, `DeferredDate`, `DeferredJustification`, and `NotRequiredJustification`.

Original due dates remain stored for deferred and Not Required tasks. Deferred dates must be later than the original due date and have justification. Not Required requires justification. Existing code columns remain inert in old lists (no destructive column deletion); the application no longer selects, writes, or exports them. Existing Not Required records without justification are labeled as legacy in the attention view and require justification when saved again.

### User directory fields

`RecordId`, `Title` (display name), `LoginKey`, `Email`, and `AppRole` (`Manager` or `User`).

### Update fields

`RecordId`, `ProjectKey`, `UpdateType`, `Summary`, `EntryDate`, `AuthorName`, `AuthorEmail`, and `AuthorKey`.

### Risk fields

`RecordId`, `ProjectKey`, `RiskTitle`, `Severity`, `Probability`, `Mitigation`, `OwnerName`, `OwnerKey`, `RiskStatus`, and `DueDate`.

## Forge compatibility

- Vite and `vite-plugin-singlefile` inline all application code and assets.
- Sanitizer hardening preserves JavaScript literals inside `iframe srcdoc`.
- Forge manifest, developer console, and test-recorder runtimes remain installed and hash-verified.
- The Forge developer and recorder buttons are hidden with the runtime's actual DOM selectors.
- The finished file has zero external subresource references.


### Phase/progress compatibility

The additive `ProgressMode` project field stores `phases` (default) or `tasks`. Owners may update only this project field through the dedicated repository method and may create tasks on their own projects with an initial Not Started status, no deadline, and themselves as owner. Manager-only project/assignment/deadline rules otherwise remain in place. In a tenant using restrictive list ACLs, these owner actions also require appropriate SharePoint write permissions; application code does not elevate permissions.

Legacy phase keys are normalized on load without deleting records: Development and Production/Procurement become Procurement (EMD); Operation and Sustainment becomes Deployment (P&D). Requirement and Acquisition retain their keys. Both browser storage and SharePoint use the same mapping.
