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

## Task documents and SME rollout

Task documents use the tasks list's native AttachmentFiles REST collection, not a separate library or file content in list text columns. New task lists are created with attachments enabled. On existing lists, a site owner must enable attachments in list Advanced settings if they are disabled; the app reports this condition rather than silently changing list settings. No extra file list or lookup-key synchronization is needed. Task deletion/recycling follows SharePoint's native item/attachment lifecycle. List attachments are suitable for task supporting documents; use a document library instead if a future requirement calls for independent document versioning, metadata, or approval workflows.

Uploads use the existing authenticated SharePoint session, form digest, and binary request body. List and item permissions govern read/write access. Standard users need appropriate SharePoint item-edit permissions to upload, even though the app permits them to change only selected task fields. The app role restrictions are not server-side ACLs; grant SMEs read access at SharePoint where practical. Shared testing-manager access remains a testing feature and should be disabled for production as described above. SMEs cannot use that feature to promote themselves.

Run the additive schema update with an account allowed to create fields before read-only SME users open the updated build. EstimatedHours is a Number column; existing tasks load with a blank estimate. AppRole is already text and needs no schema change to store SME.

The app does not overwrite or delete attachments. If an upload fails, successfully uploaded files remain attached; refresh the list before retrying uncertain outcomes. Files over the app's 20 MB limit or duplicate filenames must be renamed/reduced or managed separately. Local preview attachments stay in the browser's IndexedDB and are not uploaded to SharePoint.

REST reference: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest#working-with-files-attached-to-list-items-by-using-rest


### Prompt-free removal and documents
On first load after this update, the app's existing automatic provisioning adds `Archived` to tracker lists and `ArchivedDocuments` to Tasks (the first loader needs permission to add list columns). Existing records and attachments remain intact. Delete now archives records or hides document attachments through normal metadata updates, following Uncertalytics; it does not issue DELETE, MERGE, recycle, or bulk-delete requests. Archived content remains in SharePoint for site-owner recovery. Deleted document names remain reserved; rename a replacement upload. Downloads fetch the file bytes and save a local blob instead of navigating to SharePoint's file-opening page. Browser download policy and tenant security controls still apply.

Directory Save uses a direct click handler so it works when the embedded host blocks native form submission. SharePoint write errors remain visible in the form and roles are read back before success is shown.


### People Picker invitations (no Power Automate)
Managers can search organization people, select a resolved individual, choose a tracker role, and invite them. SharePoint People Picker and ensureuser resolve the identity; the tracker saves and verifies the directory role.

Every invitee joins the site's associated Members SharePoint group (ISEA METENG Members on this deployment). Existing members are not added again. Membership is read back, and effective site and tracker-page access must pass before exactly one page-targeted `SP.Web.ShareObject` email request. Flank Speed has retired `SP.Utilities.Utility.SendEmail`, so it cannot be used as a silent mail replacement. The app does not separately share the site or page; Members already supplies access. No Owners membership is granted. The inviter needs SharePoint permission to manage the group and request the invitation; application Manager status does not elevate SharePoint permissions.

All tracker roles, including SME, receive the Members group's underlying SharePoint permissions. The selected directory role controls tracker features only. Pages, documents, and lists accessible to Members are covered; content with unique permissions that exclude Members requires separate administration. Configure manager-only directory writes separately as appropriate.

Use the direct published tracker ASPX page as the email destination. Set window.MOD_TRACKER_CONFIG.appUrl to override detection. The known ISEA METENG deployment defaults to SitePages/Modernization Tracker.aspx; other deployments use available parent/current/referrer page hints. Site roots and Pages library views are rejected. The page must exist and be published; owners must publish or republish it themselves.

Email is requested only after membership and access verification. Delivery is not confirmed. A failure preserves the saved directory role and allows Retry site invitation. Reinviting a user with the same role does not duplicate the directory record. Existing roles are changed through Edit. Retrying can request another email. The invitation link is resolved internally; the link field and draft/copy controls are removed. Local preview cannot grant SharePoint access or send an automatic invitation.

Directory editing uses Edit → Update User / Delete User. Delete archives the tracker directory entry, not SharePoint membership or permissions. Self-deletion is disabled.


### Reference Documents
Reference Documents appears below My work for all roles. Users and managers can upload (20 MB per file), create nested folders, rename entries, and move them between folders. SMEs can browse, search, and download. Search spans the shared library. Downloads fetch bytes without opening a SharePoint document page.

The deployment adds the ModernizationReferenceDocuments list (or the configured prefix). Open the updated app once with site-owner/list-creation rights so automatic provisioning can create it and its fields. The list uses attachments for file bytes and metadata for folder organization; incomplete uploads remain archived and hidden. Moves and renames update metadata without copying document bytes. SharePoint permissions remain the server-side boundary; ensure Members can read/add/edit this list. Local preview stores metadata in localStorage and file bytes in IndexedDB.

The user directory now has compact scrollable rows and a name/email/role search while retaining each user's role badge and Edit menu.


Reference actions use a text-only Edit menu: Download, Rename, Move, and Delete (SMEs can only Download). New folder asks only for a name and uses the currently open folder. Delete archives the reference entry and preserves its underlying bytes in SharePoint; readback verifies removal. Nonempty folders must have their contents moved/deleted first. The local preview likewise removes directory metadata without purging stored bytes.
