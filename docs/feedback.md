# Private bug reports and feature requests

The bug icon next to the theme control opens feedback for every Tracker role. **Feedback History** lets site owners and Tracker Managers read and edit submitted bug reports and feature requests. Site owners can also delete individual reports or clear the history. Viewers and Project Engineers can submit reports privately.

Feedback uses a dedicated `Modernization-Tracker - Feedback` SharePoint list (or `<customPrefix> Tracker - Feedback` for a custom namespace), separate from portfolio loads, exports, and migration backups. The HTML deployment cannot write to SharePoint until a signed-in owner performs setup.

## Set up from the application

1. Deploy the updated HTML and open the feedback icon while signed in as a site owner with **Manage Lists** and **Manage Permissions** on the hosting web.
2. When the list is missing, **Set up private feedback** appears automatically. No list changes occur just from opening the dialog.
3. Check the selected submission groups. Include all groups containing tracker users. Scoped deployments preselect existing `Tracker Users`, `Tracker Project Engineers`, and `Tracker Managers` groups. Legacy deployments also preselect the associated Members group. Other existing SharePoint groups can be selected for additional users.
4. Select **Create private feedback list**. The application creates the list, breaks inheritance, creates the fields, and verifies the list grants. The associated Owners group receives Full Control. Submission groups receive Add Items and the required API/page access. The Tracker Managers group receives permission to view and edit reports. The temporary direct creator grant is removed.
5. On success, the dialog reports that the list is ready in **Site Contents** and enables the inbox and submission controls. The list is visible in Site Contents to accounts with access, while ordinary users cannot read its contents.

For a deployed list, open the feedback dialog as a site owner and select **Enable manager history access**. This updates the existing list in place. It verifies the status column and grants review access specifically to the SharePoint **Tracker Managers** group. In deployments that store manager roles in the application directory, ensure those Managers also belong to the SharePoint group for this list.

If setup is interrupted, reopening feedback offers **Finish setup** for an app-created pending list. Only an empty list bearing the app's setup marker can be resumed; existing populated or unrelated lists are never overwritten. An unsafe pre-existing submission role is rejected. Access-denied responses are not treated as missing storage. If the creator-grant removal or another late step fails, inspect the existing list with a site administrator before collecting feedback.

Owner startup checks list existence, unique permissions, required fields, and reading permissions. Missing storage produces the setup panel. A malformed or non-private existing list produces an actionable error rather than being silently recreated. Ordinary users make no setup or inbox requests; a failed submission retains their draft and explains that an owner must finish setup.

The original `scripts/setup-feedback.ps1` remains an optional administrator alternative for a new list. It is no longer required for normal setup through the application.

## Deployment acceptance checks

Submit a report as a Viewer and a Project Engineer. Verify that both accounts can submit and that SharePoint denies their requests to read or edit reports. As a non-owner Tracker Manager, open **Feedback History**, edit the Viewer’s report and check the saved status. Verify SharePoint denies the Manager a delete. As an owner, confirm a different report can be deleted individually.

Automated tests exercise provisioning and permission checks against a simulated SharePoint API; they do not establish live tenant behavior. If the tenant rejects add-only submission, keep the list private and use a server-side submission endpoint instead of granting read access.

Local preview uses disposable module memory, labels submissions accordingly, and does not provision SharePoint storage. Reloading clears preview submissions.

## Embedded-host submission and list naming

Send feedback uses a direct click handler, like the directory’s Save action, so embedded hosts that cancel native form activation do not swallow the send operation. Empty fields show validation errors, pending setup shows a status message, and failed writes keep the draft.

The list uses the same App name - Content convention as the other tracker lists. Opening feedback as an owner with Manage Lists renames an existing `ModernizationFeedback` or `NPSL Tracker - Feedback` list in place by GUID after privacy checks. It preserves records, URLs and permissions. Duplicate aliases stop resolution instead of creating a replacement. Add-only users submit by display title with a legacy fallback only on an explicit CSOM list-resolution error before a write; they do not read the list catalog. Feedback remains excluded from shared portfolio loading and permission repair.

## Managing individual submissions

In **Feedback History**, Managers and owners can change the type, summary, details and status (**New**, **In progress**, **Resolved**, **Closed**). **Save changes** updates that report. **Cancel edit** discards the draft. History retains the original submitter and submission date.

Existing lists remain readable without a schema migration: reports without a status display **New**. The first owner save adds and verifies an optional `FeedbackStatus` choice column in the existing Feedback list, requiring Manage Lists for that initial upgrade. New lists include it during setup. The upgrade preserves existing records, naming and permissions. It never recreates the list or resets its ACLs.

Each update rechecks the reviewer’s access and uses the record’s SharePoint ETag to detect concurrent edits. Failed saves preserve the draft and display an error. A site owner can use **Delete** to confirm one report by title and move it to the SharePoint recycle bin.

## Clearing feedback history

Site owners can clear the displayed submissions after confirming in the dialog. Tracker moves those items to the SharePoint recycle bin. A report submitted after the history was loaded stays in the list. Site owners can restore recycled items through SharePoint.

Owner permission is rechecked when clearing begins. Ordinary users and non-owner tracker Managers cannot clear submissions. A partial failure stops the operation and reports how many items were cleared; uncertain writes are not automatically retried. Reopen the inbox to see the current state before retrying.

## Add-only submission transport

Submission uses SharePoint CSOM `/_vti_bin/client.svc/ProcessQuery` with AddItem, SetFieldValue and Update. It sends the signed-in user's cookies and the site's form digest, and never loads or queries the created item. The REST `/items` create endpoint returns the created item; that read-back can fail for the intentionally add-only submission role, even after the item was written. The site owner confirmed that the report affected by the HTTP 500 “Item does not exist” error was present in the inbox. Creation succeeded despite the failed confirmation. The write-only transport avoids returning that protected item; the new transport still requires live tenant validation.

The application confirms success only when CSOM returns an explicit `ErrorInfo: null`. Errors inside an HTTP 200 response are failures too. Only an explicit missing-list error during GetByTitle can try a legacy name. Timeouts, HTTP errors and generic “Item does not exist” responses never retry creation automatically. Before resubmitting a report affected by the old error, ask an owner to check whether it already arrived. No read permissions are added to submission groups.

Protocol reference: [Microsoft CSOM update example](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/ec9d185e-97c4-49e6-bbb3-8de04445db47). Automated coverage checks the request, digest/cookies, response handling and no-retry behavior; it does not replace submitting as an actual add-only SharePoint user.

## Tracker Users access group

Scoped feedback setup uses Tracker Users, Tracker Project Engineers and Tracker Managers. Keep **ModernizationFeedback Submit Only** for Tracker Users. Use the feedback dialog’s manager access action for review permissions on an existing list. Project permission repair operates separately.

## Screenshots and documents

Every user can optionally choose or drop files into feedback, or paste a screenshot into the form. Image previews and file names appear before sending, with a Remove action for each file. A submission can contain up to **5 files, each at most 1 MiB (1,048,576 bytes)**. Supported formats are PNG, JPEG, GIF, WebP, PDF, DOC/DOCX, XLS/XLSX, PPT/PPTX, TXT, LOG, CSV and ZIP. HTML, SVG and executable files are not accepted. These limits apply to the actual file bytes; they keep each CSOM request below the documented base64 upload limit.

After deploying, a site owner opens feedback, completes **Enable manager history access** if offered, then selects **Enable private attachments**. This adds an optional `FeedbackAttachments` plain-text Note column to the existing Feedback list and creates **Modernization-Tracker - Feedback Attachments**, a separate private document library. The library is created under a setup-only name and published under the upload name only after privacy and schema verification. It follows the same application naming convention as the lists. The action verifies the Feedback permissions, copies the approved submission and owner grants, gives verified Tracker Managers an attachment review role that also includes Open Items for document downloads, verifies the new library has unique permissions, and removes the temporary creator grant. It does not broaden access to other applications, documents or reports. Native attachments on the Feedback list stay disabled because adding one to an existing list item requires edit permissions that ordinary submitters deliberately do not have.

Existing text-only reports and submissions work before this upgrade. If an ordinary user selects files before setup, the message tells them to ask an owner to enable private attachments, or remove the files and send text feedback.

Files use collision-resistant generated names, are created with `Overwrite=false`, and inherit the private library's permissions. Ordinary submitters receive the same **Submit Only** permission level as Feedback, without View Items, Edit Items or Delete Items. Owners retain Full Control. Tracker Managers receive the dedicated Attachment Review role (the Feedback Review permissions plus Open Items), without Delete Items. Files remain separate from project and update documents, shared portfolio loads, exports and backups.

Uploads use CSOM `Files.Add(FileCreationInformation)` through an ObjectPath action without loading or querying the returned File. The report is created **after all files are confirmed** and stores only their names, sizes, safe MIME types and immutable storage keys. The application never automatically repeats an uncertain create. Confirmed uploads are reused if the same draft is explicitly retried. An uncertain file upload stops the operation and requires owner investigation; no report create is attempted. If files uploaded but the report was not confirmed, the same draft cannot resend those files. Ask an owner to check Feedback History before starting a new report. This prevents duplicate reports sharing the same file keys. SharePoint requests are not transactional: orphan files can remain after interrupted or abandoned submissions, and an owner can remove these from the private library after checking the report manifests. Closing or reloading the page does not preserve upload receipts.

Feedback History shows each report's attachments. Only verified owners and Tracker Managers can preview raster images or download files; each download rechecks the user's role and the report's manifest before reading the protected library. File links stay in the current SharePoint web. Files are downloaded as files, rather than rendering documents or HTML inside the application.

Owner deletion and clearing recycle each linked file **before** recycling its report. If a file fails, the report and its manifest remain and the error explains the partial operation. An owner can restore previously recycled files and reports together from SharePoint's recycle bin. Managers cannot delete or clear reports or files through Tracker.

### Additional tenant acceptance checks

1. Enable private attachments as a site owner. Confirm the new library has unique permissions: the approved owner and submitter grants plus the manager Attachment Review grant; Viewers and Project Engineers must have Submit Only, not Read/Contribute.
2. As a genuine Viewer, submit a small screenshot and a PDF. Confirm success once, then verify that direct SharePoint reads, downloads of a known file URL, edits and deletes are denied to that account. Repeat submission as a Project Engineer and a Manager.
3. As an owner and a non-owner Tracker Manager, open Feedback History, preview the screenshot, and download both files. Verify Managers cannot delete/clear, and that no attachment appears under project documents or project updates.
4. Delete one report as an owner and verify both the report and its files are in the recycle bin. Restore both if testing recovery. Verify clearing keeps a report submitted after History was loaded.
5. Test a rejected file, an upload failure, and a report response interrupted after upload. Confirm the draft remains and there is no automatic second report or file upload.

Automated tests verify the request protocol, allowed file paths, permission gates, upload receipts and error handling against simulated APIs. They cannot establish the tenant's add-only document-library behavior. Validate with the actual Viewer account before rollout; **do not add View/Edit permissions as a workaround** if the tenant rejects the no-readback upload. A server-side submission endpoint is the alternative in that case.

References: [FileCreationInformation](https://learn.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.client.filecreationinformation?view=sharepoint-csom), [CSOM ObjectPath response](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csom/18661c4a-8d59-4eb2-8742-786c7a674d25), [SharePoint file upload limits](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/upload-large-files-sample-app-for-sharepoint).
