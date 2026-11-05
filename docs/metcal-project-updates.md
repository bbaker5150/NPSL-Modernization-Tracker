# METCAL responsibility and project updates

PMA 260B, CHENG Team and Program Office normalize to **NAVAIR METCAL 260B** on load. Existing IDs, records and assignments are retained. Saving a record writes the current name. METCAL tasks offer Chief Engineer Office, Team 2 Lead, or Team 3 Lead; blank responsibilities are shown as Not yet assigned. Switching away from METCAL clears the subassignment. Needs Attention shows responsibility prominently, supports responsibility filters and search, and carries the filter and responsibility into workbook exports.

Project Engineers can edit tasks on assigned projects, including deferred dates and their required justification. The task editor disables the assignment/creation date and original due date. The authorized repository also preserves those saved values on engineer writes, and generates the creation date for new tasks rather than accepting a supplied date. Managers can change them. These are application rules: SharePoint grants apply to whole items, not individual columns. Direct SharePoint editing by an account with item Edit permission is outside this field-level UI restriction.

## Updates

The project drawer has Overview, Work breakdown and Updates. The Risks tab is removed. Existing risk records appear in Updates as previously recorded risks, with their owner explicitly labeled; they remain in the Risks list and exports. A missing historical timestamp is not invented.

The simplified composer accepts an update, an optional task reference restricted to the same project, and optional documents. Existing Status, Risk, Decision and Handoff labels remain visible in history. Author identity comes from the current signed-in user; new updates retain an ISO UTC timestamp and display it in the reader's local timezone. Older date-only updates continue to show their recorded date. A task reference opens the task; a deleted task keeps its recorded name as a non-clickable reference.

Documents attach to the **Modernization-Tracker - Updates** list item, separate from Tasks/project documents and Reference Documents. They inherit that update's existing permissions. Local preview files use a separate `updates/<record-id>` namespace in IndexedDB. Failed file uploads retain only the unsuccessful files for retry against the already-posted update. Readers can download files; managers and the assigned Project Engineer can attach them. The 50 MB limit and filename checks match other tracker documents.

## Deployment and acceptance

Open the new HTML as a site owner with Manage Lists and Manage Permissions first. The existing additive setup adds Tasks.SubAssignment and Updates.TaskId, TaskTitle, PostedAt, EditedAt, EditorName and EditorKey to the existing lists, preserving contents and ACLs. Setup also enables attachments on the Updates list if disabled. Other users should open the app after the owner completes this upgrade. No new list or separate permission group is required.

Verify with a Project Engineer assigned to a project: assignment date and original due date are locked, deferred date/justification remain editable, and updates can be posted both with and without a task reference. Upload an update-only file, reopen the tracker, download it, and confirm it does not appear under project/task documents. Verify that an unrelated engineer cannot access that project's update files and a Viewer cannot upload. Check all three METCAL responsibilities plus Not yet assigned in Needs Attention.

Workbook exports include responsibility, update timestamps and task references. Backup/migration includes update documents separately from task documents; the existing attachment inclusion option now covers both. Use the updated tracker at both ends when transferring these fields/files. Automated API fixtures, backup round trips and embedded browser smoke tests cover these behaviors; live SharePoint acceptance remains necessary.

## Editing and deleting individual updates

Managers and the assigned Project Engineer can edit or delete entries regardless of who posted them. Editing a regular update changes its text and optional task reference. History retains the original author, posting time and documents and shows the last editor and edit time. Historical risks can be managed from the same tab and stay in their original Risks list. The repository rechecks the current role and project assignment for each action. Viewers have read-only access.

If SharePoint denies an edit to someone else’s entry, check both the affected item’s project permissions and the list’s **Advanced settings → Item-level Permissions**. An author-only write setting on Updates or Risks can block an assigned engineer who has item Edit access. Use the all-items setting with the existing project-specific grants when the team needs cross-author editing. [Microsoft SharePoint list security settings](https://learn.microsoft.com/en-us/sharepoint/dev/schema/security-element)

Save uses the version captured when editing began, plus SharePoint's ETag check, to reject stale drafts. Errors retain the draft. Copy it before reloading the tracker to resolve a conflict. Documents added later refresh the update version. Delete requires confirmation and archives only that update (`Archived=true`); its attachments stay on the archived SharePoint item and leave the project history together. Other updates and project/task documents are untouched. An administrator can restore the archived update by clearing its Archived field; it is not sent to the recycle bin.

Verify with a manager and the assigned engineer: edit one update, change/remove its optional task link, reload and check original author plus editor attribution and documents. Cancel deletion, then delete that update while another remains. Verify Viewer controls stay hidden and conflicting saves are rejected. Owner startup adds the three edit-attribution columns to the existing Updates list before other users open this version.

## Viewer startup compatibility

The edit-attribution fields (`EditedAt`, `EditorName`, `EditorKey`) are optional for reading an older Updates list. If they are missing, startup checks actual site permissions: a site owner upgrades the existing list; ordinary users continue without schema or permission writes. Reads omit only those confirmed-missing optional columns. Posting updates remains available to existing authorized writers, while editing is disabled until the audit columns exist. No update edit is saved without its attribution. Delete still uses the existing Archived field and ETag.

All other required columns, including Archived and identity/role fields, remain required. Missing or invisible lists and missing core columns produce a specific list/field diagnostic; denied reads remain errors rather than an empty portfolio. Feedback is excluded from startup requirements. The Viewer role and SharePoint ACLs are unchanged.

The regression test loads the built HTML in an iframe using a fixed pre-editing SharePoint schema snapshot and a Viewer account. It verifies real startup/API selection, read-only project history, no schema/ACL/record writes, and explicit failure for missing core fields, missing lists and denied item reads. This supplements local-mode browser tests, which cannot expose deployment schema mismatches.
