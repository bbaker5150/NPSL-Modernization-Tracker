# METCAL responsibility and project updates

PMA 260B, CHENG Team and Program Office normalize to **NAVAIR METCAL 260B** on load. Existing IDs, records and assignments are retained. Saving a record writes the current name. METCAL tasks offer Chief Engineer Office, Team 2 Lead, or Team 3 Lead; blank responsibilities are shown as Not yet assigned. Switching away from METCAL clears the subassignment. Needs Attention shows responsibility prominently, supports responsibility filters and search, and carries the filter and responsibility into workbook exports.

Project Engineers can edit tasks on assigned projects, including deferred dates and their required justification. The task editor disables the assignment/creation date and original due date. The authorized repository also preserves those saved values on engineer writes, and generates the creation date for new tasks rather than accepting a supplied date. Managers can change them. These are application rules: SharePoint grants apply to whole items, not individual columns. Direct SharePoint editing by an account with item Edit permission is outside this field-level UI restriction.

## Updates

The project drawer has Overview, Work breakdown and Updates. The Risks tab is removed. Existing risk records appear in Updates as previously recorded risks, with their owner explicitly labeled; they remain in the Risks list and exports. A missing historical timestamp is not invented.

Updates support Status, Risk, Decision and Handoff, an optional task reference restricted to the same project, and optional documents. Author identity comes from the current signed-in user; new updates retain an ISO UTC timestamp and display it in the reader's local timezone. Older date-only updates continue to show their recorded date. A task reference opens the task; a deleted task keeps its recorded name as a non-clickable reference.

Documents attach to the **Modernization-Tracker - Updates** list item, separate from Tasks/project documents and Reference Documents. They inherit that update's existing permissions. Local preview files use a separate `updates/<record-id>` namespace in IndexedDB. Failed file uploads retain only the unsuccessful files for retry against the already-posted update. Readers can download files; managers and the assigned Project Engineer can attach them. The 50 MB limit and filename checks match other tracker documents.

## Deployment and acceptance

Open the new HTML as a site owner with Manage Lists and Manage Permissions first. The existing additive setup adds Tasks.SubAssignment and Updates.TaskId, TaskTitle and PostedAt to the existing lists, preserving contents and ACLs. Setup also enables attachments on the Updates list if disabled. Other users should open the app after the owner completes this upgrade. No new list or separate permission group is required.

Verify with a Project Engineer assigned to a project: assignment date and original due date are locked, deferred date/justification remain editable, and updates can be posted both with and without a task reference. Upload an update-only file, reopen the tracker, download it, and confirm it does not appear under project/task documents. Verify that an unrelated engineer cannot access that project's update files and a Viewer cannot upload. Check all three METCAL responsibilities plus Not yet assigned in Needs Attention.

Workbook exports include responsibility, update timestamps and task references. Backup/migration includes update documents separately from task documents; the existing attachment inclusion option now covers both. Use the updated tracker at both ends when transferring these fields/files. Automated API fixtures, backup round trips and embedded browser smoke tests cover these behaviors; live SharePoint acceptance remains necessary.
