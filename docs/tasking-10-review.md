# Modernization tracker tasking review

Reviewed against the seven requirements in Mod Tasking(10).docx, with the subsequent instruction to replace the separate project-owner selector with inline saved-user autocomplete. This review covers repository behavior and automated tests; it does not certify the deployed Flank Speed tenant.

| Requirement | Implementation and verification |
| --- | --- |
| 1. Project documents | Below Upcoming work; Edit menu offers Download, Rename and Delete to managers and authorized owners. Component coverage checks ordering, menu contents and rename. SharePoint-store coverage checks copying original bytes before archiving the old attachment. |
| 2. Work breakdown paperclip | Appears immediately before the phase n/n count. Its hover text and accessible description now identify each file and its task. Project-view document changes refresh this data. |
| 3. Organization | Program Office, NPSL and ISE options persist through the project store. Cards display organization; pipeline-board organization sorting is tested. Legacy Program Office task status is normalized to In Progress. |
| 4. Organization attention cards | Replaces Portfolio Progress with three organization-specific attention cards. Tests open each card and check its heading and project scope. |
| 5. Upload limit | Task and reference uploads share the 50 × 1024 × 1024-byte limit. Tests accept the boundary and reject one byte above it, empty files and invalid sizes. Actual tenant upload limits and large-file network behavior require deployed testing. |
| 6. Table sorting | All seven data columns are tested ascending and descending. Text uses alphabetical order, health uses the defined category order, stage uses pipeline order, and progress is numeric. |
| 7. Roles and owner editing | User and Manager roles; legacy SME maps to User. Standard task edits now preserve status and task-definition fields, allowing deferred dates and justifications. Task attachment upload remains available as previously requested. Checked owners can manage their own project tasking; revocation, protected ownership controls and unrelated-project denial are tested. |

## Corrections from this review

- Removed standard-user task-status editing and disabled the quick-complete control for users without project-edit access.
- Replaced the generic paperclip message with task/file details and refreshed it after document mutations.
- Scoped project document listings to the tasks the caller can access.
- Made the portfolio tolerate a missing owner name.
- Expanded tests beyond Project-column sorting to every column, each organization attention card, board sorting, document menus, owner revocation and exact upload-size boundaries.
- Included the owner autocomplete fix: selecting a saved person stores their name, email and login identity together; unresolved text cannot be saved. The compact permission checkbox sits below the field.

## Deployment limits and decisions left unchanged

Application checks execute in the browser. SharePoint ACLs are the server-side security boundary. Hiding lists or hiding the testing-manager control does not secure them. The requested Ctrl+M testing flow remains available; disable `testingManagerPassword` for production and review previously granted manager roles.

Existing projects without Organization default to NPSL. The old Program Office task status does not reliably identify a project's current organization, so this review does not bulk-reclassify existing projects. Managers should review their organization values.

Attachment deletion currently archives the name in application metadata rather than physically deleting the SharePoint bytes. Renaming copies bytes to a new attachment and then archives the previous name. This is not transactional: an interruption between steps can leave both names visible. No server-side permission changes or physical deletion changes were made as part of this review.

The separate request to isolate the tracker from the wider ISEA site remains deferred. No site memberships, invitations or live tenant data were changed during this review.
