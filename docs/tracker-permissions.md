# Tracker access on metsoft

Scoped access is enabled automatically for `/sites/metsoft`. Other deployments can set `MOD_TRACKER_CONFIG.scopedAccess = true`. The old ISEA METENG deployment keeps its existing behavior for export/migration. No site memberships are changed during startup.

## Site configuration

Use these exact SharePoint group names. All three groups need **Read** on the site, the published ASPX page, and the HTML/assets it loads.

| Group | App role | Tracker list permissions |
| --- | --- | --- |
| Tracker Viewers | Viewer | Read on all seven lists |
| Tracker Project Engineers | Project Engineer | Read on all seven lists; individual Contribute grants on assigned projects and their content |
| Tracker Managers | Manager | Tracker Project Access Manager (Contribute plus Manage Permissions) on Projects, Tasks, Updates, Risks; Contribute on Users, Acronyms, ReferenceDocuments |

Tracker Managers owns the Viewers and Project Engineers groups. The site owner owns Tracker Managers. Allow group membership editing only by the group owner and disable join/leave requests. SharePoint enforces these settings: an ordinary tracker manager cannot promote someone to Manager unless also authorized to manage that group. The app never grants site Members membership in scoped mode.

Keep site Owners with Full Control. Do not give the entire engineers group list-wide Contribute.

## First deployment and imported records

1. Deploy the new HTML and open the tracker as a site owner. Site owners are recognized through actual SharePoint site permissions and can administer the tracker without the testing password.
2. Review the saved directory roles and project engineer assignments. Under **Users and managers → Tracker permissions**, choose **Apply tracker permissions**. This grants the three group memberships from the saved roles; it sends no emails.
3. The owner setup enables folders in Tasks, Updates, and Risks. Each project gets a folder in each of those lists. Folder names use the destination project item ID, so they are stable and cannot collide. Existing records and their attachments stay where they are; their item permissions are updated individually. New records are created inside the project's folder and inherit access.
4. Setup shows the current project/item. If interrupted, rerun it. Completed group changes and ACLs remain; each scope is inspected and reconciled again. Requests have a 60-second deadline. A timed-out write may have completed. Do not run setup concurrently with role/assignment changes.
5. Run setup again after importing records. Permission folders are excluded from record backups; SharePoint permissions and memberships are not migrated by backup/import.

Setup preserves existing group grants, Read, Full Control, and other manually assigned permission levels. On project items, their task/update/risk items, and project folders, **direct-user Contribute is reserved for the assigned engineer**. Setup removes other direct-user Contribute bindings on those scopes, including archived content. Review any manually maintained direct Contribute exceptions before applying.

## Routine behavior

- Invitations and saved role edits synchronize and verify tracker group memberships. Stale tracker-role memberships are removed. Errors stop the operation; the app does not report a role as saved if group/ACL changes fail. SharePoint operations are not transactional, so a partial change requires retry/setup.
- Assigning/reassigning an engineer applies Contribute to that project's item, task/update/risk folders, and existing child records. Former engineer grants are removed. Task attachments inherit their task item's permissions. Managers assigned as engineers already receive access through the Managers group.
- Demotion/removal reconciles all projects to catch grants left by an interrupted reassignment, then removes stale tracker-group memberships. This can take longer than a simple directory edit. Deleting a user does not remove unrelated site/Microsoft 365 group memberships or manually assigned permission levels.
- The current user's effective app role comes from their tracker group membership (or site-owner permissions), not a writable AppRole string alone. Read-only users don't write to the Users list at startup. Testing password promotion is disabled in scoped mode, including direct calls to its handler.
- Engineers can delete tasks/content but project deletion is manager-only in scoped mode because it revokes engineer ACLs. Engineers cannot change engineer assignment through the app. Project deadlines and other manager-only fields retain application restrictions. **SharePoint item permissions do not provide column-level security**: Contribute can edit the whole assigned record through SharePoint APIs. A trusted backend or separately secured records would be needed to enforce field-level restrictions outside the app.
- Reference documents remain readable/downloadable for all roles and writable by managers.

## Email and final verification

The invitation retains the supported page-targeted `SP.Web.ShareObject` request and its native confirmation. No retired SendEmail endpoint or HTML message markup is used. Group membership and email confirmation are separate: an account allowed to manage a group may still lack permission to share the page. In that case membership can succeed while the invitation reports an error; a site owner must send/retry the page invitation. This update does not grant managers extra site/page sharing rights.

Before removing transitional access, test with non-owner accounts:

| Account | Expected behavior |
| --- | --- |
| Viewer | Open page/HTML, read tracker and reference documents; cannot save records |
| Assigned engineer | Edit assigned project; create/edit task, update, risk and task attachment; cannot edit another project |
| Tracker manager | Manage tracker records and invite Viewer/Engineer; manager promotion follows group ownership |
| Former engineer | After reassignment/demotion, cannot edit the old project's records through the app **or direct REST** |

Existing broad access must be reviewed separately. Permissions are cumulative. Users who remain site Members (including through a Microsoft 365 group), Owners, or recipients of unrelated direct write grants may still edit broadly. Remove those user memberships/grants after testing. **Breaking inheritance copies existing group grants into project scopes**: removing a Members grant only from the parent list will not remove copies on child scopes. Remove the users' broad group membership or audit every unique scope during cutover. This update intentionally does not delete existing site groups or their permissions.

Live Flank Speed group ownership, folder behavior, page-sharing policy, and effective user permissions must be verified on the tenant. Automated tests use mocked SharePoint responses; they are not evidence of live access enforcement.

API references: [List items in folders](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest#create-list-item-in-a-folder), [REST permission assignment](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/set-custom-permissions-on-a-list-by-using-the-rest-interface).

## Recovering interrupted project-folder setup

If **Apply tracker permissions** stopped with **Project folder identity does not match**, deploy the folder-verification update and run the action again. Verification now reads the underlying list item by ID rather than relying on custom metadata in the folder-navigation response, and accepts either numeric or string `1` for a folder.

Setup can finish initialization of an existing empty `tracker-project-<ID>` folder when its ProjectKey is explicitly blank. It checks the exact folder path, item ID, folder type, zero child items/folders, and the saved parent project's identity before writing only the missing ProjectKey. The write uses the returned ETag and is read back before folder permissions are applied. Ordinary task creation cannot repair folder metadata. Non-empty folders, another project's key, missing identity evidence, and concurrent changes stop setup with the affected list and path in the error.

Completed permission changes remain in place; rerunning rechecks them. Do not delete imported records or broaden group permissions to work around this error. SharePoint's host confirmation prompts still apply. Automated regression tests cover the recovery paths; the specific live folder state must be verified by rerunning setup on metsoft.
