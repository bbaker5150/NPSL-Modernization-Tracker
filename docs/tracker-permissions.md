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

## Startup access errors and schema setup

Opening Tracker with Read access does not require site-wide Edit. Startup checks the seven lists and their fields. A list omitted from the signed-in user's catalog can appear missing even when it exists. Before creating lists, adding columns, or seeding schema-related records, setup now verifies actual site-level **Manage Lists and Manage Permissions**. Tracker Manager membership or list-level permissions alone do not authorize schema setup. Permission-check failures stop setup without writes. Administrators with both site permissions retain automatic first-load setup.

If a reader encounters an unavailable-schema message, an owner should first check that person's Read access to Projects, Tasks, Updates, Risks, Users, Acronyms, and Reference Documents, and confirm the configured site and list names. Open the app as an authorized site administrator only after confirming whether schema setup is actually needed. Do not grant viewers site-wide Edit to suppress an error.

This startup guard does not change role groups, list ACLs, or invitation checks. The existing invitation workflow still checks site-level View Items; migrating to page/folder/list-only access requires a separate change and tenant testing. The historical full permission setup UI described below is no longer exposed in the current app; do not assume its button is available.

## Manager and engineer permission checklist

1. Retain Owners with Full Control. Keep standard Read and Contribute unchanged.
2. Keep the exact group names above. Set Tracker Managers as owner of Tracker Viewers and Tracker Project Engineers. A site owner controls Tracker Managers. Restrict membership editing to group owners, disable join/leave requests, and ensure authorized managers can read the memberships needed by role synchronization.
3. Create or verify **Tracker Project Access Manager** by copying Contribute and adding Manage Permissions (plus dependencies SharePoint selects). Assign this custom level only on Projects, Tasks, Updates, and Risks, never to Tracker Managers across the whole site. This intentionally delegates permissions administration on these lists to trusted managers.
4. On all seven lists, establish unique permissions if needed, preserve Owners, and grant Viewers and Project Engineers Read. Grant Managers the custom level on the four project-content lists and standard Contribute on Users, Acronyms, and Reference Documents. Never grant the engineers group list-wide Contribute/Edit.
5. Give all three role groups Read on the published Tracker ASPX and its Site Assets folder; the HTML and dependencies can inherit from that folder. Homepage access is separately managed. Keep the existing site-access dependency in mind before removing site grants.
6. Use Tracker's normal Users and managers workflow to assign roles, then assign an engineer to a project. The app synchronizes direct Contribute on that project and its task/update/risk scopes. Existing deployments need their project folders/ACLs intact; report incomplete-setup errors rather than broadening engineer access. The legacy setup control is absent from the current UI.
7. Audit existing unique project/child scopes and other user memberships. Copied broad grants survive removal at a parent list. Preserve intended individual engineer grants while reviewing unwanted access.
8. Test as non-owner Viewer, Engineer A, Engineer B, and Manager. Verify Read, assigned-project changes and attachment operations, denial of edits to another engineer's project, role/assignment changes, and revocation after reassignment. Invitations require group-management and page-sharing rights and the existing site-level check. A successful UI test does not prove the direct REST permissions: verify those too.

Engineers still have baseline Read across all Tracker data; assigned-only visibility is an app filter. SharePoint Contribute allows modification of all fields of permitted records. This model does not hide permitted HTML files or list data from direct access.

## First deployment and imported records

1. Deploy the new HTML and open the tracker as a site owner. Site owners are recognized through actual SharePoint site permissions and can administer the tracker without the testing password.
2. Review the saved directory roles and project engineer assignments. Under **Users and managers → Tracker permissions**, choose **Apply tracker permissions**. This grants the three group memberships from the saved roles; it sends no emails.
3. The owner setup enables folders in Tasks, Updates, and Risks. Each project gets a folder in each of those lists. Folder names use the destination project item ID, so they are stable and cannot collide. Existing records and their attachments stay where they are; their item permissions are updated individually. New records are created inside the project's folder and inherit access.
4. Setup shows the current project/item. If interrupted, rerun it. Completed group changes and ACLs remain; each scope is inspected and reconciled again. Requests have a 60-second deadline. A timed-out write may have completed. Do not run setup concurrently with role/assignment changes.
5. Run setup again after importing records. Permission folders are excluded from record backups; SharePoint permissions and memberships are not migrated by backup/import.

Setup preserves existing group grants, Read, Full Control, and other manually assigned permission levels. On project items, their task/update/risk items, and project folders, **direct-user Contribute is reserved for the assigned engineer**. Setup removes other direct-user Contribute bindings on those scopes, including archived content. Review any manually maintained direct Contribute exceptions before applying.

## Routine behavior

- Invitations and saved role edits synchronize and verify tracker group memberships. Stale tracker-role memberships are removed. Errors stop the operation; the app does not report a role as saved if group/ACL changes fail. SharePoint operations are not transactional, so a partial change requires retry/setup.
- Viewer → Manager only updates Tracker groups and the verified directory role; it does not scan projects, create folders, or change task permissions. Saving an unchanged Viewer/Manager role also avoids project work. All three live group memberships are read first. If the account actually belongs to Project Engineers, the app still cleans up its old engineer grants even when its saved role says Viewer. Native confirmations remain for necessary group changes.
- Assigning/reassigning an engineer applies Contribute to that project's item, task/update/risk folders, and existing child records. Former engineer grants are removed. Task attachments inherit their task item's permissions. Managers assigned as engineers already receive access through the Managers group.
- Demotion/removal reads the four project-content lists and inspects their existing unique ACLs to find this user’s direct Contribute grants, including archived records and former assignments. It removes only matching grants, verifies each removal, and then updates tracker-group memberships. It never creates folders, breaks inheritance, adds grants or repairs other users during this cleanup. A Manager → Viewer change with no direct engineer grants needs only the two group membership writes. Read checks do not request permission changes. Deleting a user does not remove unrelated site/Microsoft 365 memberships or other manually assigned permission levels.
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

## Retrying after a host interruption

Firepit may suspend or reload the embedded app when you leave it. SharePoint retains each completed permission change. Run **Apply tracker permissions** again after a reload: setup reads the current child-item ACLs four at a time and skips uniquely permissioned tasks, updates, and risks whose direct engineer Contribute bindings already match. It reports **Checking**, **already correct**, and **updated** separately. Mismatches use the existing reconciliation and verification path; legacy items still inheriting from the list are prepared for project-specific access. Group grants and unrelated permission levels remain untouched.

Retries still scan live permissions from the beginning, but do not repeat the full update path for matching items. There is no local permission checkpoint or cached authorization decision: manual SharePoint changes and interrupted requests are checked on retry. Unreadable or malformed ACLs stop the current batch before mutations, rather than being treated as correct. A test reproduces interruption at item 131 and verifies that retry only updates items 131 and 132 when the first 130 already match.

Within the same loaded app, leaving and returning to Users and managers retains the running job and progress, and cannot start a duplicate job. This cannot prevent Firepit from suspending or destroying the app; after a host reload, start the live check again.

API references: [List items in folders](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest#create-list-item-in-a-folder), [REST permission assignment](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/set-custom-permissions-on-a-list-by-using-the-rest-interface).

## Recovering interrupted project-folder setup

If **Apply tracker permissions** stopped with **Project folder identity does not match**, deploy the folder-verification update and run the action again. Verification now finds the folder by its exact FileRef path in the backing list and reads the canonical list item by ID. It does not use GetFolderByServerRelativePath/ListItemAllFields, which returned no usable item ID in the reported deployment. It accepts either numeric or string `1` for a folder. A missing or malformed lookup response is not treated as an empty list.

Setup can finish initialization of an existing empty `tracker-project-<ID>` folder when its ProjectKey is explicitly blank. It checks the exact folder path, item ID, folder type, an explicit zero ItemCount from the associated Folder resource, and the saved parent project's identity before writing only the missing ProjectKey. The write uses the returned ETag and is read back before folder permissions are applied. Ordinary task creation cannot repair folder metadata. Non-empty folders, another project's key, missing identity evidence, and concurrent changes stop setup with the affected list and path in the error.

Folder verification remains part of project setup and assignment. Role-change cleanup, including Project Engineer → Manager and Manager → Viewer, uses targeted revocation of existing direct engineer grants instead of rerunning project setup. A clean Viewer → Manager promotion remains group-only. Cleanup preflights ACL reads before any mutation, rechecks planned removals, and completes before group changes; the directory role is saved only after access changes and group verification succeed. Interrupted changes reread live ACLs and skip completed removals. Firepit confirmations remain host-controlled: existing engineer grants may still require individual confirmed removals, but unrelated project setup writes no longer occur.

Completed permission changes remain in place; rerunning rechecks them. Do not delete imported records or broaden group permissions to work around this error. SharePoint's host confirmation prompts still apply. Automated regression tests cover the recovery paths; the specific live folder state must be verified by rerunning setup on metsoft.

The follow-up **Could not identify the tasks project folder** error indicates a missing folder item ID, not an incorrect list display title. The root URL comes from SharePoint list metadata; changing a display title to `NPSL Tracker - Tasks` leaves the existing `/Lists/ModernizationTasks` path intact. Do not rename or move that URL to match the display title.

The **PrimitiveValue / StartObject** HTTP 400 from `AddValidateUpdateItemUsingPath` was a folder-creation payload error: this API requires `LeafName` to be a ResourcePath object (`{ "DecodedUrl": "tracker-project-<ID>" }`), not a string. The corrected build uses that object for tasks, updates, and risks folders. Deploy the new HTML and rerun **Apply tracker permissions**; the existing lookup and identity checks reuse verified folders and retain completed work. HTTP-boundary regression tests cover serialization through the real store, list-ID resolution, and request helper, plus ordinary child-record creation. Live tenant verification is still required. See Microsoft's [LeafName property contract](https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-csom/mt796264(v=office.15)).

The **ItemChildCount does not exist** HTTP 400 came from requesting computed columns as list-item REST properties. Folder identity reads now select only Id, FileSystemObjectType, FileRef, and ProjectKey. Only recovery of an explicitly blank ProjectKey reads `items(<ID>)/Folder?$select=ServerRelativeUrl,ItemCount`; the returned folder path must match and ItemCount must explicitly be zero. Missing/null counts, other paths, populated folders, and failed reads stop recovery without changing the key. Correctly keyed folders need no count request. This uses the documented [list item's Folder](https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-csom/jj171077(v=office.15)) and [Folder.ItemCount](https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj246437(v=office.15)) properties.

HTTP tests now run creation, canonical identity verification, and a repeated setup call for Tasks, Updates, and Risks, with the reported unsupported-property error enforced by the mock server. These tests verify the request contracts; they do not replace a live tenant check. After deploying this update, run **Apply tracker permissions** again to resume.
