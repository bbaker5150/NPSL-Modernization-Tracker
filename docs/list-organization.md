# Tracker lists in Site Contents

Use **App name - Content** for each application's list titles. This tracker uses the following names so all seven sort together in Site Contents:

| Existing title | Visible title |
| --- | --- |
| ModernizationProjects | NPSL Tracker - Projects |
| ModernizationTasks | NPSL Tracker - Tasks |
| ModernizationUpdates | NPSL Tracker - Updates |
| ModernizationRisks | NPSL Tracker - Risks |
| ModernizationAcronyms | NPSL Tracker - Acronyms |
| ModernizationUsers | NPSL Tracker - Users |
| ModernizationReferenceDocuments | NPSL Tracker - Reference Documents |

## Existing site

1. Deploy the updated HTML and reload it as a site owner with Manage Lists permission.
2. In **Users and managers**, expand **Site Contents organization** and select **Organize tracker lists**.
3. Wait for verification of all seven lists, then refresh Site Contents and sort by name.

The action changes only each list's display title and Hidden flag. Existing IDs, URLs, records, attachments, folders, and permission assignments remain intact. Existing links to list settings by ID still work. Other apps' lists are not modified. Making a list visible does not grant access to it; SharePoint permissions still determine who can open or edit it.

The app resolves either naming convention to the existing list GUID before accessing data. Normal users only read the catalog; organizing requires Manage Lists. New installations create visible lists with the friendly names immediately. A custom `listPrefix` uses `<prefix> Tracker - <content>`; keep the existing prefix unchanged. The old `hideLists` option is ignored.

Before writing, organization checks all seven lists. If both names exist for the same list, it stops for the owner to resolve the duplicate; it never merges, deletes, or substitutes lists. If a request fails after some renames, rerun the action to finish. It verifies each list's ID, name, and visibility after updating it.

Deploy this build before organizing. Older HTML builds or external flows that look up a list by its old title must be updated to use the list ID or its new title. Do not roll back to a title-only build after renaming without restoring the old titles first. Backups, imports, and permission setup in this build support both names and use GUIDs, including batch writes.

## Acceptance checks

- Open Site Contents as an owner and confirm all seven names above are present.
- Reload the tracker; confirm existing projects, tasks, directory roles, and documents are still present.
- With a viewer account, confirm reading still works and editing is denied as before.
- With a manager or assigned engineer, confirm permitted edits and document access still work.
- Preview an export/import and run permission setup as needed; list renaming does not itself reapply permissions or send invitations.
