# Tracker lists in Site Contents

Use **App name - Content** for each application's list titles. This tracker uses the following names so all seven sort together in Site Contents:

| Existing title | Visible title |
| --- | --- |
| ModernizationProjects | Modernization-Tracker - Projects |
| ModernizationTasks | Modernization-Tracker - Tasks |
| ModernizationUpdates | Modernization-Tracker - Updates |
| ModernizationRisks | Modernization-Tracker - Risks |
| ModernizationAcronyms | Modernization-Tracker - Acronyms |
| ModernizationUsers | Modernization-Tracker - Users |
| ModernizationReferenceDocuments | Modernization-Tracker - Reference Documents |

## Existing deployments

Open the updated app as a site owner. After the app loads, it checks site-level Manage Permissions and Manage Lists before updating display titles and visibility. Other users do not write list metadata. Reload as an owner to retry a partial update.

All seven lists are resolved before the first rename. The original internal names and former `NPSL Tracker - …` display names remain accepted. Duplicate aliases stop maintenance instead of selecting an arbitrary list. Requests use the existing list GUID, preserving list URLs, records, attachments, folders and permissions. This is not a data migration.

The Site Contents organization, Tracker permissions and Site owner tools panels are no longer shown in the app. Role changes and project-engineer changes still synchronize permissions through the existing app workflows. This release does not rerun the retired full permission setup.

New lists use these names and are visible in Site Contents. Other applications and custom list namespaces are unchanged.
