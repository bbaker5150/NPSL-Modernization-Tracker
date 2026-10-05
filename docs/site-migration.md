# ISEA METENG → metsoft migration

Destination page: https://flankspeed.sharepoint-mil.us/sites/metsoft/SitePages/Modernization-Tracker.aspx

The migration runs in the deployed tracker using the signed-in owner's SharePoint session. A GitHub build does **not** move tenant data.

## Run

1. Merge this change and deploy the newly built single-file HTML to the new site's app location. Keep the old tracker available and unchanged.
2. Confirm the deployed app's storage configuration points at `https://flankspeed.sharepoint-mil.us/sites/metsoft`, with the same list prefix used on ISEA METENG (normally `Modernization`). Check the wrapper/configuration too; copied configuration must not point at the old site. Open the destination as a SharePoint owner and complete its normal storage setup if prompted.
3. Open **Users and managers → Site owner tools · Migrate from ISEA METENG → Preview migration**. The tools appear only on metsoft for an account with both Manage Lists and Manage Permissions. The account also needs read access to every source tracker list and its attachments.
4. Review the counts and conflicts. Pause edits in both apps. A fresh site's seeded glossary and automatically registered user can appear as conflicts. Explicitly approve source values for conflicting destination records only after reviewing them.
5. Select **Copy and verify**, and keep the page open on this view until completion. Do not change data from other tabs while it runs.
6. Download the verification report, then reload. Compare project/task counts and open several task/reference documents. Test the new site with an intended standard user before distributing the new link. Keep the original site as the recovery source until acceptance is complete.

## Scope and behavior

- Copies the current application fields in Projects, Tasks, Updates, Risks, Acronyms, Users, and ReferenceDocuments, including archived records and archived task attachment metadata.
- Copies task and reference item attachments as binary data, including archived attachments. Reference folder relationships, RecordId, and ProjectKey remain intact; numeric SharePoint item IDs are newly assigned.
- Matches active users by login/email identity as well as RecordId to avoid duplicating the account registered on first destination launch. Archived users match only by RecordId, preserving deleted/recreated account history without blocking an active account. Multiple active records for one account stop the preview with the site, item IDs and roles to review; no role is chosen automatically. Invalid relationships also stop the preview.
- Reads all pagination pages. Verifies copied field values and SHA-256 file hashes, then checks that the source snapshot and document contents still match before reporting success.
- Never writes to the source or deletes destination data. Destination-only records/files are retained. Existing different record values require explicit approval. Existing updates use the retrieved ETag; different same-name attachments stop without overwrite.
- A failure can leave a partial copy. Correct the reported problem, preview again, and resume. Matching records/files are reused. Do not repeatedly use an old preview.
- Transient GET failures retry up to three times with backoff; short Retry-After cooldowns are honored and longer cooldowns stop without an early retry. List reads use pages of 500 records. After a timed-out write, the migration checks destination values/file hashes up to three times and continues if the write is confirmed. It never blindly repeats an uncertain write. If confirmation fails, refresh the preview before resuming. The error screen retains the last operation to help diagnose persistent host timeouts.
- The verification report contains record counts, source/destination addresses, attachment names, IDs, sizes, and hashes—not full record contents or file bodies. Store it appropriately.

## Separate from migration

SharePoint group membership, site/list/item permissions, sharing links, recycle-bin items, version history, system Created/Modified/Author/Editor fields, custom columns outside the current tracker schema, and library files outside the tracker lists are not copied. Application dates, author identity fields, roles, and assignments **are** copied. Existing URLs typed into notes remain unchanged. This is not a site backup or full-fidelity SharePoint export.

Grant intended users appropriate permissions on the new site separately. An app Manager role does not grant SharePoint permissions. The HTML, ASPX wrapper, and any external app assets must already be deployed on the destination.

The embedded host must permit ordinary authenticated reads from the ISEA METENG source and writes to the metsoft destination. A host-policy or SharePoint denial stops the migration; the feature does not bypass host confirmations or tenant access controls. It has automated tests with simulated SharePoint responses, but requires verification in the actual Flank Speed tenant.
