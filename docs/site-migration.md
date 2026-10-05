# ISEA METENG → metsoft backup/import

Destination page: https://flankspeed.sharepoint-mil.us/sites/metsoft/SitePages/Modernization-Tracker.aspx

Use a downloaded backup package to separate reading the old site from writing the new one. The import never contacts the old site. A GitHub build does **not** move tenant data.

## Run

1. Close the old live migration tab before starting. Do not run two migration jobs at once, and do not delete the partial destination copy.
2. Merge this change and deploy the **same newly built single-file HTML on both sites**. Keep each site's own storage configuration: ISEA METENG on the source and metsoft on the destination, with the same tracker list prefix (normally `Modernization`). Open each app as a SharePoint owner. Complete normal storage setup on metsoft if prompted.
3. Pause edits in the old tracker. On ISEA METENG, open **Users and managers → Site owner tools · Export tracker backup**. Check the paused-edits box and select **Download backup**. Keep the downloaded `.npsl-backup.json` file. Export only reads the source; it does not modify it.
4. Pause edits in metsoft. On metsoft, open **Users and managers → Site owner tools · Import tracker backup → Choose backup file**. Select the downloaded package and choose **Preview import**.
5. Review new, already matching, and conflicting records. Previously migrated records are reused. The fresh site's seeded glossary and registered account may conflict; review the listed roles and explicitly approve backup values if appropriate. Check the paused-edits box, then select **Import and verify**.
6. Keep the page open on this view until completion. Download the verification report and reload. Compare project/task counts, open task/reference documents, and test with an intended standard user before distributing the new link. Keep the original site and backup until acceptance is complete.

The tools appear only on these two sites for an account with both **Manage Lists** and **Manage Permissions**. An app Manager role alone does not grant these SharePoint permissions.

## What the package contains

- Current app fields from Projects, Tasks, Updates, Risks, Acronyms, Users, and ReferenceDocuments, including archived records, app roles, assignments, application dates and archived-task attachment metadata.
- Binary task and reference attachments, including archived attachments, encoded in the same JSON file. Reference folder relationships, RecordId and ProjectKey are preserved; destination numeric SharePoint item IDs are newly assigned.
- A package checksum and SHA-256 hashes for individual files. Validation checks list completeness, field types, identities and project/reference relationships before import writes anything.

This package contains actual tracker data and documents in a readable, unencrypted format. Keep it in an approved location. The browser workflow supports up to **150 MiB of document contents** and a **256 MiB package**. Exceeding either limit stops export instead of downloading an incomplete package.

## Progress, recovery and verification

- Export reads each list with its attachment names, downloads the documents, and checks the list for changes during export. Keep source edits paused through the final cutover; changes made after export are not in the package.
- Import reads only the local package and metsoft. It writes records in batches of up to 25 and checks every operation's response; an HTTP 200 for the outer batch is not treated as proof that all records saved. SharePoint batches can partially succeed.
- Field values are verified with list-level reads, avoiding per-record verification requests. Each destination document is downloaded once for comparison with its backup hash during a normal successful import.
- Documents use the same authenticated direct file addresses as the tracker's normal Download action, obtained from SharePoint attachment metadata. They do not use the migration's former `AttachmentFiles/getByFileName(...)/$value` route. Expanded metadata supplies source addresses without a separate lookup per document; newly uploaded destination files resolve their addresses before verification. Failed downloads identify the document and item, and HTML sign-in pages are rejected instead of packaged as file content.
- Already matching records/files are reused. Destination-only records/files stay in place. Conflicting records require explicit approval; updates use ETags to detect concurrent changes. Different same-name documents stop the import and are never overwritten.
- A failed import can leave a partial copy. Keep it, correct the reported problem, and **preview the same package again** before importing. Do not delete records or repeatedly submit an old preview.
- Each network attempt has a 45-second deadline covering both headers and response body, even if the embedded host ignores cancellation. The UI shows the current request and elapsed time. Transient reads retry up to three times with backoff and short server-directed cooldowns. Longer cooldowns stop rather than retry early.
- A timed-out write is never blindly resent: the importer checks destination values or file hashes up to three times. If confirmation fails or only part of a batch saved, it stops and requires a fresh preview. An expired browser request does not roll back server work; let the request settle before resuming.
- Success is reported only after all package records and document hashes verify. The downloadable report contains counts, document names, sizes and hashes, not full record contents or file bodies.

Active users match by normalized login/email or RecordId; archived users match only by RecordId. Multiple active records for one account still require review rather than silently choosing a role.

## Separate from this transfer

SharePoint memberships, site/list/item permissions, sharing links, recycle-bin items, version history, system Created/Modified/Author/Editor fields, custom columns outside the tracker schema, and library files outside the tracker lists are not included. Existing URLs typed into notes remain unchanged. This is an application backup, not a full SharePoint site backup.

Deploy the HTML, ASPX wrapper and external app assets separately. Grant intended users appropriate permissions on metsoft. The embedded host must allow ordinary authenticated REST reads, batch writes and attachment uploads to the current site. The workflow does not bypass host confirmations or tenant access controls.

Automated tests use simulated SharePoint responses. Actual Flank Speed access and copied data must still be checked in the tenant before cutover.
