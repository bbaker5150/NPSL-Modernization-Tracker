# Firepit 9.2 deployment

Build with `npm ci && npm run build:singlefile`, then publish `build-singlefile/modernization-project-tracker.html` through the existing Firepit/App Page deployment. Keep the existing list prefix and configured SharePoint web. Firepit's `window.firepit.sharepoint.siteUrl` supplies the exact hosting web, including subsites. A configured different web stops requests with an explicit error; the app never silently changes the data destination.

The bridge handles SharePoint authentication and form digests. Tracker sends REST operations through its `fetch` method and does not forward an app-supplied Authorization or X-RequestDigest header. Bridge failures retain their message; there is no direct-network fallback. Requests have a 60-second deadline. If a write times out, refresh and verify its result before retrying.

## Supported changes

- Current-user discovery, list/field queries, ordinary record writes, and attachment downloads use the hosting-web bridge.
- File downloads use the SharePoint file REST endpoint; the host still enforces actual file ownership and access.
- Group membership writes use `sitegroups/getbyid(...)`; form field updates set `bNewDocumentUpdate: false`.
- Project permission setup creates folders through the list's folder collection and verifies the resulting list item, folder identity, emptiness, parent project, and concurrency token before completing metadata or ACL changes.
- Themes tolerate opaque-frame storage restrictions. The final HTML is tested with an opaque sandbox, no direct network access, and no eval permission.

## Workflows that need SharePoint outside Firepit

- **Private feedback submission and first-time feedback/attachment provisioning:** unavailable in this Firepit build. Add-only submissions rely on a CSOM operation that the host blocks. These actions stop before uploading files or changing permissions; no Read permission is added to submitter roles. Use an approved SharePoint feedback form outside Firepit or contact a site owner. Existing authorized reviewer history and file downloads remain available.
- **Invitation email:** use the SharePoint page's Share dialog. Tracker's invitation action stops before role changes and before sending mail because the host blocks `SP.Web.ShareObject`. Continue managing Tracker roles separately in **Users and managers**.
- **Creating permission levels:** an owner must create the documented levels in SharePoint first. Existing verified levels can be assigned from Tracker. Glossary setup requires `<listPrefix>Acronyms Add Only` (default `ModernizationAcronyms Add Only`) with `Low: 200707, High: 32`: View/Add Items, View Forms, Open, View Pages, Use Remote APIs. Existing feedback manager-review setup requires `<listPrefix>Feedback Review` with `Low: 200711, High: 32`; see [feedback setup](feedback.md). Missing levels stop before the corresponding ACL changes.
- **Live transfer between sites:** export a Tracker backup on the source web, then open a fresh Tracker session hosted on the destination web and import that backup. A session must not read a foreign web and then attempt writes; Firepit can lock that session read-only. Backup/import does not transfer site permissions.

Lock Down Mode, a read-only Firepit session, and unavailable REST routes remain host policy errors. They are not repaired by granting broader SharePoint permissions. Reload into an authorized writable hosting session after resolving the host policy.

## Validation boundary

`npm test`, `npm run verify:singlefile`, and `npm run smoke:singlefile` cover request routing, rejected writes, cancellation, file bytes, folder identity safeguards, privacy guards, legacy behavior, and the production HTML in a mocked opaque Firepit frame. The smoke uses an independent bridge contract fixture, not Firepit implementation code. Real tenant permissions, host approval prompts, and live upload/write outcomes still require deployment verification.
