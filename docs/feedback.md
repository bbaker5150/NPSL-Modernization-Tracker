# Private bug reports and feature requests

The bug icon immediately left of the theme control opens feedback for every tracker role. The bottom action row contains one **Send feedback** button, with **Owner inbox** immediately to its left for verified site owners. Site administrators, users with effective site-level Manage Permissions (the tracker’s existing owner definition), and members of the site's associated Owners group qualify. Owner access is rechecked whenever the dialog opens; verification failures are shown and can be retried by reopening it. Tracker Manager status alone does not grant access.

Feedback uses a dedicated `Modernization-Tracker - Feedback` SharePoint list (or `<customPrefix> Tracker - Feedback` for a custom namespace), separate from portfolio loads, exports, and migration backups. The HTML deployment cannot write to SharePoint until a signed-in owner performs setup.

## Set up from the application

1. Deploy the updated HTML and open the feedback icon while signed in as a site owner with **Manage Lists** and **Manage Permissions** on the hosting web.
2. When the list is missing, **Set up private feedback** appears automatically. No list changes occur just from opening the dialog.
3. Check the selected submission groups. Include all groups containing tracker users. Scoped deployments preselect existing `Metrology App User`, `Tracker Project Engineers`, and `Tracker Managers` groups. Legacy deployments also preselect the associated Members group. Other existing SharePoint groups can be selected for additional users.
4. Select **Create private feedback list**. The application creates the list, breaks inheritance before granting submission access, creates the fields, and verifies a custom submission-only permission level. The associated site Owners group receives Full Control. Selected groups receive Add Items and the API/page access permissions, without View Items, Edit Items, Delete Items, or Manage Permissions. The temporary direct creator grant is removed.
5. On success, the dialog reports that the list is ready in **Site Contents** and enables the inbox and submission controls. The list is visible in Site Contents to accounts with access, while ordinary users cannot read its contents.

Setup uses the current SharePoint identity and existing groups; it does not create a new login or grant site-wide permissions. The associated Owners group is the reading principal. An owner recognized through direct or security-group site permissions still needs membership/access through that Owners group to read the uniquely secured list; site collection administrators retain their administrative access.

If setup is interrupted, reopening feedback offers **Finish setup** for an app-created pending list. Only an empty list bearing the app's setup marker can be resumed; existing populated or unrelated lists are never overwritten. An unsafe pre-existing submission role is rejected. Access-denied responses are not treated as missing storage. If the creator-grant removal or another late step fails, inspect the existing list with a site administrator before collecting feedback.

Owner startup checks list existence, unique permissions, required fields, and reading permissions. Missing storage produces the setup panel. A malformed or non-private existing list produces an actionable error rather than being silently recreated. Ordinary users make no setup or inbox requests; a failed submission retains their draft and explains that an owner must finish setup.

The original `scripts/setup-feedback.ps1` remains an optional administrator alternative for a new list. It is no longer required for normal setup through the application.

## Deployment acceptance checks

Using separate Viewer, Project Engineer, and non-owner Manager accounts, verify both report types submit successfully, while direct SharePoint REST GET requests to the list's `/items` endpoint and individual items are denied. Verify authors cannot read their own items or edit/delete them, and that an owner can see all submissions in the inbox. Never grant Read, Contribute, or “read own items” to submission groups. Verify that the list does not inherit general site permissions.

Automated tests exercise provisioning and permission checks against a simulated SharePoint API; they do not establish live tenant behavior. If the tenant rejects add-only submission, keep the list private and use a server-side submission endpoint instead of granting read access.

Local preview uses disposable module memory, labels submissions accordingly, and does not provision SharePoint storage. Reloading clears preview submissions.

## Embedded-host submission and list naming

Send feedback uses a direct click handler, like the directory’s Save action, so embedded hosts that cancel native form activation do not swallow the send operation. Empty fields show validation errors, pending setup shows a status message, and failed writes keep the draft.

The list uses the same App name - Content convention as the other tracker lists. Opening feedback as an owner with Manage Lists renames an existing `ModernizationFeedback` or `NPSL Tracker - Feedback` list in place by GUID after privacy checks. It preserves records, URLs and permissions. Duplicate aliases stop resolution instead of creating a replacement. Add-only users submit by display title with a legacy fallback only on a definitive 404; they do not read the list catalog. Feedback remains excluded from shared portfolio loading and permission repair.

## Clearing the owner inbox

Inside the inbox, the redundant Owner inbox navigation button is hidden. **Clear inbox** opens an in-app confirmation for the submissions currently loaded. It clears the shared inbox for all owners by moving those list items to the SharePoint recycle bin; it does not delete the list or change its permissions. New submissions received after the inbox was loaded are not included. Site owners can restore recycled items through SharePoint’s normal recycle-bin controls.

Owner permission is rechecked when clearing begins. Ordinary users and non-owner tracker Managers cannot clear submissions. A partial failure stops the operation and reports how many items were cleared; uncertain writes are not automatically retried. Reopen the inbox to see the current state before retrying.
