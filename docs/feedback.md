# Private bug reports and feature requests

The bug icon immediately left of the theme control opens feedback for every tracker role. Site administrators and members of the site's associated Owners group also see **Owner inbox**. Tracker Manager status alone does not grant access. Feedback uses a separate `<listPrefix>Feedback` list, never the tracker data loader, Excel export, or migration backup.

## SharePoint setup

Before deploying, connect PnP.PowerShell to the same SharePoint web as the tracker as a site owner, then run `scripts/setup-feedback.ps1`. Pass the actual SharePoint groups covering **all** tracker users through `-SubmitterGroups` and the app's configured prefix through `-ListPrefix` (default `Modernization`). For the scoped deployment, groups normally include `Metrology App User`, `Tracker Project Engineers`, and `Tracker Managers`; legacy deployments may use the associated Members group. Resolve the actual names in your site.

The script creates a dedicated list, breaks permission inheritance, disables search crawling and attachments, grants Full Control to the associated Owners group, removes the temporary creator grant, and assigns an exact Add Items permission mask to the submission groups. It does not grant View Items, Edit Items, Delete Items, or Manage Permissions. SharePoint's built-in Author and Created fields capture authenticated identity and time. Do not replace this role with Contribute, Read, or “read own items”: even authors must not be able to retrieve submissions. Site collection administrators retain their inherent administrative access.

The script intentionally refuses to overwrite an existing list or reuse a permission level. If interrupted, inspect the partially created list and role before removing empty setup artifacts and retrying. Never reset or delete a list containing feedback. Never enable inheritance or grant tracker managers a reading role on this list. Users with permission to administer the site can change ACLs; govern that privileged access separately.

## Required deployment acceptance checks

Using separate Viewer, Project Engineer, and non-owner Manager accounts, verify that both report types submit successfully, while direct SharePoint REST GET requests to the list's `/items` endpoint and individual items return access denied. Verify authors cannot read their own items or edit/delete them, and that the owner can see all submissions in the inbox. Test permissions in the target tenant: this repository's automated tests mock SharePoint and cannot establish live ACL behavior. If the tenant rejects add-only submission, keep the list private and use a server-side submission endpoint; do not grant read access to work around it.

Confirm the icon's placement, modal keyboard navigation, light/dark theme, narrow screen layout, and retained draft after a failed request. Network timeouts may occur after a write completes; check the owner inbox before retrying to avoid duplicates.

Local preview stores feedback only in module memory and labels it as disposable. It offers no multi-user security and must not collect real private feedback. Reloading clears those preview submissions.
