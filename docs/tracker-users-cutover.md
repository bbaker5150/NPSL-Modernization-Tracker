# Tracker Users access cutover

Scoped Tracker invitations now enroll users in **Tracker Users**, replacing the suite-wide **Metrology App User** group. All Tracker roles use this basic-access group; **Tracker Managers** and **Tracker Project Engineers** retain their existing elevated roles. The Home page and other applications are not changed by this release. Group names live in `src/lib/trackerGroups.js`.

## Existing SharePoint preparation

- Tracker Users exists, is owned by Tracker Managers if managers should enroll users, and allows membership edits only by its owner. Include intended Managers and Engineers as well as Viewers.
- It has Read on the Tracker ASPX, its specific assets folder/files, and Projects, Tasks, Updates, Risks, Users, Acronyms and Reference Documents. Do not grant it site-wide access merely to make startup pass.
- It has the existing **ModernizationFeedback Submit Only** permission on Feedback. Reading reports remains reserved for site owners.
- Tracker Managers retains Contribute on Users, Acronyms and Reference Documents and the custom project-access-manager level on Projects, Tasks, Updates and Risks. Its Users-list Add Items permission is required to save an invitee before enrollment/email.

## Deploy and verify

1. Deploy this HTML and refresh active Tracker sessions so invitations run the new code. This build also includes the pending Viewer startup compatibility fix: absent optional edit-attribution columns do not block existing data reads.
2. Open as a site owner. In Users and managers → Tracker permissions, run **Apply tracker permissions** and wait for completion. It adds/verifies Tracker Users Read on projects, task/update/risk folders and individually secured legacy records. Existing owner, manager, engineer and unrelated group grants are preserved. List-level/page/asset/Feedback grants remain the owner's separately configured prerequisites.
3. Test a non-owner account that belongs to Tracker Users but not Metrology App User or another broad-access group. Verify portfolio/history/attachments and private feedback submission. Verify no project editing or owner inbox. Test other apps through their direct page and data URLs against actual tenant permissions; METBASE uses a different SharePoint site.
4. As a non-owner Manager, invite a new Viewer. Confirm the Users record saves, Tracker Users membership is verified, the invitation is accepted, and Metrology App User membership is unchanged. Repeat Engineer/Manager invitation tests with an account authorized to manage the relevant elevated group. Manager promotion still requires ownership/administrative authority over Tracker Managers.
5. Confirm engineers can edit their assigned projects, original due/assigned dates remain locked, and deferred dates remain editable. Demoting to Viewer removes elevated access while retaining Tracker Users membership. A Viewer label alone does not establish membership.
6. Complete any intended cleanup of old suite-group grants or memberships only after testing. Ordinary repair preserves unrelated grants; it does not revoke Metrology App User membership or its existing resource grants. Removing a Tracker directory entry does not revoke basic access. Use the corresponding SharePoint group/resource controls for that purpose.

The invitation flow uses the signed-in person's SharePoint permissions. It validates Tracker Users membership and target-page/list/asset access before requesting a single page-targeted email. Missing Tracker Users, denied membership access or an unconfirmed add stops the flow. No automatic fallback enrolls the person in Metrology App User. Earlier directory or elevated-group changes can remain after a later failure; retry after addressing the reported step.

Automated API fixtures and packaged-browser tests verify this behavior without modifying a live tenant. They do not establish that a person lacks every possible direct or inherited grant to another app.
