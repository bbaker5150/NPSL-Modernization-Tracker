import { feedbackAttachmentTitle } from './feedbackAttachments';
import { legacyListTitle } from './trackerLists';
import { TRACKER_GROUPS } from './trackerGroups';
export { feedbackAttachmentTitle } from './feedbackAttachments';

const PENDING = 'NPSL Tracker private feedback attachments setup pending v1';
const READY = 'Private feedback files. Managed by NPSL Tracker.';
const SUBMIT_MASK = '32:200706';
const REVIEW_MASK = '32:200711';
const ATTACHMENT_REVIEW_PERMISSIONS = { Low: '200743', High: '32' };
const FIELD = 'FeedbackAttachments';
const FIELD_XML = '<Field Type="Note" DisplayName="Feedback attachments" Name="FeedbackAttachments" StaticName="FeedbackAttachments" RichText="FALSE" AppendOnly="FALSE" Required="FALSE" />';
const unwrap = body => body?.d || body;
const validId = value => Number.isSafeInteger(Number(value)) && Number(value) > 0;
const mask = bits => {
  if (!bits || !/^\d+$/.test(String(bits.Low)) || !/^\d+$/.test(String(bits.High)) || BigInt(bits.Low) > 4294967295n || BigInt(bits.High) > 4294967295n) throw new Error('SharePoint did not return verifiable feedback permissions.');
  return `${BigInt(bits.High)}:${BigInt(bits.Low)}`;
};
const bindings = assignment => assignment.RoleDefinitionBindings?.results || assignment.RoleDefinitionBindings;

// Attachments are optional storage. Neither tracker startup nor text-only
// submission needs this library, its schema, or permission to inspect its ACL.
export function feedbackAttachmentsSetup(store, isOwner, feedbackPath, verifyFeedback) {
  const title = feedbackAttachmentTitle(store.prefix);
  const setupTitle = `${title} - Setup`;
  const reviewRoleName = `${legacyListTitle(store.prefix, 'feedback')} Attachment Review`;
  const namedPath = name => `/_api/web/lists/getbytitle('${encodeURIComponent(name.replace(/'/g, "''"))}')`;
  const roleQuery = '/roleassignments?$select=PrincipalId,RoleDefinitionBindings/Id,RoleDefinitionBindings/BasePermissions&$expand=RoleDefinitionBindings';
  let path;
  async function requireOwner() {
    if (!await isOwner()) throw new Error('Only site owners can set up private feedback attachments.');
  }
  async function pages(next) {
    const result = [], seen = new Set();
    while (next) {
      if (seen.has(next)) throw new Error('SharePoint repeated a feedback permissions page.');
      seen.add(next);
      const body = await store.get(next);
      const rows = body?.value || body?.d?.results;
      if (!Array.isArray(rows)) throw new Error('SharePoint did not return the requested feedback setup records.');
      result.push(...rows);
      const link = body['@odata.nextLink'] || body['odata.nextLink'] || body.d?.__next;
      if (!link) break;
      const url = new URL(link, `${store.webUrl}/`);
      if (!url.href.startsWith(`${store.webUrl}/_api/`)) throw new Error('Unexpected feedback setup pagination address.');
      next = url.href.slice(store.webUrl.length);
    }
    return result;
  }
  function acl(rows) {
    const result = new Map();
    for (const row of rows) {
      const roles = bindings(row), principal = Number(row.PrincipalId);
      if (!validId(principal) || !Array.isArray(roles) || !roles.length) throw new Error('SharePoint did not return verifiable feedback permission assignments.');
      const entry = result.get(principal) || new Map();
      for (const role of roles) {
        if (!validId(role.Id)) throw new Error('SharePoint did not return a valid feedback permission level.');
        const value = mask(role.BasePermissions), roleId = Number(role.Id);
        if (entry.has(roleId) && entry.get(roleId) !== value) throw new Error('SharePoint returned conflicting feedback permission levels.');
        entry.set(roleId, value);
      }
      result.set(principal, entry);
    }
    return result;
  }
  function equalAcl(actual, expected) {
    return actual.size === expected.size && [...expected].every(([principal, roles]) => {
      const found = actual.get(principal);
      return found?.size === roles.size && [...roles].every(([role, bits]) => found.get(role) === bits);
    });
  }
  async function source({ ensureReview = false, requireManagerReady = false } = {}) {
    const state = await verifyFeedback();
    if (!state?.ready) throw new Error('Finish setting up private feedback before enabling attachments.');
    if (requireManagerReady && state.managerReviewNeeded) throw new Error('Enable manager history access before setting up private attachments.');
    const sourcePath = await feedbackPath();
    const owners = unwrap(await store.get('/_api/web/associatedownergroup?$select=Id,Title'));
    const roleCatalog = '/_api/web/roledefinitions?$select=Id,Name,RoleTypeKind,BasePermissions&$top=500';
    const roles = await pages(roleCatalog);
    const full = roles.find(role => Number(role.RoleTypeKind) === 5);
    if (!validId(owners?.Id) || !validId(full?.Id)) throw new Error('The site Owners group and Full Control permission level must be configured before enabling feedback attachments.');
    const expected = acl(await pages(`${sourcePath}${roleQuery}`));
    if (expected.get(Number(owners.Id))?.get(Number(full.Id)) !== mask(full.BasePermissions)) throw new Error('The feedback list must grant its associated Owners group Full Control before enabling attachments.');
    const groups = await pages(`/_api/web/sitegroups?$select=Id,Title&$filter=Title eq '${TRACKER_GROUPS.Manager}'`);
    const managers = groups.find(group => group.Title === TRACKER_GROUPS.Manager && validId(group.Id));
    let managerReview = false;
    for (const [principal, grants] of expected) {
      if (principal === Number(owners.Id)) continue;
      for (const [roleId, bits] of grants) {
        if (principal === Number(managers?.Id) && bits === REVIEW_MASK) { managerReview = true; continue; }
        const catalogRole = roles.find(role => Number(role.Id) === roleId);
        const limitedAccess = Number(catalogRole?.RoleTypeKind) === 1 && (BigInt(bits.split(':')[1]) & 45n) === 0n;
        if (bits !== SUBMIT_MASK && !limitedAccess) throw new Error('Feedback submission permissions are not add-only. Only the associated Owners group and verified Tracker Managers may read or edit feedback files.');
      }
    }
    if (!managerReview) throw new Error('Enable manager history access before setting up private attachments.');
    let reviewer = roles.find(role => role.Name === reviewRoleName);
    if (!reviewer && ensureReview) {
      await store.post('/_api/web/roledefinitions', { body: { Name: reviewRoleName, Description: 'Read and download private Tracker feedback files without deleting them.', BasePermissions: ATTACHMENT_REVIEW_PERMISSIONS } });
      reviewer = (await pages(roleCatalog)).find(role => role.Name === reviewRoleName);
    }
    if (!reviewer && !ensureReview) return { sourcePath, expected: null, reviewRoleNeeded: true, ownerId: Number(owners.Id), fullId: Number(full.Id) };
    if (!validId(reviewer?.Id) || mask(reviewer.BasePermissions) !== mask(ATTACHMENT_REVIEW_PERMISSIONS)) throw new Error('The feedback attachment review permission level could not be verified. Existing permission levels were not changed.');
    const managerRoles = expected.get(Number(managers.Id));
    for (const [roleId, bits] of managerRoles) if (bits === REVIEW_MASK) managerRoles.delete(roleId);
    managerRoles.set(Number(reviewer.Id), mask(reviewer.BasePermissions));
    return { sourcePath, expected, ownerId: Number(owners.Id), fullId: Number(full.Id) };
  }
  async function metadata() {
    const matches = [];
    for (const name of [title, setupTitle]) {
      let list;
      try { list = unwrap(await store.get(`${namedPath(name)}?$select=Id,Title,Description,ItemCount,HasUniqueRoleAssignments,BaseTemplate,NoCrawl`)); }
      catch (error) { if (error.status === 404) continue; throw error; }
      if (!/^[\da-f]{8}(-[\da-f]{4}){3}-[\da-f]{12}$/i.test(list?.Id || '') || list.Title !== name) throw new Error('SharePoint did not return the expected feedback attachments library.');
      matches.push(list);
    }
    if (matches.length > 1) throw new Error('Both published and pending feedback attachment libraries exist. Ask a site administrator to resolve the conflict; no library was selected.');
    if (!matches.length) return null;
    const list = matches[0];
    path = `/_api/web/lists(guid'${list.Id}')`;
    return list;
  }
  async function field(sourcePath) {
    // Missing direct field lookups can return HTTP 400. A collection query
    // distinguishes an absent column from a failed or unauthorized request.
    const body = await store.get(`${sourcePath}/fields?$filter=InternalName eq '${FIELD}'&$select=InternalName,TypeAsString,RichText,AppendOnly`);
    const fields = body?.value || body?.d?.results;
    if (!Array.isArray(fields)) throw new Error('Could not verify the feedback attachments column.');
    const value = fields.find(entry => entry.InternalName === FIELD);
    if (!value) return null;
    if (value?.InternalName !== FIELD || value.TypeAsString !== 'Note' || value.RichText !== false || value.AppendOnly === true) throw new Error('FeedbackAttachments must be a plain multiple-lines-of-text field with append changes disabled. No existing field was changed.');
    return value;
  }
  async function ensureField(sourcePath) {
    if (!await field(sourcePath)) await store.post(`${sourcePath}/fields/createfieldasxml`, {
      verbose: true, body: { parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: FIELD_XML, Options: 24 } },
    });
    if (!await field(sourcePath)) throw new Error('SharePoint did not verify the feedback attachments field.');
  }
  function isPendingEmpty(list) {
    return list.Title === setupTitle && list.Description === PENDING && Number(list.BaseTemplate) === 101 && list.ItemCount === 0;
  }
  async function verifyLibrary(list, expected, pending = false) {
    if (list.Title !== (pending ? setupTitle : title) || Number(list.BaseTemplate) !== 101 || list.Description !== (pending ? PENDING : READY) || list.HasUniqueRoleAssignments !== true || list.NoCrawl !== true) throw new Error('The feedback attachments library is not verified private storage. Ask a site administrator to review it; existing permissions were not reset.');
    if (!equalAcl(acl(await pages(`${path}${roleQuery}`)), expected)) throw new Error('The feedback attachments permissions do not exactly match the approved Feedback permissions. Ask a site administrator to review them; existing permissions were not reset.');
  }
  const merge = body => store.post(path, { body, headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' } });
  const addRole = (principal, role) => store.post(`${path}/roleassignments/addroleassignment(principalid=${principal},roledefid=${role})`);
  const removeRole = (principal, role) => store.post(`${path}/roleassignments/removeroleassignment(principalid=${principal},roledefid=${role})`);

  return {
    async status() {
      await requireOwner();
      if (!store.webUrl) return { ready: true, title };
      const state = await verifyFeedback();
      if (!state?.ready) return { ready: false, title };
      const list = await metadata();
      if (!list) return { ready: false, title };
      if (isPendingEmpty(list)) return { ready: false, title, partial: true };
      const { sourcePath, expected, reviewRoleNeeded } = await source();
      if (reviewRoleNeeded) return { ready: false, title };
      await verifyLibrary(list, expected);
      return { ready: !!await field(sourcePath), title };
    },
    async provision(progress = () => {}) {
      await requireOwner();
      if (!store.webUrl) return { ready: true, title };
      const response = unwrap(await store.get('/_api/web/EffectiveBasePermissions'));
      if ((BigInt((response.EffectiveBasePermissions || response).Low || 0) & 33556480n) !== 33556480n) throw new Error('Setup requires Manage Lists and Manage Permissions on this SharePoint site.');
      let approved = await source({ requireManagerReady: true });
      const user = await store.currentUser();
      if (!validId(user?.id)) throw new Error('SharePoint could not verify the signed-in owner.');
      // Verify an existing field before creating or granting access to storage.
      await field(approved.sourcePath);
      let list = await metadata();
      if (list && list.Description !== READY && !isPendingEmpty(list)) throw new Error('The feedback attachments library already exists. Setup will not overwrite its files or permissions.');
      approved = await source({ ensureReview: true, requireManagerReady: true });
      if (list?.Description === READY) {
        await verifyLibrary(list, approved.expected);
        await ensureField(approved.sourcePath);
        return this.status();
      }
      if (list && !isPendingEmpty(list)) throw new Error('The feedback attachments library already exists. Setup will not overwrite its files or permissions.');
      progress('Creating private feedback attachment storage…');
      if (!list) {
        // Uploads resolve only the final title. Publish that title only after
        // privacy and schema verification, never during inherited site access.
        await store.post('/_api/web/lists', { body: { Title: setupTitle, Description: PENDING, BaseTemplate: 101, Hidden: true, NoCrawl: true } });
        list = await metadata();
      }
      if (!list || !isPendingEmpty(list)) throw new Error('Could not verify an empty feedback attachments library. Existing files and permissions were not changed.');
      // Only a new or interrupted, empty app-created library may be reset.
      await store.post(`${path}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`);
      await merge({ NoCrawl: true });
      let current = acl(await pages(`${path}${roleQuery}`));
      if (!current.get(approved.ownerId)?.has(approved.fullId)) await addRole(approved.ownerId, approved.fullId);
      for (const [principal, roles] of current) {
        for (const role of roles.keys()) {
          if (role === approved.fullId && [approved.ownerId, Number(user.id)].includes(principal)) continue;
          await removeRole(principal, role);
        }
      }
      await ensureField(approved.sourcePath);
      progress('Applying private feedback attachment permissions…');
      current = acl(await pages(`${path}${roleQuery}`));
      for (const [principal, roles] of approved.expected) {
        for (const role of roles.keys()) if (!current.get(principal)?.has(role)) await addRole(principal, role);
      }
      // Do not leave the automatic direct creator grant alongside group access.
      current = acl(await pages(`${path}${roleQuery}`));
      for (const role of current.get(Number(user.id))?.keys() || []) {
        if (!approved.expected.get(Number(user.id))?.has(role)) await removeRole(Number(user.id), role);
      }
      const fresh = await source();
      if (!equalAcl(fresh.expected, approved.expected)) throw new Error('Feedback permissions changed during setup. Reopen setup to verify the current permissions before continuing.');
      await verifyLibrary(await metadata(), fresh.expected, true);
      await merge({ Title: title, Description: READY, Hidden: false });
      const verified = await this.status();
      if (!verified.ready) throw new Error('SharePoint did not verify private feedback attachment storage. Reopen setup to check its state.');
      return verified;
    },
  };
}
