import { TrackerLists, displayListTitle } from './trackerLists';
const unwrap = body => body?.d || body;
const quote = value => encodeURIComponent(value.replace(/'/g, "''"));
const id = value => Number.isInteger(Number(value)) && Number(value) > 0;
const PENDING = 'NPSL Tracker private feedback setup pending v1';
const READY = 'Private bug reports and feature requests. Managed by NPSL Tracker.';
// AddListItems, ViewFormPages, Open, ViewPages; UseRemoteAPIs. NO ViewListItems.
export const SUBMIT_PERMISSIONS = { Low: '200706', High: '32' };
const sameMask = bits => String(bits?.Low) === SUBMIT_PERMISSIONS.Low && String(bits?.High) === SUBMIT_PERMISSIONS.High;

export function feedbackSetup(store, legacyTitle, isOwner) {
  const title = displayListTitle(store.prefix, 'feedback');
  const catalog = new TrackerLists(store.prefix, path => store.get(path), (path, options) => store.post(path, options));
  let path;
  const merge = body => store.post(path, { body, headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' } });
  async function requireOwner() {
    if (!await isOwner()) throw new Error('Only site owners can set up private feedback.');
  }
  async function pages(next) {
    const result = [];
    while (next) {
      const body = await store.get(next);
      const rows = body.value || body.d?.results;
      if (!Array.isArray(rows)) throw new Error('SharePoint did not return the requested setup records.');
      result.push(...rows);
      const link = body['@odata.nextLink'] || body['odata.nextLink'] || body.d?.__next;
      if (!link) break;
      const url = new URL(link, `${store.webUrl}/`);
      if (!url.href.startsWith(`${store.webUrl}/_api/`)) throw new Error('Unexpected setup pagination address.');
      next = url.href.slice(store.webUrl.length);
    }
    return result;
  }
  async function metadata() {
    try { path = await catalog.path('feedback'); return unwrap(await store.get(`${path}?$select=Id,Title,Description,ItemCount,HasUniqueRoleAssignments,BaseTemplate,Hidden`)); }
    catch (error) { if (error.status === 404) return null; throw error; }
  }
  return {
    async status() {
      await requireOwner();
      if (!store.webUrl) return { ready: true, title };
      const list = await metadata();
      if (!list) return { ready: false, title };
      if (list.Description === PENDING) return { ready: false, title, partial: true };
      if (!list.HasUniqueRoleAssignments) throw new Error('The feedback list inherits site permissions. A site administrator must make it private before it is used.');
      const fields = await pages(`${path}/fields?$select=InternalName&$top=500`);
      if (!['FeedbackKind', 'FeedbackDetails'].every(name => fields.some(field => field.InternalName === name))) throw new Error('The feedback list is missing required fields. Ask a site administrator to review its setup.');
      const owners = unwrap(await store.get('/_api/web/associatedownergroup?$select=Id,Title'));
      const assignments = await pages(`${path}/roleassignments?$select=PrincipalId,RoleDefinitionBindings/BasePermissions&$expand=RoleDefinitionBindings`);
      if (!id(owners?.Id) || assignments.some(assignment => {
        const roles = assignment.RoleDefinitionBindings?.results || assignment.RoleDefinitionBindings;
        if (!Array.isArray(roles)) return true;
        return roles.some(role => role.BasePermissions?.Low == null || (assignment.PrincipalId !== owners.Id && (BigInt(role.BasePermissions.Low) & 1n) !== 0n));
      })) throw new Error('The feedback list grants reading access outside the site Owners group, or its permissions could not be verified. Ask a site administrator to review it.');
      // Rename the existing list by GUID; never replace records or ACLs.
      if (list.Title !== title || list.Hidden) {
        const permissions = unwrap(await store.get('/_api/web/EffectiveBasePermissions'));
        if ((BigInt((permissions.EffectiveBasePermissions || permissions).Low || 0) & 2048n) !== 0n) {
          await merge({ Title: title, Hidden: false });
          const verified = await metadata();
          if (verified.Id !== list.Id || verified.Title !== title || verified.Hidden !== false) throw new Error('Could not verify the feedback list name. Reopen feedback to retry.');
        }
      }
      return { ready: true, title };
    },
    async options() {
      await requireOwner();
      const owners = unwrap(await store.get('/_api/web/associatedownergroup?$select=Id,Title'));
      if (!id(owners?.Id)) throw new Error('Configure the site’s associated Owners group before setting up feedback.');
      const groups = (await pages('/_api/web/sitegroups?$select=Id,Title&$top=500')).filter(group => id(group.Id) && group.Id !== owners.Id);
      const defaults = ['Metrology App User', 'Tracker Project Engineers', 'Tracker Managers'];
      if (!store.scopedAccess) {
        const members = unwrap(await store.get('/_api/web/associatedmembergroup?$select=Id,Title'));
        if (members?.Title) defaults.push(members.Title);
      }
      return { owners, groups, selected: groups.filter(group => defaults.includes(group.Title)).map(group => group.Id) };
    },
    async provision(selectedIds, progress = () => {}) {
      await requireOwner();
      // Creating schema and unique ACLs requires both site permissions.
      const response = unwrap(await store.get('/_api/web/EffectiveBasePermissions'));
      const low = BigInt((response.EffectiveBasePermissions || response).Low || 0);
      if ((low & 33556480n) !== 33556480n) throw new Error('Setup requires Manage Lists and Manage Permissions on this SharePoint site.');
      const { owners, groups } = await this.options();
      const selected = [...new Set(selectedIds.map(Number))];
      if (!selected.length || selected.some(value => !groups.some(group => group.Id === value))) throw new Error('Select the existing SharePoint groups that contain all tracker users.');
      const user = await store.currentUser();
      if (!id(user.id)) throw new Error('SharePoint could not verify the signed-in owner.');
      const roles = await pages('/_api/web/roledefinitions?$select=Id,Name,RoleTypeKind,BasePermissions&$top=500');
      const full = roles.find(role => role.RoleTypeKind === 5);
      if (!id(full?.Id)) throw new Error('The Full Control permission level is unavailable.');
      const roleName = `${legacyTitle} Submit Only`;
      let submitRole = roles.find(role => role.Name === roleName);
      if (submitRole && !sameMask(submitRole.BasePermissions)) throw new Error('The existing submission permission level is not add-only. No list changes were made; have a site administrator review it.');
      let list = await metadata();
      // Never reset an existing populated list or an unrelated list's permissions.
      if (list && (list.Description !== PENDING || list.ItemCount !== 0 || list.BaseTemplate !== 100)) throw new Error('This feedback list already exists. Setup will not overwrite its records or permissions.');
      progress('Creating private feedback storage…');
      if (!list) {
        await store.post('/_api/web/lists', { body: { Title: title, Description: PENDING, BaseTemplate: 100, Hidden: false, EnableAttachments: false, NoCrawl: true } });
        list = await metadata();
        if (!list || list.Description !== PENDING || list.ItemCount !== 0) throw new Error('Could not verify the new feedback list. Reopen setup to check its state.');
      }
      // No submission grants are made until schema and private ACLs are verified.
      await store.post(`${path}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`);
      await merge({ NoCrawl: true, EnableAttachments: false, Hidden: false });
      let assignments = await pages(`${path}/roleassignments?$select=PrincipalId,RoleDefinitionBindings/Id&$expand=RoleDefinitionBindings`);
      if (!assignments.some(assignment => assignment.PrincipalId === owners.Id && (assignment.RoleDefinitionBindings?.results || assignment.RoleDefinitionBindings || []).some(role => role.Id === full.Id))) await store.post(`${path}/roleassignments/addroleassignment(principalid=${owners.Id},roledefid=${full.Id})`);
      for (const assignment of assignments) {
        if (![owners.Id, user.id].includes(assignment.PrincipalId)) await store.post(`${path}/roleassignments/getbyprincipalid(${assignment.PrincipalId})`, { headers: { 'X-HTTP-Method': 'DELETE' } });
      }
      const fields = await pages(`${path}/fields?$select=InternalName&$top=500`);
      for (const [name, xml] of [
        ['FeedbackKind', '<Field Type="Choice" DisplayName="Feedback kind" Name="FeedbackKind" StaticName="FeedbackKind" Required="TRUE"><CHOICES><CHOICE>Bug</CHOICE><CHOICE>Feature</CHOICE></CHOICES></Field>'],
        ['FeedbackDetails', '<Field Type="Note" DisplayName="Feedback details" Name="FeedbackDetails" StaticName="FeedbackDetails" Required="TRUE" RichText="FALSE" />'],
      ]) if (!fields.some(field => field.InternalName === name)) await store.post(`${path}/fields/createfieldasxml`, { verbose: true, body: { parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: xml, Options: 24 } } });
      if (!submitRole) {
        await store.post('/_api/web/roledefinitions', { body: { Name: roleName, Description: 'Submit feedback without viewing, editing, or deleting submissions.', BasePermissions: SUBMIT_PERMISSIONS } });
        submitRole = unwrap(await store.get(`/_api/web/roledefinitions/getbyname('${quote(roleName)}')?$select=Id,BasePermissions`));
      }
      if (!id(submitRole?.Id) || !sameMask(submitRole.BasePermissions)) throw new Error('SharePoint did not confirm an add-only permission level. No submission access was granted.');
      const secured = await metadata();
      const verifiedFields = await pages(`${path}/fields?$select=InternalName&$top=500`);
      if (!secured?.HasUniqueRoleAssignments || !['FeedbackKind', 'FeedbackDetails'].every(name => verifiedFields.some(field => field.InternalName === name))) throw new Error('The private list or its fields could not be verified. Submission access was not granted.');
      progress('Applying submission-only access…');
      for (const principal of selected) await store.post(`${path}/roleassignments/addroleassignment(principalid=${principal},roledefid=${submitRole.Id})`);
      assignments = await pages(`${path}/roleassignments?$select=PrincipalId,RoleDefinitionBindings/Id&$expand=RoleDefinitionBindings`);
      const bindings = assignment => assignment.RoleDefinitionBindings?.results || assignment.RoleDefinitionBindings || [];
      const expected = new Map([[owners.Id, full.Id], [user.id, full.Id], ...selected.map(principal => [principal, submitRole.Id])]);
      if (assignments.some(assignment => !expected.has(assignment.PrincipalId) || bindings(assignment).some(role => role.Id !== expected.get(assignment.PrincipalId))) || ![owners.Id, ...selected].every(principal => assignments.some(assignment => assignment.PrincipalId === principal && bindings(assignment).some(role => role.Id === expected.get(principal))))) throw new Error('Could not verify the feedback permission assignments. Ask a site administrator to review this list before accepting submissions.');
      await merge({ Description: READY, Title: title, Hidden: false });
      // Remove the automatic creator grant. Only Owners retain reading access.
      if (assignments.some(assignment => assignment.PrincipalId === user.id)) await store.post(`${path}/roleassignments/getbyprincipalid(${user.id})`, { headers: { 'X-HTTP-Method': 'DELETE' } });
      return { ready: true, title };
    },
  };
}
