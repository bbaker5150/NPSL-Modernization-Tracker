import { isOwnedByUser } from './identity';
import { projectEngineers } from './projectEngineers';
import { TRACKER_GROUPS, TRACKER_USERS_GROUP } from './trackerGroups';
import { legacyListTitle } from './trackerLists';
export { TRACKER_GROUPS } from './trackerGroups';

const CHILDREN = ['tasks', 'updates', 'risks'];
const suffixes = { projects: 'Projects', tasks: 'Tasks', updates: 'Updates', risks: 'Risks' };
const unwrap = body => body?.d || body;
const quote = value => encodeURIComponent(String(value).replace(/'/g, "''"));
const positiveId = value => Number.isInteger(Number(value)) && Number(value) > 0;
const engineerPrincipalSet = value => {
  const ids = value == null ? [] : value instanceof Set ? [...value] : Array.isArray(value) ? value : [value];
  if (ids.some(id => !Number.isSafeInteger(Number(id)) || Number(id) <= 0)) {
    throw new Error('Could not verify every project engineer principal. No permission changes were made.');
  }
  return new Set(ids.map(Number));
};

// No privileged service identity: every operation runs as the signed-in person.
// Group ownership and the four list ACLs must be configured by the site owner.
export class TrackerPermissions {
  constructor(store) { this.store = store; }
  root(key) { return `/_api/web/lists/getbytitle('${quote(`${String(this.store.prefix).replace(/[^A-Za-z0-9]/g, '') || 'Modernization'}${suffixes[key]}`)}')`; }
  get(path, headers) { return this.store.get(path, headers); }
  post(path, options) { return this.store.post(path, options); }
  async pages(path) {
    const rows = [];
    while (path) {
      const body = await this.get(path);
      const page = body?.value || body?.d?.results || body?.results;
      if (!Array.isArray(page)) throw new Error('SharePoint did not return the requested permission records. Retry after checking access.');
      rows.push(...page);
      const next = body['@odata.nextLink'] || body['odata.nextLink'] || body.d?.__next;
      if (!next) break;
      const url = new URL(next, `${this.store.webUrl}/`);
      if (!url.href.startsWith(`${this.store.webUrl}/`)) throw new Error('Unexpected SharePoint pagination address.');
      path = url.href.slice(this.store.webUrl.length);
    }
    return rows;
  }
  async currentRole(user) {
    const groups = await this.pages('/_api/web/currentuser/groups?$select=Title&$top=500');
    const permissions = unwrap(await this.get('/_api/web/effectivebasepermissions'));
    const bits = permissions?.EffectiveBasePermissions || permissions;
    const siteOwner = user.isSiteAdmin || (Number(bits?.Low) & 33554432) !== 0;
    return { siteOwner, role: siteOwner || groups.some(g => g.Title === TRACKER_GROUPS.Manager) ? 'Manager'
      : groups.some(g => g.Title === TRACKER_GROUPS['Project Engineer']) ? 'Project Engineer' : 'Viewer' };
  }
  async group(role) {
    const title = TRACKER_GROUPS[role];
    if (!title) throw new Error('Select a valid tracker role.');
    return this.groupByName(title);
  }
  async groupByName(title) {
    const group = unwrap(await this.get(`/_api/web/sitegroups/getbyname('${quote(title)}')?$select=Id,Title`));
    if (!positiveId(group?.Id) || group.Title !== title) throw new Error(`Verify the SharePoint group ${title} before assigning this role.`);
    return group;
  }
  async membership(group, login) {
    return (await this.pages(`/_api/web/sitegroups(${group.Id})/users?$select=Id,LoginName&$filter=LoginName eq '${quote(login)}'`))
      .find(row => row.LoginName?.toLowerCase() === login.toLowerCase());
  }
  async ensureAppMembership(person) {
    if (!person.loginName?.trim()) throw new Error('Select a resolved SharePoint identity.');
    const title = TRACKER_USERS_GROUP;
    try {
      const group = unwrap(await this.get(`/_api/web/sitegroups/getbyname('${quote(title)}')?$select=Id,Title`));
      if (!positiveId(group?.Id) || group.Title !== title) throw new Error(`The ${title} group is unavailable. Ask a site owner to create or verify it.`);
      // A failed read must not be treated as a missing member.
      if (!(await this.membership(group, person.loginName))) {
        await this.post(`/_api/web/sitegroups(${group.Id})/users`, { body: { LoginName: person.loginName } });
      }
      if (!(await this.membership(group, person.loginName))) throw new Error('SharePoint did not confirm group membership.');
      return group;
    } catch (error) {
      throw new Error(`Could not verify membership in ${title}. Membership may already be saved. Ask a site owner to verify the group and allow Tracker Managers to view and manage its membership, then retry. No invitation email was requested. ${error.message}`);
    }
  }
  async ensureElevatedMembership(person, group, member) {
    if (!member) {
      try { await this.post(`/_api/web/sitegroups(${group.Id})/users`, { body: { LoginName: person.loginName } }); }
      catch (error) { throw new Error(`Could not add membership in ${group.Title}. Ask its group owner or a site owner. ${error.message}`); }
    }
    if (!(await this.membership(group, person.loginName))) throw new Error(`SharePoint did not verify ${group.Title} membership. Retry the role change.`);
  }

  async ensureAssignmentRole(person, savedRole, siteAdmin) {
    const managers = await this.group('Manager');
    let isManager = siteAdmin || Boolean(await this.membership(managers, person.loginName));
    if (!isManager) {
      // Site owners can manage Tracker without belonging to Tracker Managers.
      try {
        const response = unwrap(await this.get(`/_api/web/getusereffectivepermissions(@u)?@u='${quote(person.loginName)}'`));
        const low = (response?.GetUserEffectivePermissions || response)?.Low;
        if (!['number', 'string'].includes(typeof low) || !/^\d+$/.test(String(low)) || Number(low) > 0xffffffff) throw new Error('SharePoint did not verify site permissions for the selected project engineer.');
        isManager = (Number(low) & 33554432) !== 0;
      } catch (error) {
        // Delegated group owners may only inspect list-level permissions. Keep
        // a saved Manager role in that case; assignment never adds to Managers.
        if (error.status !== 403) throw error;
        isManager = savedRole === 'Manager';
      }
    }
    if (!isManager && savedRole === 'Manager') throw new Error('The saved Manager role differs from this account’s SharePoint permissions. Verify its Manager role in Users and managers before assigning the project.');
    await this.ensureAppMembership(person);
    if (isManager) return 'Manager';
    const engineers = await this.group('Project Engineer');
    await this.ensureElevatedMembership(person, engineers, await this.membership(engineers, person.loginName));
    return 'Project Engineer';
  }
  async syncGroups(person, role, { beforeChange, ensureViewerMembership = false } = {}) {
    if (!person.loginName) throw new Error('Select a resolved SharePoint identity.');
    // Read every membership before changing any. Failure is never interpreted as absence.
    const state = [];
    for (const name of Object.keys(TRACKER_GROUPS)) {
      const group = await this.group(name);
      state.push({ name, group, member: await this.membership(group, person.loginName) });
    }
    const target = state.find(entry => entry.name === role);
    if (role && role !== 'Viewer' && !target) throw new Error('Select a valid tracker role.');
    // Let role changes revoke engineer grants using the live memberships we
    // just read, before removing groups or granting the replacement role.
    if (beforeChange) await beforeChange(state);
    // Existing Viewers can be repaired by saving their role again. Any
    // transition from a live elevated group must preserve basic Tracker access.
    if (role === 'Viewer' && (ensureViewerMembership || state.some(entry => entry.member))) {
      await this.ensureAppMembership(person);
    }
    // Remove stale privileged membership before granting the replacement role.
    for (const entry of state.filter(entry => entry.member && entry !== target)) {
      try {
        await this.post(`/_api/web/sitegroups(${entry.group.Id})/users/removebyid(${entry.member.Id})`);
        if (await this.membership(entry.group, person.loginName)) throw new Error('Membership is still present.');
      } catch (error) { throw new Error(`Could not remove membership in ${entry.group.Title}. Its group owner must complete this role change. ${error.message}`); }
    }
    if (target) await this.ensureElevatedMembership(person, target.group, target.member);
    return target?.group || null;
  }
  async revokeEngineerAccess(person, memberships = []) {
    if (!person.loginName) throw new Error('Select a resolved SharePoint identity.');
    // Prefer the principal already verified in the live tracker memberships.
    // An interrupted earlier demotion may have removed every membership, so
    // resolve that case with a read, never ensureuser or a new permission grant.
    const members = memberships.filter(entry => entry.member).map(entry => entry.member);
    if (!members.length) members.push(...await this.pages(`/_api/web/siteusers?$select=Id,LoginName&$filter=LoginName eq '${quote(person.loginName)}'`));
    if (!members.length || members.some(member => !positiveId(member.Id) || member.LoginName?.toLowerCase() !== person.loginName.toLowerCase()) || new Set(members.map(member => Number(member.Id))).size !== 1) {
      throw new Error('Could not uniquely verify this user’s SharePoint identity before removing engineer access.');
    }
    const principalId = Number(members[0].Id), roleId = await this.contributionRole();
    const hasGrant = rows => rows.some(row => Number(row.PrincipalId) === principalId && Number(row.Member.PrincipalType) === 1 &&
      (row.RoleDefinitionBindings.results || row.RoleDefinitionBindings).some(binding => Number(binding.Id) === Number(roleId)));
    // Include folders, archived records, and former assignments, but never
    // initialize scopes or repair other people's access during a role change.
    const inventories = await Promise.all(['projects', ...CHILDREN].map(async key => {
      const rows = await this.pages(`${this.root(key)}/items?$select=Id,HasUniqueRoleAssignments&$top=500`);
      if (rows.some(row => !positiveId(row.Id) || typeof row.HasUniqueRoleAssignments !== 'boolean')) throw new Error(`Cannot verify unique permission scopes in ${key}. No engineer grants were removed.`);
      return rows.filter(row => row.HasUniqueRoleAssignments).map(row => `${this.root(key)}/items(${Number(row.Id)})`);
    }));
    const scopes = [...new Set(inventories.flat())], removals = [];
    // Preflight all ACL reads before the first mutation. Limit parallel reads
    // so a large legacy task list does not flood SharePoint or the host bridge.
    for (let offset = 0; offset < scopes.length; offset += 4) {
      const checks = await Promise.allSettled(scopes.slice(offset, offset + 4).map(async scope => ({ scope, rows: await this.assignments(scope) })));
      const failed = checks.find(result => result.status === 'rejected');
      if (failed) throw failed.reason;
      removals.push(...checks.filter(result => hasGrant(result.value.rows)).map(result => result.value.scope));
    }
    let removed = 0;
    for (const scope of removals) {
      const info = unwrap(await this.get(`${scope}?$select=HasUniqueRoleAssignments`));
      if (info.HasUniqueRoleAssignments !== true) throw new Error('Project access changed during this role update. Retry to check the current permissions.');
      if (!hasGrant(await this.assignments(scope))) continue;
      await this.post(`${scope}/roleassignments/removeroleassignment(principalid=${principalId},roledefid=${roleId})`);
      if (hasGrant(await this.assignments(scope))) throw new Error('SharePoint did not verify removal of the former engineer grant. Retry this role change.');
      removed++;
    }
    return removed;
  }
  async contributionRole() {
    const roles = await this.pages('/_api/web/roledefinitions?$select=Id,RoleTypeKind&$filter=RoleTypeKind eq 3');
    const role = roles.find(row => row.RoleTypeKind === 3);
    if (!positiveId(role?.Id)) throw new Error('SharePoint Contribute permission level is unavailable.');
    return role.Id;
  }
  async prepareGlossaryAccess() {
    const { siteOwner } = await this.currentRole(await this.store.currentUser());
    if (!siteOwner) throw new Error('Only a SharePoint site owner can set up glossary access.');
    const scope = await this.store.listApi('acronyms');
    const [basic, engineers, managers] = await Promise.all([
      this.groupByName(TRACKER_USERS_GROUP), this.group('Project Engineer'), this.group('Manager'),
    ]);
    const roles = await this.pages('/_api/web/roledefinitions?$select=Id,Name,RoleTypeKind,BasePermissions');
    const contribute = roles.find(role => role.RoleTypeKind === 3);
    if (!positiveId(contribute?.Id) || (Number(contribute.BasePermissions?.Low) & 15) !== 15) throw new Error('Check the SharePoint Contribute permission level before setting up glossary access.');
    const roleName = `${legacyListTitle(this.store.prefix, 'acronyms')} Add Only`;
    // View/Add Items, View Forms, Open, View Pages, and Use Remote APIs.
    const addMask = { Low: '200707', High: '32' };
    const info = unwrap(await this.get(`${scope}?$select=HasUniqueRoleAssignments,ReadSecurity,WriteSecurity`));
    if (typeof info.HasUniqueRoleAssignments !== 'boolean') throw new Error('SharePoint did not verify permission inheritance on the Acronyms list.');
    if (Number(info.ReadSecurity) !== 1 || Number(info.WriteSecurity) !== 1) {
      throw new Error('Check item-level permissions in Acronyms list settings → Advanced settings. Select Read all items and Create and edit all items. The list grants then control who can add and who can edit. Retry glossary setup.');
    }
    const current = await this.assignments(scope);
    for (const row of current.filter(row => [Number(basic.Id), Number(engineers.Id)].includes(Number(row.PrincipalId)))) {
      for (const binding of row.RoleDefinitionBindings.results || row.RoleDefinitionBindings) {
        const assigned = roles.find(role => Number(role.Id) === Number(binding.Id));
        const bits = assigned?.BasePermissions?.Low;
        // Editing an item can archive it, so both Edit Items and Delete Items
        // must be manager-only. Leave existing broad grants for owner review.
        if (bits == null || !/^\d+$/.test(String(bits)) || (Number(bits) & (4 | 8 | 2048 | 33554432))) {
          throw new Error('Review the Tracker Users and Tracker Project Engineers grants on the Acronyms list. Give these groups Read and glossary add access only. Existing grants were preserved.');
        }
      }
    }
    let addRole = roles.find(role => role.Name === roleName);
    if (!addRole) {
      await this.post('/_api/web/roledefinitions', { body: { Name: roleName, Description: 'Read and add glossary entries.', BasePermissions: addMask } });
      addRole = unwrap(await this.get(`/_api/web/roledefinitions/getbyname('${quote(roleName)}')?$select=Id,Name,BasePermissions`));
    }
    if (!positiveId(addRole?.Id) || addRole.Name !== roleName || String(addRole.BasePermissions?.Low) !== addMask.Low || String(addRole.BasePermissions?.High) !== addMask.High) {
      throw new Error(`Review the ${roleName} permission level. SharePoint did not verify its read and add permissions.`);
    }
    const grants = [
      { principalId: Number(basic.Id), roleId: Number(addRole.Id) },
      { principalId: Number(engineers.Id), roleId: Number(addRole.Id) },
      { principalId: Number(managers.Id), roleId: Number(contribute.Id) },
    ];
    if (!info.HasUniqueRoleAssignments) await this.post(`${scope}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`);
    for (const grant of grants) {
      if (!this.matchesBaseline(current, [grant])) await this.post(`${scope}/roleassignments/addroleassignment(principalid=${grant.principalId},roledefid=${grant.roleId})`);
    }
    if (!this.matchesBaseline(await this.assignments(scope), grants)) throw new Error('SharePoint did not verify the glossary grants. Retry glossary setup.');
  }
  async baselineGrants() {
    const roles = await this.pages('/_api/web/roledefinitions?$select=Id,Name,RoleTypeKind,BasePermissions');
    const read = roles.find(row => row.RoleTypeKind === 2);
    const manager = roles.find(row => row.Name === 'Tracker Project Access Manager');
    const required = 33554447n; // Manage Permissions plus View/Add/Edit/Delete Items.
    let managerBits = 0n;
    try { managerBits = BigInt(manager?.BasePermissions?.Low || 0); } catch { /* reject below */ }
    if (!positiveId(read?.Id) || !positiveId(manager?.Id) || (managerBits & required) !== required) {
      throw new Error('A site owner must configure Read and Tracker Project Access Manager (Contribute plus Manage Permissions) before repairing project access.');
    }
    const grants = [];
    for (const [title, roleId] of [[TRACKER_USERS_GROUP, read.Id], [TRACKER_GROUPS['Project Engineer'], read.Id], [TRACKER_GROUPS.Manager, manager.Id]]) {
      const group = unwrap(await this.get(`/_api/web/sitegroups/getbyname('${quote(title)}')?$select=Id,Title`));
      if (!positiveId(group?.Id) || group.Title !== title) throw new Error(`A site owner must configure the SharePoint group ${title} before repairing project access.`);
      grants.push({ principalId: Number(group.Id), roleId: Number(roleId) });
    }
    return grants;
  }
  matchesBaseline(rows, grants) {
    return grants.every(grant => rows.some(row => Number(row.PrincipalId) === grant.principalId && Number(row.Member.PrincipalType) === 8 &&
      (row.RoleDefinitionBindings.results || row.RoleDefinitionBindings).some(binding => Number(binding.Id) === grant.roleId)));
  }
  async assignments(scope) {
    const rows = await this.pages(`${scope}/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/PrincipalType,RoleDefinitionBindings/Id&$top=500`);
    for (const row of rows) {
      const bindings = row.RoleDefinitionBindings?.results || row.RoleDefinitionBindings;
      if (!positiveId(row.PrincipalId) || !positiveId(row.Member?.PrincipalType) || !Array.isArray(bindings) || bindings.some(binding => !positiveId(binding.Id))) {
        throw new Error(`SharePoint returned incomplete permission assignments for ${scope}. No matching-permission shortcut was used.`);
      }
    }
    return rows;
  }
  matchesEngineer(rows, principalId, roleId) {
    const desired = engineerPrincipalSet(principalId);
    const engineers = new Set(rows.filter(row => Number(row.Member.PrincipalType) === 1 &&
      (row.RoleDefinitionBindings.results || row.RoleDefinitionBindings).some(binding => Number(binding.Id) === Number(roleId)))
      .map(row => Number(row.PrincipalId)));
    return engineers.size === desired.size && [...engineers].every(id => desired.has(id));
  }
  async applyScope(scope, principalId, roleId, baseline = null) {
    const desired = engineerPrincipalSet(principalId);
    baseline ??= await this.baselineGrants();
    const info = unwrap(await this.get(`${scope}?$select=HasUniqueRoleAssignments`));
    if (typeof info.HasUniqueRoleAssignments !== 'boolean') throw new Error('Could not verify permission inheritance. No scope changes were made.');
    if (!info.HasUniqueRoleAssignments) await this.post(`${scope}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`);
    const bindings = row => row.RoleDefinitionBindings?.results || row.RoleDefinitionBindings || [];
    const hasContribute = row => bindings(row).some(binding => Number(binding.Id) === Number(roleId));
    const current = await this.assignments(scope);
    let changed = false;
    // Add baseline roles without replacing any existing owner or manual grants.
    for (const grant of baseline) {
      if (!this.matchesBaseline(current, [grant])) {
        await this.post(`${scope}/roleassignments/addroleassignment(principalid=${grant.principalId},roledefid=${grant.roleId})`);
        changed = true;
      }
    }
    // These app scopes reserve direct-user Contribute for assigned engineers.
    // Preserve retained teammates, groups, Full Control, and other manual roles.
    for (const row of current) {
      if (Number(row.Member?.PrincipalType) === 1 && !desired.has(Number(row.PrincipalId)) && hasContribute(row)) {
        await this.post(`${scope}/roleassignments/removeroleassignment(principalid=${row.PrincipalId},roledefid=${roleId})`);
        changed = true;
      }
    }
    for (const id of desired) {
      if (!current.some(row => Number(row.PrincipalId) === id && Number(row.Member?.PrincipalType) === 1 && hasContribute(row))) {
        await this.post(`${scope}/roleassignments/addroleassignment(principalid=${id},roledefid=${roleId})`);
        changed = true;
      }
    }
    const checked = changed ? await this.assignments(scope) : current;
    if (!this.matchesBaseline(checked, baseline)) throw new Error('SharePoint did not verify baseline group access. Retry Apply tracker permissions.');
    if (!this.matchesEngineer(checked, desired, roleId)) {
      throw new Error('SharePoint did not verify every project engineer permission change. Retry Apply tracker permissions.');
    }
  }
  async metadata(key) {
    return unwrap(await this.get(`${this.root(key)}?$select=EnableFolderCreation,RootFolder/ServerRelativeUrl&$expand=RootFolder`));
  }
  async enableFolders() {
    for (const key of CHILDREN) {
      if (!(await this.metadata(key)).EnableFolderCreation) {
        await this.post(this.root(key), { headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' }, body: { EnableFolderCreation: true } });
        if (!(await this.metadata(key)).EnableFolderCreation) throw new Error(`Could not enable project folders in ${key}.`);
      }
    }
  }
  folderName(project) {
    // The persisted numeric project ID is stable, path-safe, and collision-free.
    if (!positiveId(project.spId)) throw new Error('Save the project before preparing access.');
    return `tracker-project-${project.spId}`;
  }
  async folder(key, project, create = false) {
    if (typeof project.projectKey !== 'string' || !project.projectKey.trim()) throw new Error('The project has no identity key. Restore its ProjectKey before preparing access.');
    const meta = await this.metadata(key);
    if (!meta.EnableFolderCreation) throw new Error('A site owner must run Apply tracker permissions once to enable project folders.');
    const root = meta.RootFolder?.ServerRelativeUrl;
    if (!root) throw new Error(`Could not resolve the ${key} list folder.`);
    const name = this.folderName(project);
    const path = `${root}/${name}`;
    // List folders are list items. Resolve them in the list itself: some hosts
    // return no usable item ID from GetFolder.../ListItemAllFields.
    const lookup = await this.get(`${this.root(key)}/items?$select=Id&$filter=FileRef eq '${quote(path)}'&$top=2`);
    const matches = lookup?.value || lookup?.d?.results || lookup?.results;
    if (!Array.isArray(matches)) throw new Error(`SharePoint did not return list records while finding the ${key} project folder at ${path}. No folder was created.`);
    if (matches.length > 1 || lookup['@odata.nextLink'] || lookup['odata.nextLink'] || lookup.d?.__next) throw new Error(`More than one result was returned for the ${key} project folder at ${path}. No folder was selected.`);
    let id;
    if (matches.length === 1) id = matches[0].Id;
    else {
      if (!create) throw new Error(`The ${key} project folder is not set up. Run Apply tracker permissions first.`);
      id = await this.addInFolder(key, root, { Title: name, ProjectKey: project.projectKey }, name);
    }
    if (!positiveId(id)) throw new Error(`Could not identify the ${key} project folder at ${path}.`);
    const scope = `${this.root(key)}/items(${Number(id)})`;
    const read = async () => unwrap(await this.get(`${scope}?$select=Id,FileSystemObjectType,FileRef,ProjectKey`, { Accept: 'application/json;odata=minimalmetadata' }));
    const identityError = detail => new Error(`Project folder identity does not match in ${key} at ${path}. ${detail} No permissions were changed for this folder.`);
    const validateFolder = row => {
      if (Number(row.Id) !== Number(id) || ![1, '1'].includes(row.FileSystemObjectType) || String(row.FileRef || '').toLowerCase() !== path.toLowerCase()) {
        throw identityError('SharePoint did not verify the expected folder type, item ID, and path.');
      }
    };
    let row = await read();
    validateFolder(row);
    if (row.ProjectKey !== project.projectKey) {
      // A failed first setup can leave an empty infrastructure folder without
      // its custom key. Only the setup path may finish that initialization.
      // Never overwrite another project's key or adopt a populated folder.
      const unassigned = row.ProjectKey === null || row.ProjectKey === '';
      if (!create || !unassigned) throw identityError(`Expected project key ${project.projectKey}; received ${row.ProjectKey == null ? '(missing)' : row.ProjectKey}. Rerun setup only after resolving any conflicting folder data.`);
      // Computed ItemChildCount/FolderChildCount columns are not REST item
      // properties on every list. Use the associated Folder resource, and
      // only when recovering a blank key; absent counts never mean zero.
      const folderInfo = unwrap(await this.get(`${scope}/Folder?$select=ServerRelativeUrl,ItemCount`));
      if (String(folderInfo?.ServerRelativeUrl || '').toLowerCase() !== path.toLowerCase() || ![0, '0'].includes(folderInfo?.ItemCount)) {
        throw identityError('SharePoint did not verify that the unassigned folder is empty. Resolve its contents before retrying setup.');
      }
      const parent = unwrap(await this.get(`${this.root('projects')}/items(${project.spId})?$select=Id,ProjectKey`));
      if (Number(parent.Id) !== Number(project.spId) || parent.ProjectKey !== project.projectKey) throw identityError('The saved parent project could not be verified.');
      const etag = row['odata.etag'] || row['@odata.etag'] || row.__metadata?.etag;
      if (!etag || etag === '*') throw identityError('SharePoint did not return a concurrency token for the empty folder.');
      await this.post(scope, { body: { ProjectKey: project.projectKey }, headers: { 'IF-MATCH': etag, 'X-HTTP-Method': 'MERGE' } });
      row = await read();
      validateFolder(row);
      if (row.ProjectKey !== project.projectKey) throw identityError('SharePoint did not save the folder project key.');
    }
    return { path, scope };
  }
  async addInFolder(key, path, fields, leafName) {
    // UsingPath takes ResourcePath objects for BOTH FolderPath and LeafName.
    // A plain LeafName string produces SharePoint's PrimitiveValue/StartObject 400.
    const result = await this.post(`${this.root(key)}/AddValidateUpdateItemUsingPath`, { body: {
      listItemCreateInfo: { FolderPath: { DecodedUrl: `${new URL(this.store.webUrl).origin}${path}` }, UnderlyingObjectType: leafName ? 1 : 0, ...(leafName ? { LeafName: { DecodedUrl: leafName } } : {}) },
      formValues: this.store.formValues(fields), bNewDocumentUpdate: false,
    } });
    const rows = result?.value || result?.d?.AddValidateUpdateItemUsingPath?.results || result?.d?.AddValidateUpdateItemUsingPath || result?.AddValidateUpdateItemUsingPath?.results || result?.AddValidateUpdateItemUsingPath || [];
    const failed = rows.find(row => row.HasException || row.ErrorMessage);
    if (failed) throw new Error(`SharePoint rejected ${failed.FieldName}: ${failed.ErrorMessage || 'validation failed'}`);
    const id = Number(rows.find(row => /^(id)$/i.test(row.FieldName))?.FieldValue || rows.find(row => row.ItemId > 0)?.ItemId);
    if (!positiveId(id)) throw new Error('SharePoint did not return the new record ID. Reload before retrying.');
    return id;
  }
  async syncProject(project, users, onProgress = () => {}, baseline = null) {
    baseline ??= await this.baselineGrants();
    const engineerIds = [];
    if (!project.archived) {
      const team = projectEngineers(project);
      const resolved = [];
      for (const engineer of team) {
        // Name-only legacy assignments can be displayed, but never authorize access.
        if (!engineer.loginName && !engineer.email) continue;
        const matches = users.filter(row => isOwnedByUser({ ownerKey: engineer.loginName, ownerEmail: engineer.email }, row));
        if (matches.length > 1 || (Array.isArray(project.engineers) && matches.length === 0)) {
          throw new Error(`Could not uniquely verify ${engineer.name || engineer.email || engineer.loginName} in the Tracker directory. Check the project team before applying access.`);
        }
        const user = matches[0];
        // An intentional demotion to Viewer must not be undone by a saved assignment.
        // Managers already receive the baseline Manager group grant.
        if (!user || user.role === 'Viewer' || user.role === 'Manager') continue;
        if (user.role !== 'Project Engineer' || !user.loginName?.trim()) {
          throw new Error(`Verify ${user.title || engineer.name}'s Project Engineer role before assigning project access.`);
        }
        if (!resolved.some(row => row.loginName.toLowerCase() === user.loginName.toLowerCase())) resolved.push(user);
      }
      // Verify the entire team before any scope is changed. A missing teammate
      // must never be treated as permission to remove their existing grants.
      if (resolved.length) {
        const group = await this.group('Project Engineer');
        for (const engineer of resolved) {
          const member = await this.membership(group, engineer.loginName);
          if (!positiveId(member?.Id) || engineerIds.includes(Number(member.Id))) {
            throw new Error(`Save ${engineer.title}'s Project Engineer role before assigning project access. SharePoint did not uniquely verify every engineer.`);
          }
          engineerIds.push(Number(member.Id));
        }
      }
    }
    // Keep the scalar contract for legacy single-assignee callers and retries.
    const principalId = engineerIds.length > 1 ? engineerIds : engineerIds[0] || null;
    const roleId = await this.contributionRole();
    onProgress(`Updating access: ${project.title || project.projectKey}`);
    await this.applyScope(`${this.root('projects')}/items(${project.spId})`, principalId, roleId, baseline);
    for (const key of CHILDREN) {
      const folder = await this.folder(key, project, true);
      await this.applyScope(folder.scope, principalId, roleId, baseline);
      // Include archived records: former engineers must not retain their ACLs.
      const rows = await this.pages(`${this.root(key)}/items?$select=Id,FileSystemObjectType,HasUniqueRoleAssignments,FileDirRef&$filter=ProjectKey eq '${quote(project.projectKey)}'&$top=500`);
      const candidates = rows.filter(row => ![1, '1'].includes(row.FileSystemObjectType) && (row.HasUniqueRoleAssignments || row.FileDirRef !== folder.path));
      let unchanged = 0, updated = 0;
      // Bound concurrent reads to four. Mutations remain sequential, and each
      // batch must be readable before any item in that batch is changed.
      for (let offset = 0; offset < candidates.length; offset += 4) {
        const batch = candidates.slice(offset, offset + 4);
        onProgress(`Checking ${key} access: ${project.title || project.projectKey} · ${offset + 1}–${offset + batch.length} of ${candidates.length}`);
        const checks = await Promise.allSettled(batch.map(async row => ({
          row, scope: `${this.root(key)}/items(${row.Id})`,
          acl: row.HasUniqueRoleAssignments ? await this.assignments(`${this.root(key)}/items(${row.Id})`) : null,
        })));
        const failed = checks.find(result => result.status === 'rejected');
        if (failed) throw failed.reason;
        for (const { value: { row, scope, acl } } of checks) {
          if (acl && this.matchesEngineer(acl, principalId, roleId) && this.matchesBaseline(acl, baseline)) { unchanged++; continue; }
          onProgress(`Updating ${key} access: ${project.title || project.projectKey} · item ${row.Id}`);
          await this.applyScope(scope, principalId, roleId, baseline);
          updated++;
        }
        onProgress(`Checked ${key} access: ${project.title || project.projectKey} · ${offset + batch.length} of ${candidates.length}; ${unchanged} already correct, ${updated} updated`);
      }
    }
  }
  async createChild(key, fields) {
    const matches = await this.pages(`${this.root('projects')}/items?$select=Id,ProjectKey&$filter=ProjectKey eq '${quote(fields.ProjectKey)}'&$top=2`);
    if (matches.length !== 1) throw new Error('The parent project could not be resolved uniquely.');
    const project = { spId: matches[0].Id, projectKey: matches[0].ProjectKey };
    const folder = await this.folder(key, project);
    return this.addInFolder(key, folder.path, fields);
  }
}
