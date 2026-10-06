import { isOwnedByUser } from './identity';

export const TRACKER_GROUPS = {
  Viewer: 'Tracker Viewers',
  'Project Engineer': 'Tracker Project Engineers',
  Manager: 'Tracker Managers',
};
const CHILDREN = ['tasks', 'updates', 'risks'];
const suffixes = { projects: 'Projects', tasks: 'Tasks', updates: 'Updates', risks: 'Risks' };
const unwrap = body => body?.d || body;
const quote = value => encodeURIComponent(String(value).replace(/'/g, "''"));
const positiveId = value => Number.isInteger(Number(value)) && Number(value) > 0;

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
    const group = unwrap(await this.get(`/_api/web/sitegroups/getbyname('${quote(title)}')?$select=Id,Title`));
    if (!positiveId(group?.Id)) throw new Error(`Create the SharePoint group ${title} before assigning this role.`);
    return group;
  }
  async membership(group, login) {
    return (await this.pages(`/_api/web/sitegroups(${group.Id})/users?$select=Id,LoginName&$filter=LoginName eq '${quote(login)}'`))
      .find(row => row.LoginName?.toLowerCase() === login.toLowerCase());
  }
  async syncGroups(person, role, { beforeChange } = {}) {
    if (!person.loginName) throw new Error('Select a resolved SharePoint identity.');
    // Read every membership before changing any. Failure is never interpreted as absence.
    const state = [];
    for (const name of Object.keys(TRACKER_GROUPS)) {
      const group = await this.group(name);
      state.push({ name, group, member: await this.membership(group, person.loginName) });
    }
    const target = state.find(entry => entry.name === role);
    if (role && !target) throw new Error('Select a valid tracker role.');
    // Let role changes revoke engineer grants using the live memberships we
    // just read, before removing groups or granting the replacement role.
    if (beforeChange) await beforeChange(state);
    // Remove stale privileged membership before granting the replacement role.
    for (const entry of state.filter(entry => entry.member && entry !== target)) {
      try {
        await this.post(`/_api/web/sitegroups(${entry.group.Id})/users/removebyid(${entry.member.Id})`);
        if (await this.membership(entry.group, person.loginName)) throw new Error('Membership is still present.');
      } catch (error) { throw new Error(`Could not remove membership in ${entry.group.Title}. Its group owner must complete this role change. ${error.message}`); }
    }
    if (target && !target.member) {
      try { await this.post(`/_api/web/sitegroups(${target.group.Id})/users`, { body: { LoginName: person.loginName } }); }
      catch (error) { throw new Error(`Could not add membership in ${target.group.Title}. Ask its group owner or a site owner. ${error.message}`); }
    }
    if (target && !(await this.membership(target.group, person.loginName))) throw new Error(`SharePoint did not verify ${target.group.Title} membership. Retry the role change.`);
    return target?.group;
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
    const engineers = rows.filter(row => Number(row.Member.PrincipalType) === 1 &&
      (row.RoleDefinitionBindings.results || row.RoleDefinitionBindings).some(binding => Number(binding.Id) === Number(roleId)));
    return engineers.every(row => Number(row.PrincipalId) === Number(principalId)) &&
      (!principalId || engineers.some(row => Number(row.PrincipalId) === Number(principalId)));
  }
  async applyScope(scope, principalId, roleId) {
    const info = unwrap(await this.get(`${scope}?$select=HasUniqueRoleAssignments`));
    if (!info.HasUniqueRoleAssignments) await this.post(`${scope}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`);
    const bindings = row => row.RoleDefinitionBindings?.results || row.RoleDefinitionBindings || [];
    const current = await this.assignments(scope);
    let changed = false;
    // These app scopes reserve direct-user Contribute for the assigned engineer.
    // Preserve groups, Full Control, Read, and any other manually assigned roles.
    for (const row of current) {
      if (row.Member?.PrincipalType === 1 && row.PrincipalId !== principalId && bindings(row).some(binding => binding.Id === roleId)) {
        await this.post(`${scope}/roleassignments/removeroleassignment(principalid=${row.PrincipalId},roledefid=${roleId})`);
        changed = true;
      }
    }
    if (principalId && !current.some(row => row.PrincipalId === principalId && bindings(row).some(binding => binding.Id === roleId))) {
      await this.post(`${scope}/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleId})`);
      changed = true;
    }
    const verified = (changed ? await this.assignments(scope) : current).filter(row => row.Member?.PrincipalType === 1 && bindings(row).some(binding => binding.Id === roleId));
    if (verified.some(row => row.PrincipalId !== principalId) || (principalId && !verified.some(row => row.PrincipalId === principalId))) {
      throw new Error('SharePoint did not verify the project engineer permission change. Retry Apply tracker permissions.');
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
  async syncProject(project, users, onProgress = () => {}) {
    const engineer = users.find(row => row.role === 'Project Engineer' && isOwnedByUser(project, row));
    let principalId = null;
    if (engineer && !project.archived) {
      const group = await this.group('Project Engineer');
      const member = await this.membership(group, engineer.loginName);
      if (!member) throw new Error(`Save ${engineer.title}'s Project Engineer role before assigning project access.`);
      principalId = member.Id;
    }
    const roleId = await this.contributionRole();
    onProgress(`Updating access: ${project.title || project.projectKey}`);
    await this.applyScope(`${this.root('projects')}/items(${project.spId})`, principalId, roleId);
    for (const key of CHILDREN) {
      const folder = await this.folder(key, project, true);
      await this.applyScope(folder.scope, principalId, roleId);
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
          if (acl && this.matchesEngineer(acl, principalId, roleId)) { unchanged++; continue; }
          onProgress(`Updating ${key} access: ${project.title || project.projectKey} · item ${row.Id}`);
          await this.applyScope(scope, principalId, roleId);
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
