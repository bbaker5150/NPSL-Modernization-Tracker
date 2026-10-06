import { isOwnedByUser } from './identity';

export const TRACKER_GROUPS = {
  Viewer: 'Tracker Viewers',
  'Project Engineer': 'Tracker Project Engineers',
  Manager: 'Tracker Managers',
};
const CHILDREN = ['tasks', 'updates', 'risks'];
const suffixes = { projects: 'Projects', tasks: 'Tasks', updates: 'Updates', risks: 'Risks' };
const unwrap = body => body?.d || body;
const values = body => body?.value || body?.d?.results || body?.results || [];
const quote = value => encodeURIComponent(String(value).replace(/'/g, "''"));
const positiveId = value => Number.isInteger(Number(value)) && Number(value) > 0;

// No privileged service identity: every operation runs as the signed-in person.
// Group ownership and the four list ACLs must be configured by the site owner.
export class TrackerPermissions {
  constructor(store) { this.store = store; }
  root(key) { return `/_api/web/lists/getbytitle('${quote(`${String(this.store.prefix).replace(/[^A-Za-z0-9]/g, '') || 'Modernization'}${suffixes[key]}`)}')`; }
  get(path) { return this.store.get(path); }
  post(path, options) { return this.store.post(path, options); }
  async pages(path) {
    const rows = [];
    while (path) {
      const body = await this.get(path);
      rows.push(...values(body));
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
  async syncGroups(person, role) {
    if (!person.loginName) throw new Error('Select a resolved SharePoint identity.');
    // Read every membership before changing any. Failure is never interpreted as absence.
    const state = [];
    for (const name of Object.keys(TRACKER_GROUPS)) {
      const group = await this.group(name);
      state.push({ name, group, member: await this.membership(group, person.loginName) });
    }
    const target = state.find(entry => entry.name === role);
    if (role && !target) throw new Error('Select a valid tracker role.');
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
  async contributionRole() {
    const roles = await this.pages('/_api/web/roledefinitions?$select=Id,RoleTypeKind&$filter=RoleTypeKind eq 3');
    const role = roles.find(row => row.RoleTypeKind === 3);
    if (!positiveId(role?.Id)) throw new Error('SharePoint Contribute permission level is unavailable.');
    return role.Id;
  }
  async assignments(scope) {
    return this.pages(`${scope}/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/PrincipalType,RoleDefinitionBindings/Id&$top=500`);
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
    const meta = await this.metadata(key);
    if (!meta.EnableFolderCreation) throw new Error('A site owner must run Apply tracker permissions once to enable project folders.');
    const root = meta.RootFolder?.ServerRelativeUrl;
    if (!root) throw new Error(`Could not resolve the ${key} list folder.`);
    const name = this.folderName(project);
    const path = `${root}/${name}`;
    const api = `/_api/web/GetFolderByServerRelativePath(decodedurl='${quote(path)}')/ListItemAllFields`;
    try { const row = unwrap(await this.get(`${api}?$select=Id,FileSystemObjectType,ProjectKey`)); if (row.FileSystemObjectType !== 1 || row.ProjectKey !== project.projectKey) throw new Error('Project folder identity does not match.'); return { path, scope: `${this.root(key)}/items(${row.Id})` }; }
    catch (error) { if (error.status !== 404 || !create) throw error; }
    await this.addInFolder(key, root, { Title: name, ProjectKey: project.projectKey }, name);
    // Resolve the folder again after creation; never treat an ordinary item as a folder.
    return this.folder(key, project);
  }
  async addInFolder(key, path, fields, leafName) {
    const result = await this.post(`${this.root(key)}/AddValidateUpdateItemUsingPath`, { body: {
      listItemCreateInfo: { FolderPath: { DecodedUrl: `${new URL(this.store.webUrl).origin}${path}` }, UnderlyingObjectType: leafName ? 1 : 0, ...(leafName ? { LeafName: leafName } : {}) },
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
      for (const row of rows.filter(row => row.FileSystemObjectType !== 1)) {
        if (row.HasUniqueRoleAssignments || row.FileDirRef !== folder.path) {
          onProgress(`Updating ${key} access: ${project.title || project.projectKey} · item ${row.Id}`);
          await this.applyScope(`${this.root(key)}/items(${row.Id})`, principalId, roleId);
        }
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
