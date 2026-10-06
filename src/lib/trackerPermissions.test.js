import { describe, expect, it, vi } from 'vitest';
import { TrackerPermissions, TRACKER_GROUPS } from './trackerPermissions';
import { SharePointStore } from './spStore';
import { authorizedStore } from './access';

function fixture() {
  const groups = Object.entries(TRACKER_GROUPS).map(([role, Title], index) => ({ Id: index + 1, Title, role, members: [] }));
  const calls = [];
  const store = { webUrl: 'https://tenant.sharepoint.com/sites/metsoft', prefix: 'Modernization' };
  store.get = vi.fn(async path => {
    calls.push(['GET', path]);
    const named = groups.find(group => path.includes(`getbyname('${encodeURIComponent(group.Title)}')`));
    if (named) return { Id: named.Id, Title: named.Title };
    const member = path.match(/sitegroups\((\d+)\)\/users/);
    if (member) return { value: groups.find(group => group.Id === Number(member[1])).members };
    throw new Error(`Unexpected GET: ${path}`);
  });
  store.post = vi.fn(async (path, options) => {
    calls.push(['POST', path, options]);
    const group = groups.find(group => path.includes(`sitegroups(${group.Id})`));
    if (path.includes('/removebyid(')) group.members = [];
    else if (group && options?.body?.LoginName) group.members.push({ Id: 11, LoginName: options.body.LoginName });
    else throw new Error(`Unexpected POST: ${path}`);
  });
  return { store, groups, calls, permissions: new TrackerPermissions(store) };
}
const person = { id: 'u1', loginName: 'i:0#.f|membership|engineer@example.com', title: 'Engineer', role: 'Project Engineer' };

describe('tracker group permissions', () => {
  it('replaces tracker role membership, verifies it, and never touches site Members', async () => {
    const { permissions, groups, calls } = fixture();
    groups[0].members = [{ Id: 11, LoginName: person.loginName }];
    const group = await permissions.syncGroups(person, 'Project Engineer');
    expect(group.Title).toBe('Tracker Project Engineers');
    expect(groups[0].members).toEqual([]);
    expect(groups[1].members).toHaveLength(1);
    expect(JSON.stringify(calls)).not.toContain('associatedmembergroup');
    const writes = calls.filter(([method]) => method === 'POST');
    expect(writes[0][1]).toContain('removebyid');
    expect(writes[1][1]).toContain('sitegroups(2)/users');
  });
  it('requires the Manager group owner for promotions and does not pretend they succeeded', async () => {
    const { permissions, store } = fixture();
    store.post.mockRejectedValue(new Error('403 Forbidden'));
    await expect(permissions.syncGroups(person, 'Manager')).rejects.toThrow('Tracker Managers');
  });
  it('does not grant a replacement role if removing the previous role fails', async () => {
    const { permissions, groups, store } = fixture();
    groups[2].members = [{ Id: 11, LoginName: person.loginName }];
    store.post.mockRejectedValue(new Error('403 Forbidden'));
    await expect(permissions.syncGroups(person, 'Viewer')).rejects.toThrow('Could not remove membership');
    expect(store.post).toHaveBeenCalledTimes(1);
  });
  it('removes all managed memberships without adding a fallback group when deleting', async () => {
    const { permissions, groups, store } = fixture();
    groups[0].members = groups[1].members = [{ Id: 11, LoginName: person.loginName }];
    await permissions.syncGroups(person, null);
    expect(groups.every(group => !group.members.length)).toBe(true);
    expect(store.post.mock.calls.every(([path]) => path.includes('removebyid'))).toBe(true);
  });
  it('fails closed before any mutation when membership cannot be read', async () => {
    const { permissions, store } = fixture();
    store.get.mockRejectedValue(new Error('403 Forbidden'));
    await expect(permissions.syncGroups(person, 'Viewer')).rejects.toThrow('403');
    expect(store.post).not.toHaveBeenCalled();
  });
  it('uses real current group membership or site permissions as the current role', async () => {
    const { permissions, store } = fixture();
    store.get.mockImplementation(async path => path.includes('/groups') ? { value: [{ Title: 'Tracker Project Engineers' }] } : { Low: '1', High: '0' });
    expect(await permissions.currentRole({})).toEqual({ siteOwner: false, role: 'Project Engineer' });
    store.get.mockImplementation(async path => path.includes('/groups') ? { value: [] } : { Low: '33554433', High: '0' });
    expect(await permissions.currentRole({})).toEqual({ siteOwner: true, role: 'Manager' });
  });
});

describe('project ACLs and inherited creation', () => {
  it('replaces the old engineer while preserving groups and unrelated permission levels', async () => {
    const { permissions, store } = fixture();
    const scope = '/_api/project/item';
    let acl = [
      { PrincipalId: 11, Member: { PrincipalType: 1 }, RoleDefinitionBindings: [{ Id: 3 }, { Id: 2 }] },
      { PrincipalId: 50, Member: { PrincipalType: 8 }, RoleDefinitionBindings: [{ Id: 3 }] },
      { PrincipalId: 90, Member: { PrincipalType: 1 }, RoleDefinitionBindings: [{ Id: 99 }] },
    ];
    store.get.mockImplementation(async path => path.includes('/roleassignments') ? { value: acl } : { HasUniqueRoleAssignments: false });
    store.post.mockImplementation(async path => {
      if (path.includes('removeroleassignment')) acl = acl.map(row => row.PrincipalId === 11 ? { ...row, RoleDefinitionBindings: [{ Id: 2 }] } : row);
      if (path.includes('addroleassignment')) acl.push({ PrincipalId: 12, Member: { PrincipalType: 1 }, RoleDefinitionBindings: [{ Id: 3 }] });
    });
    await permissions.applyScope(scope, 12, 3);
    expect(store.post.mock.calls[0][0]).toBe(`${scope}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`);
    expect(acl.find(row => row.PrincipalId === 11).RoleDefinitionBindings).toEqual([{ Id: 2 }]);
    expect(acl.find(row => row.PrincipalId === 50).RoleDefinitionBindings).toEqual([{ Id: 3 }]);
    expect(acl.find(row => row.PrincipalId === 90).RoleDefinitionBindings).toEqual([{ Id: 99 }]);
    await permissions.applyScope(scope, 12, 3);
    expect(store.post.mock.calls.filter(([path]) => path.includes('addroleassignment'))).toHaveLength(1);
  });
  it('detects a permission request that SharePoint did not actually apply', async () => {
    const { permissions, store } = fixture();
    store.get.mockImplementation(async path => path.includes('/roleassignments') ? { value: [] } : { HasUniqueRoleAssignments: true });
    store.post.mockResolvedValue({});
    await expect(permissions.applyScope('/scope', 12, 3)).rejects.toThrow('did not verify');
  });
  it('reconciles legacy and archived children but lets new children inherit from the folder', async () => {
    const { permissions } = fixture();
    permissions.group = vi.fn(async () => ({}));
    permissions.membership = vi.fn(async () => ({ Id: 11 }));
    permissions.contributionRole = vi.fn(async () => 3);
    permissions.applyScope = vi.fn();
    permissions.folder = vi.fn(async key => ({ path: `/${key}/tracker-project-5`, scope: `/${key}/items(50)` }));
    permissions.pages = vi.fn(async path => {
      const key = path.includes('Tasks') ? 'tasks' : path.includes('Updates') ? 'updates' : 'risks';
      return [
        { Id: 1, FileSystemObjectType: 0, FileDirRef: `/${key}`, HasUniqueRoleAssignments: false },
        { Id: 2, FileSystemObjectType: 0, FileDirRef: `/${key}/tracker-project-5`, HasUniqueRoleAssignments: false },
        { Id: 3, FileSystemObjectType: 0, FileDirRef: `/${key}`, HasUniqueRoleAssignments: true, Archived: true },
        { Id: 50, FileSystemObjectType: 1 },
      ];
    });
    await permissions.syncProject({ spId: 5, projectKey: 'p', ownerKey: person.loginName }, [person]);
    expect(permissions.applyScope).toHaveBeenCalledTimes(10); // project, three folders, six legacy/unique items
    expect(permissions.applyScope.mock.calls.every(([, id]) => id === 11)).toBe(true);
    expect(permissions.applyScope.mock.calls.some(([scope]) => scope.endsWith('/items(2)'))).toBe(false);
  });
  it('creates tasks in the assigned project folder without requesting list-wide access', async () => {
    const { permissions, store } = fixture();
    permissions.pages = vi.fn(async () => [{ Id: 5, ProjectKey: 'p' }]);
    permissions.folder = vi.fn(async () => ({ path: '/sites/metsoft/Lists/Tasks/tracker-project-5' }));
    store.formValues = fields => Object.entries(fields).map(([FieldName, FieldValue]) => ({ FieldName, FieldValue }));
    store.post.mockResolvedValue({ value: [{ FieldName: 'Id', FieldValue: '22' }] });
    expect(await permissions.createChild('tasks', { Title: 'Task', ProjectKey: 'p' })).toBe(22);
    const [path, { body }] = store.post.mock.calls[0];
    expect(path).toContain('AddValidateUpdateItemUsingPath');
    expect(body.listItemCreateInfo).toEqual({ FolderPath: { DecodedUrl: 'https://tenant.sharepoint.com/sites/metsoft/Lists/Tasks/tracker-project-5' }, UnderlyingObjectType: 0 });
    expect(body.bNewDocumentUpdate).toBe(false);
  });
  it('stops creation if folder setup is missing; never falls back to a root item', async () => {
    const { permissions, store } = fixture();
    permissions.metadata = vi.fn(async () => ({ EnableFolderCreation: false }));
    await expect(permissions.folder('tasks', { spId: 1, projectKey: 'p' }, true)).rejects.toThrow('site owner');
    expect(store.post).not.toHaveBeenCalled();
  });
  it('rejects per-field folder creation failures even on HTTP success', async () => {
    const { permissions, store } = fixture();
    store.formValues = () => [];
    store.post.mockResolvedValue({ value: [{ FieldName: 'Title', HasException: true, ErrorMessage: 'Denied' }] });
    await expect(permissions.addInFolder('tasks', '/folder', {})).rejects.toThrow('Denied');
  });
});

describe('scoped store integration', () => {
  it('does not persist a role when permission synchronization fails', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/metsoft', scopedAccess: true });
    store.directoryRows = vi.fn(async () => []);
    store.syncUserAccess = vi.fn(async () => { throw new Error('Group ownership required'); });
    store.create = vi.fn(); store.update = vi.fn();
    await expect(store.saveUser(person)).rejects.toThrow('Group ownership');
    expect(store.create).not.toHaveBeenCalled(); expect(store.update).not.toHaveBeenCalled();
  });
  it('does not archive a user if revoking access fails', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/metsoft', scopedAccess: true });
    store.directoryRows = vi.fn(async () => [{ ...person, spId: 7 }]);
    store.syncUserAccess = vi.fn(async () => { throw new Error('Cannot revoke'); });
    store.update = vi.fn();
    await expect(store.recycle('users', 7)).rejects.toThrow('Cannot revoke');
    expect(store.update).not.toHaveBeenCalled();
  });
  it('reconciles all projects on removal, including interrupted former assignments', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/metsoft', scopedAccess: true });
    store.permissionProjects = vi.fn(async () => [{ spId: 1, ownerKey: person.loginName }, { spId: 2, ownerKey: 'someone-else' }]);
    store.permissions.syncProject = vi.fn(); store.permissions.syncGroups = vi.fn();
    await store.syncUserAccess(person, null, [person]);
    expect(store.permissions.syncProject).toHaveBeenCalledTimes(2);
    expect(store.permissions.syncProject.mock.calls.every(([, users]) => users.length === 0)).toBe(true);
    expect(store.permissions.syncGroups).toHaveBeenCalledWith(person, null);
  });
  it('does not write on read-only startup and rejects testing promotion', async () => {
    const raw = {
      scopedAccess: true,
      currentUser: async () => ({ loginName: person.loginName }),
      load: async () => ({ users: [{ ...person, role: 'Viewer' }], projects: [], tasks: [], updates: [], risks: [] }),
      saveUser: vi.fn(),
    };
    const store = authorizedStore(raw, { testingManagerPassword: 'admin123' });
    expect((await store.registerCurrentUser()).role).toBe('Viewer');
    await expect(store.activateTestingManager('admin123')).rejects.toThrow('disabled');
    expect(raw.saveUser).not.toHaveBeenCalled();
  });
  it('does not let an engineer reassign project ownership through edited input', async () => {
    const project = { id: 'p', spId: 1, projectKey: 'p', ownerName: person.title, ownerKey: person.loginName, ownerEmail: 'engineer@example.com' };
    const raw = {
      currentUser: async () => person,
      load: async () => ({ users: [person], projects: [project], tasks: [], updates: [], risks: [] }),
      saveProject: vi.fn(async row => row),
    };
    const result = await authorizedStore(raw).saveProject({ ...project, ownerKey: 'someone-else', ownerName: 'New engineer', ownerEmail: 'other@example.com' });
    expect(result.ownerKey).toBe(person.loginName);
    expect(result.ownerEmail).toBe(project.ownerEmail);
  });
  it('keeps permission folders out of normal tracker data', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/metsoft' });
    store.get = vi.fn(async () => ({ value: [{ Id: 1, FileSystemObjectType: 1 }, { Id: 2, FileSystemObjectType: 0 }] }));
    expect(await store.listItems('tasks', [], row => row.Id)).toEqual([2]);
  });
  it('invites through the tracker role group and preserves one direct-page email request', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/metsoft', scopedAccess: true });
    store.permissions.syncGroups = vi.fn(async () => ({ Id: 2, Title: 'Tracker Project Engineers' }));
    store.get = vi.fn(async path => {
      if (path.includes('sitegroups(2)/users')) return { value: [{ LoginName: person.loginName }] };
      if (path.includes('getusereffectivepermissions')) return { GetUserEffectivePermissions: { Low: '33' } };
      if (path.includes('roledefinitions')) return { value: [{ Id: 124, RoleTypeKind: 2 }] };
      if (path.includes('GetFileByServerRelativePath')) return { Exists: true, Level: 1 };
      throw new Error(`Unexpected GET: ${path}`);
    });
    store.post = vi.fn(async () => ({ StatusCode: 0 }));
    const page = `${store.webUrl}/SitePages/Modernization-Tracker.aspx`;
    expect(await store.shareSiteAccess(person, 'Project Engineer', page)).toEqual({ access: 'Tracker Project Engineers', emailRequested: true });
    expect(store.get.mock.calls.some(([path]) => path.includes('associatedmembergroup'))).toBe(false);
    expect(store.post).toHaveBeenCalledTimes(1);
    expect(store.post.mock.calls[0]).toEqual(['/_api/SP.Web.ShareObject', { body: expect.objectContaining({ url: page, sendEmail: true, propagateAcl: false }) }]);
    store.post.mockRejectedValue(new Error('Sharing denied'));
    await expect(store.shareSiteAccess(person, 'Project Engineer', page)).rejects.toThrow('membership was verified, but the invitation email was not confirmed');
  });
  it('does not trust a directory Manager role when current SharePoint membership is Viewer', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/metsoft', scopedAccess: true });
    store.currentUser = vi.fn(async () => person);
    store.permissions.currentRole = vi.fn(async () => ({ role: 'Viewer', siteOwner: false }));
    store.listItems = vi.fn(async key => key === 'users' ? [{ ...person, role: 'Manager' }] : []);
    expect((await store.load()).users[0].role).toBe('Viewer');
  });
  it('reuses a project saved before its ACL setup failed', async () => {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/metsoft', scopedAccess: true });
    store.currentUser = vi.fn(async () => person);
    store.permissions.currentRole = vi.fn(async () => ({ role: 'Manager' }));
    store.permissionProjects = vi.fn(async () => [{ id: 'project', spId: 8 }]);
    store.directoryRows = vi.fn(async () => []);
    store.permissions.syncProject = vi.fn(); store.create = vi.fn(); store.update = vi.fn();
    expect((await store.saveProject({ id: 'project', projectKey: 'p', title: 'Retry' })).spId).toBe(8);
    expect(store.create).not.toHaveBeenCalled();
    expect(store.permissions.syncProject).toHaveBeenCalledWith(expect.objectContaining({ spId: 8 }), []);
  });
});

// Resolve folder records without the broken folder-navigation endpoint, and
// reproduce interrupted setup that created a folder without ProjectKey.
function folderFixture({ exists = true, key = 'p', type = 1 } = {}) {
  const { store, permissions } = fixture();
  const project = { spId: 5, projectKey: 'p' };
  const root = '/sites/metsoft/Lists/ModernizationTasks';
  const path = `${root}/tracker-project-5`;
  const row = { Id: 22, FileSystemObjectType: type, FileRef: path, ProjectKey: key, ItemChildCount: '0', FolderChildCount: '0', 'odata.etag': '"2"' };
  permissions.metadata = vi.fn(async () => ({ EnableFolderCreation: true, RootFolder: { ServerRelativeUrl: root } }));
  store.formValues = fields => Object.entries(fields).map(([FieldName, FieldValue]) => ({ FieldName, FieldValue }));
  store.get.mockImplementation(async url => {
    if (url.includes('/items?') && url.includes('$filter=FileRef')) return { value: exists ? [{ Id: 22 }] : [] };
    if (url.includes('GetFolderByServerRelativePath')) throw new Error('Do not use the broken folder-navigation endpoint');
    if (url.includes("getbytitle('ModernizationProjects')/items(5)")) return { Id: 5, ProjectKey: 'p' };
    if (url.includes('/items(22)?')) return { ...row };
    throw new Error(`Unexpected GET: ${url}`);
  });
  store.post.mockImplementation(async (url, options) => {
    if (url.includes('AddValidateUpdateItemUsingPath')) {
      exists = true;
      return { value: [{ FieldName: 'Id', FieldValue: '22' }] };
    }
    if (url.endsWith('/items(22)')) {
      if (options.headers['IF-MATCH'] !== row['odata.etag']) throw Object.assign(new Error('Concurrent folder update'), { status: 412 });
      Object.assign(row, options.body);
      return {};
    }
    throw new Error(`Unexpected POST: ${url}`);
  });
  return { store, permissions, project, row, path };
}

describe('project folder identity regression', () => {
  it.each([1, '1'])('reads canonical folder metadata and accepts folder enum %s', async type => {
    const { store, permissions, project, path } = folderFixture({ type });
    expect(await permissions.folder('tasks', project)).toEqual({ path, scope: "/_api/web/lists/getbytitle('ModernizationTasks')/items(22)" });
    expect(store.get).toHaveBeenLastCalledWith(expect.stringContaining('/items(22)?$select=Id,FileSystemObjectType,FileRef,ProjectKey'), { Accept: 'application/json;odata=minimalmetadata' });
    expect(store.get.mock.calls[0][0]).toContain("$filter=FileRef eq '%2Fsites%2Fmetsoft%2FLists%2FModernizationTasks%2Ftracker-project-5'");
    expect(store.get.mock.calls.some(([url]) => url.includes('GetFolderByServerRelativePath'))).toBe(false);
    expect(store.post).not.toHaveBeenCalled();
  });
  it('finishes an empty unassigned folder left by an interrupted setup and verifies the saved key', async () => {
    const { store, permissions, project, row } = folderFixture({ key: null });
    await permissions.folder('tasks', project, true);
    expect(row.ProjectKey).toBe('p');
    expect(store.post).toHaveBeenCalledExactlyOnceWith(expect.stringContaining('/items(22)'), { body: { ProjectKey: 'p' }, headers: { 'IF-MATCH': '"2"', 'X-HTTP-Method': 'MERGE' } });
    await permissions.folder('tasks', project, true);
    expect(store.post).toHaveBeenCalledTimes(1);
  });
  it('verifies a newly created folder by its returned ID and initializes missing metadata', async () => {
    const { store, permissions, project, row } = folderFixture({ exists: false, key: null });
    await permissions.folder('tasks', project, true);
    expect(store.post.mock.calls[0][0]).toContain('AddValidateUpdateItemUsingPath');
    expect(store.post.mock.calls[0][1].body.listItemCreateInfo).toMatchObject({ UnderlyingObjectType: 1, LeafName: 'tracker-project-5' });
    expect(row.ProjectKey).toBe('p');
    expect(store.post.mock.calls.filter(([url]) => url.includes('AddValidateUpdateItemUsingPath'))).toHaveLength(1);
  });
  it.each([
    { ProjectKey: 'another-project' },
    { ProjectKey: null, ItemChildCount: '1' },
    { ProjectKey: null, FolderChildCount: '1' },
    { ProjectKey: undefined },
    { ProjectKey: null, ItemChildCount: undefined },
    { ProjectKey: null, 'odata.etag': undefined },
    { FileSystemObjectType: 0 },
    { FileRef: '/sites/metsoft/Lists/ModernizationTasks/wrong-folder' },
    { Id: 99 },
  ])('rejects ambiguous or unsafe folder identity: %j', async changes => {
    const { store, permissions, project, row } = folderFixture();
    Object.assign(row, changes);
    await expect(permissions.folder('tasks', project, true)).rejects.toThrow('Project folder identity does not match in tasks');
    expect(store.post).not.toHaveBeenCalled();
  });
  it('does not initialize folder metadata during ordinary engineer task creation', async () => {
    const { store, permissions, project } = folderFixture({ key: null });
    await expect(permissions.folder('tasks', project)).rejects.toThrow('Project folder identity');
    expect(store.post).not.toHaveBeenCalled();
  });
  it('stops on a concurrency conflict without overwriting or applying ACLs', async () => {
    const { store, permissions, project } = folderFixture({ key: null });
    store.post.mockRejectedValue(Object.assign(new Error('Concurrent folder update'), { status: 412 }));
    await expect(permissions.folder('tasks', project, true)).rejects.toMatchObject({ status: 412 });
    expect(store.post).toHaveBeenCalledTimes(1);
    expect(store.post.mock.calls[0][1].headers['IF-MATCH']).not.toBe('*');
  });
  it('does not claim success when SharePoint did not save the repaired key', async () => {
    const { store, permissions, project } = folderFixture({ key: null });
    store.post.mockResolvedValue({});
    await expect(permissions.folder('tasks', project, true)).rejects.toThrow('did not save');
  });
  it('does not create a replacement folder on access denied', async () => {
    const { store, permissions, project } = folderFixture();
    store.get.mockRejectedValue(Object.assign(new Error('Access denied'), { status: 403 }));
    await expect(permissions.folder('tasks', project, true)).rejects.toMatchObject({ status: 403 });
    expect(store.post).not.toHaveBeenCalled();
  });
  it('rejects a missing parent project key before requesting folders', async () => {
    const { store, permissions, project } = folderFixture();
    await expect(permissions.folder('tasks', { ...project, projectKey: '' }, true)).rejects.toThrow('no identity key');
    expect(store.get).not.toHaveBeenCalled();
    expect(store.post).not.toHaveBeenCalled();
  });
});

describe('Manager promotion after interrupted folder setup', () => {
  it.each(['Viewer', 'Project Engineer'])('saves and verifies a %s → Manager change through the folder recovery path', async oldRole => {
    const backing = fixture();
    backing.groups.find(group => group.role === oldRole).members = [{ Id: 11, LoginName: person.loginName }];
    const store = new SharePointStore({ webUrl: backing.store.webUrl, scopedAccess: true });
    const previous = { ...person, role: oldRole, spId: 7 };
    const project = { spId: 5, projectKey: 'p', ownerKey: person.loginName, title: 'Migrated project' };
    const folders = Object.fromEntries(['Tasks', 'Updates', 'Risks'].map(suffix => [suffix, {
      Id: 22, FileSystemObjectType: '1', ProjectKey: null,
      FileRef: `/sites/metsoft/Lists/Modernization${suffix}/tracker-project-5`,
      ItemChildCount: '0', FolderChildCount: '0', 'odata.etag': '"2"',
    }]));
    let savedUser;
    store.directoryRows = vi.fn(async () => [previous]);
    store.permissionProjects = vi.fn(async () => [project]);
    store.update = vi.fn(async (_key, _id, fields) => { savedUser = { Id: 7, ...fields }; });
    store.get = vi.fn(async (path, headers) => {
      const suffix = Object.keys(folders).find(name => path.includes(`Modernization${name}`));
      if (path.includes('/roledefinitions')) return { value: [{ Id: 3, RoleTypeKind: 3 }] };
      if (path.includes('/items(7)?')) return savedUser;
      if (path.includes("getbytitle('ModernizationProjects')/items(5)?")) return { Id: 5, ProjectKey: 'p' };
      if (suffix && path.includes('RootFolder')) return { EnableFolderCreation: true, RootFolder: { ServerRelativeUrl: `/sites/metsoft/Lists/Modernization${suffix}` } };
      if (suffix && path.includes('/items?') && path.includes('$filter=FileRef')) return { value: [{ Id: 22 }] };
      if (path.includes('GetFolderByServerRelativePath')) throw new Error('Do not use the broken folder-navigation endpoint');
      if (suffix && path.includes('/items(22)?')) {
        expect(headers.Accept).toContain('minimalmetadata');
        return { ...folders[suffix] };
      }
      if (suffix && path.includes('/items?')) return { value: [] };
      return backing.store.get(path);
    });
    store.post = vi.fn(async (path, options) => {
      const suffix = Object.keys(folders).find(name => path.includes(`Modernization${name}`));
      if (suffix && path.endsWith('/items(22)')) {
        expect(options.headers['IF-MATCH']).toBe('"2"');
        Object.assign(folders[suffix], options.body);
        return {};
      }
      return backing.store.post(path, options);
    });
    // ACL wire behavior is separately covered above; retain the real project,
    // folder, group membership, and user persistence orchestration here.
    store.permissions.applyScope = vi.fn(async (scope, principalId) => {
      expect(principalId).toBeNull(); // remove former direct engineer rights
      if (scope.includes('/items(22)')) {
        const suffix = Object.keys(folders).find(name => scope.includes(`Modernization${name}`));
        expect(folders[suffix].ProjectKey).toBe('p'); // verified before ACL use
      }
    });
    const saved = await store.saveUser({ ...previous, role: 'Manager' });
    expect(saved.role).toBe('Manager');
    expect(store.permissions.applyScope).toHaveBeenCalledTimes(4);
    expect(backing.groups.find(group => group.role === 'Manager').members).toEqual([{ Id: 11, LoginName: person.loginName }]);
    expect(backing.groups.filter(group => group.role !== 'Manager').every(group => group.members.length === 0)).toBe(true);
    expect(store.update).toHaveBeenCalledExactlyOnceWith('users', 7, expect.objectContaining({ AppRole: 'Manager' }));
  });
});

describe('list-based project folder lookup', () => {
  it.each([
    {},
    { value: [{ Id: 22 }, { Id: 23 }] },
    { value: [], 'odata.nextLink': 'https://tenant.sharepoint.com/sites/metsoft/_api/next' },
    { value: [{ Id: null }] },
  ])('refuses malformed, incomplete, or ambiguous lookup responses without creating a folder: %j', async lookup => {
    const { store, permissions, project } = folderFixture();
    store.get.mockResolvedValueOnce(lookup);
    await expect(permissions.folder('tasks', project, true)).rejects.toThrow(/No folder|Could not identify/);
    expect(store.post).not.toHaveBeenCalled();
  });
  it('does not treat a list-level 404 as a missing project folder', async () => {
    const { store, permissions, project } = folderFixture();
    store.get.mockRejectedValueOnce(Object.assign(new Error('List unavailable'), { status: 404 }));
    await expect(permissions.folder('tasks', project, true)).rejects.toMatchObject({ status: 404 });
    expect(store.post).not.toHaveBeenCalled();
  });
});
