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
    await expect(permissions.folder('tasks', { spId: 1 }, true)).rejects.toThrow('site owner');
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
