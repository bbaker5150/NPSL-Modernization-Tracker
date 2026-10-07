import { expect, it, vi } from 'vitest';
import { TrackerPermissions } from './trackerPermissions';
import { SharePointStore } from './spStore';

const groups = [
  { Id: 20, Title: 'Metrology App User' },
  { Id: 21, Title: 'Tracker Project Engineers' },
  { Id: 22, Title: 'Tracker Managers' },
];
const roles = [{ Id: 2, RoleTypeKind: 2 }, { Id: 7, Name: 'Tracker Project Access Manager', BasePermissions: { Low: '33554447' } }];
const row = (id, role, type = 8) => ({ PrincipalId: id, Member: { PrincipalType: type }, RoleDefinitionBindings: [{ Id: role }] });
function fixture() {
  const state = new Map();
  const acl = scope => { if (!state.has(scope)) state.set(scope, [row(99, 1073741829), row(88, 19, 1)]); return state.get(scope); };
  const store = { prefix: 'Modernization', get: vi.fn(async path => {
    if (path.includes('/roledefinitions?')) return { value: structuredClone(roles) };
    if (path.includes('/sitegroups/getbyname')) {
      const group = groups.find(g => path.includes(encodeURIComponent(g.Title)));
      if (!group) throw new Error('Group missing');
      return group;
    }
    if (path.includes('/roleassignments?')) return { value: structuredClone(acl(path.split('/roleassignments?')[0])) };
    if (path.endsWith('?$select=HasUniqueRoleAssignments')) return { HasUniqueRoleAssignments: true };
    throw new Error(`Unexpected GET ${path}`);
  }), post: vi.fn(async path => {
    const scope = path.split('/roleassignments/')[0];
    const [, id, role] = path.match(/principalid=(\d+),roledefid=(\d+)/) || [];
    if (!id) throw new Error(`Unexpected POST ${path}`);
    if (path.includes('/addroleassignment')) acl(scope).push(row(Number(id), Number(role), groups.some(g => g.Id === Number(id)) ? 8 : 1));
    else state.set(scope, acl(scope).filter(r => !(r.PrincipalId === Number(id) && r.RoleDefinitionBindings[0].Id === Number(role))));
  }) };
  return { store, permissions: new TrackerPermissions(store), state, acl };
}
it('repairs owner-only access, preserves unrelated grants, and retries without writes', async () => {
  const { permissions, store, acl } = fixture();
  await permissions.applyScope('/project', 55, 3);
  expect(acl('/project')).toEqual(expect.arrayContaining([row(99, 1073741829), row(88, 19, 1), row(20, 2), row(21, 2), row(22, 7), row(55, 3, 1)]));
  expect(store.post).toHaveBeenCalledTimes(4);
  store.post.mockClear();
  await permissions.applyScope('/project', 55, 3);
  expect(store.post).not.toHaveBeenCalled();
});
it('retains verified baseline grants while reconciling the assigned engineer', async () => {
  const { permissions, acl } = fixture();
  await permissions.applyScope('/project', 55, 3);
  await permissions.applyScope('/project', 56, 3);
  expect(acl('/project')).not.toContainEqual(row(55, 3, 1));
  expect(acl('/project')).toContainEqual(row(56, 3, 1));
  expect(permissions.matchesBaseline(acl('/project'), await permissions.baselineGrants())).toBe(true);
});
it('does not claim success if SharePoint ignores a baseline grant', async () => {
  const { permissions, store } = fixture();
  store.post.mockResolvedValue({});
  await expect(permissions.applyScope('/project', null, 3)).rejects.toThrow('did not verify baseline');
});
it.each(['group', 'role'])('stops before ACL changes if required %s configuration is missing', async missing => {
  const { permissions, store } = fixture();
  const get = store.get.getMockImplementation();
  store.get.mockImplementation(path => missing === 'group' && path.includes('Metrology%20App%20User') ? Promise.reject(new Error('Group missing')) : missing === 'role' && path.includes('/roledefinitions?') ? { value: [roles[0]] } : get(path));
  await expect(permissions.syncProject({ spId: 1, projectKey: 'p' }, [])).rejects.toThrow();
  expect(store.post).not.toHaveBeenCalled();
});
it('repairs projects, folders, and unique legacy items even when engineer grants already match', async () => {
  const { permissions, store, acl } = fixture();
  permissions.contributionRole = vi.fn(async () => 3);
  permissions.folder = vi.fn(async key => ({ path: `/${key}/folder`, scope: `/${key}/folderScope` }));
  const pages = permissions.pages.bind(permissions);
  permissions.pages = vi.fn(path => path.includes('/items?$select=Id,FileSystemObjectType') ? Promise.resolve([
    { Id: 10, FileSystemObjectType: 0, HasUniqueRoleAssignments: true, FileDirRef: '/legacy' },
    { Id: 11, FileSystemObjectType: 0, HasUniqueRoleAssignments: false, FileDirRef: path.includes('Tasks') ? '/tasks/folder' : path.includes('Updates') ? '/updates/folder' : '/risks/folder' },
  ]) : pages(path));
  await permissions.syncProject({ spId: 1, projectKey: 'p' }, []);
  expect(store.post).toHaveBeenCalledTimes(21); // 3 grants × (project + 3 folders + 3 legacy items)
  expect(store.post.mock.calls.some(([path]) => path.includes('/items(11)'))).toBe(false);
  expect(acl('/tasks/folderScope')).toContainEqual(row(20, 2));
  store.post.mockClear();
  await permissions.syncProject({ spId: 1, projectKey: 'p' }, []);
  expect(store.post).not.toHaveBeenCalled();
});
it('preflights setup baseline before folder or role membership mutations', async () => {
  const store = new SharePointStore({ webUrl: 'https://tenant.test/sites/metsoft', scopedAccess: true });
  store.currentUser = vi.fn(async () => ({}));
  store.permissions.currentRole = vi.fn(async () => ({ role: 'Manager', siteOwner: true }));
  store.permissions.baselineGrants = vi.fn(async () => { throw new Error('Baseline missing'); });
  store.permissions.enableFolders = vi.fn();
  store.permissions.syncGroups = vi.fn();
  await expect(store.prepareTrackerAccess()).rejects.toThrow('Baseline missing');
  expect(store.permissions.enableFolders).not.toHaveBeenCalled();
  expect(store.permissions.syncGroups).not.toHaveBeenCalled();
});
it('resumes a partial baseline repair without duplicating completed grants', async () => {
  const { permissions, store, acl } = fixture();
  const post = store.post.getMockImplementation();
  store.post.mockImplementation(path => path.includes('principalid=21') ? Promise.reject(new Error('Interrupted')) : post(path));
  await expect(permissions.applyScope('/project', null, 3)).rejects.toThrow('Interrupted');
  expect(acl('/project')).toContainEqual(row(20, 2));
  store.post.mockImplementation(post).mockClear();
  await permissions.applyScope('/project', null, 3);
  expect(store.post).toHaveBeenCalledTimes(2);
  expect(store.post.mock.calls.some(([path]) => path.includes('principalid=20'))).toBe(false);
});
it('copies inherited owner grants when preparing a new project scope', async () => {
  const { permissions, store, acl } = fixture();
  const get = store.get.getMockImplementation(), post = store.post.getMockImplementation();
  let unique = false;
  store.get.mockImplementation(path => path.endsWith('?$select=HasUniqueRoleAssignments') ? { HasUniqueRoleAssignments: unique } : get(path));
  store.post.mockImplementation(path => {
    if (path.endsWith('/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)')) { unique = true; return {}; }
    return post(path);
  });
  await permissions.applyScope('/new-project', null, 3);
  expect(store.post.mock.calls[0][0]).toBe('/new-project/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)');
  expect(acl('/new-project')).toContainEqual(row(99, 1073741829));
  expect(permissions.matchesBaseline(acl('/new-project'), await permissions.baselineGrants())).toBe(true);
});
it('rejects a custom manager role that lacks Manage Permissions', async () => {
  const { permissions, store } = fixture();
  const get = store.get.getMockImplementation();
  store.get.mockImplementation(path => path.includes('/roledefinitions?') ? { value: [roles[0], { ...roles[1], BasePermissions: { Low: '15' } }] } : get(path));
  await expect(permissions.applyScope('/project', null, 3)).rejects.toThrow('Contribute plus Manage Permissions');
  expect(store.post).not.toHaveBeenCalled();
});
