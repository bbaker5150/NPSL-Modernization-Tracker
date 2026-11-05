import { describe, expect, it, vi } from 'vitest';
import { SharePointStore } from './spStore';
import { TRACKER_GROUPS, TRACKER_USERS_GROUP } from './trackerGroups';

const account = { Id: 13, Title: 'New Engineer', LoginName: 'claims|engineer@example.test', Email: 'engineer@example.test', PrincipalType: 1, IsSiteAdmin: false };
const project = { id: 'record-p', projectKey: 'p', title: 'P', ownerName: 'Untrusted name', ownerKey: 'other', ownerEmail: 'wrong@example.test' };
const binding = id => ({ PrincipalId: id, Member: { PrincipalType: 1 }, RoleDefinitionBindings: [{ Id: 3 }] });

function fixture({ savedRole, liveManager = false, sitePermissions = '0', actorRole = 'Manager', resolved = account } = {}) {
  const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/metsoft', scopedAccess: true });
  const groups = [TRACKER_USERS_GROUP, TRACKER_GROUPS['Project Engineer'], TRACKER_GROUPS.Manager].map((Title, i) => ({
    Id: i + 1, Title, members: Title === TRACKER_GROUPS.Manager && liveManager ? [account] : [],
  }));
  let user = savedRole && { Id: 7, RecordId: 'old-record', LoginKey: account.LoginName, Email: account.Email, AppRole: savedRole };
  let saved;
  const acl = new Map();
  function assignments(scope) {
    if (!acl.has(scope)) acl.set(scope, [binding(99)]); // former engineer
    return acl.get(scope);
  }
  store.currentUser = vi.fn(async () => ({ loginName: 'actor' }));
  store.permissions.currentRole = vi.fn(async () => ({ role: actorRole }));
  store.directoryRows = vi.fn(async () => user ? [{ id: user.RecordId, loginName: user.LoginKey, role: user.AppRole, spId: user.Id, email: user.Email }] : []);
  store.permissionProjects = vi.fn(async () => saved ? [{ id: saved.RecordId, spId: 8 }] : []);
  store.permissions.baselineGrants = vi.fn(async () => []);
  store.permissions.folder = vi.fn(async key => ({ scope: '/project-folder/' + key, path: '/folders/' + key }));
  store.get = vi.fn(async path => {
    if (path.includes("ModernizationUsers')/items(7)")) return user;
    if (path.includes('/sitegroups/getbyname')) return groups.find(group => decodeURIComponent(path).includes(`('${group.Title}')`));
    const membership = path.match(/sitegroups\((\d+)\)\/users/);
    if (membership) return { value: groups[Number(membership[1]) - 1].members };
    if (path.includes('/getuserbyid(13)?$select=Id,LoginName,Title,Email,PrincipalType,IsSiteAdmin')) return account;
    if (path.includes('/getusereffectivepermissions')) {
      if (sitePermissions instanceof Error) throw sitePermissions;
      return { GetUserEffectivePermissions: { Low: sitePermissions } };
    }
    if (path.includes('/roledefinitions')) return { value: [{ Id: 3, RoleTypeKind: 3 }] };
    if (path.includes('/items?')) return { value: [] };
    if (path.includes('?$select=HasUniqueRoleAssignments')) return { HasUniqueRoleAssignments: true };
    if (path.includes('/roleassignments?')) return { value: assignments(path.split('/roleassignments?')[0]) };
    throw new Error(`Unexpected GET ${path}`);
  });
  store.post = vi.fn(async (path, options) => {
    if (path.endsWith('/ensureuser?$select=Id,LoginName,Title,Email,PrincipalType,IsSiteAdmin')) return resolved;
    const membership = path.match(/sitegroups\((\d+)\)\/users$/);
    if (membership && options.body.LoginName === account.LoginName) {
      groups[Number(membership[1]) - 1].members.push(account);
      return;
    }
    const change = path.match(/^(.*)\/roleassignments\/(add|remove)roleassignment\(principalid=(\d+),roledefid=3\)$/);
    if (change) {
      const [, scope, action, id] = change;
      if (action === 'add') assignments(scope).push(binding(Number(id)));
      else acl.set(scope, assignments(scope).filter(entry => entry.PrincipalId !== Number(id)));
      return;
    }
    throw new Error(`Unexpected POST ${path}`);
  });
  store.update = vi.fn(async (key, id, fields) => {
    if (key === 'users') user = { Id: id, ...fields };
    if (key === 'projects') saved = fields;
  });
  store.create = vi.fn(async (key, fields) => {
    if (key === 'users') { user = { Id: 7, ...fields }; return 7; }
    if (key === 'projects') { saved = fields; return 8; }
    throw new Error(`Unexpected create ${key}`);
  });
  return { store, groups, acl, user: () => user, saved: () => saved };
}

describe('explicit manager assignment of a project engineer', () => {
  it.each([undefined, 'Viewer'])('enrolls a resolved %s, replaces the former project grants, and retries without duplicates', async savedRole => {
    const { store, groups, acl, user, saved } = fixture({ savedRole });
    const sync = store.permissions.syncProject.bind(store.permissions);
    store.permissions.syncProject = vi.fn().mockRejectedValueOnce(new Error('ACL interrupted')).mockImplementation(sync);
    await expect(store.saveProject(project, account.LoginName)).rejects.toThrow('Project saved, but its access setup is incomplete');
    expect(saved()).toMatchObject({ OwnerKey: account.LoginName, OwnerName: account.Title, OwnerEmail: account.Email });
    expect(user()).toMatchObject({ RecordId: savedRole ? 'old-record' : 'sp-user-13', LoginKey: account.LoginName, AppRole: 'Project Engineer' });
    expect((await store.saveProject(project, account.LoginName)).spId).toBe(8);
    expect(groups.map(group => group.members.length)).toEqual([1, 1, 0]);
    expect(store.post.mock.calls.filter(([path]) => path.includes('/sitegroups('))).toHaveLength(2);
    expect(store.create.mock.calls.map(([key]) => key)).toEqual(savedRole ? ['projects'] : ['users', 'projects']);
    expect(acl.size).toBe(4); // this project plus its tasks, updates and risks folders
    for (const grants of acl.values()) expect(grants).toEqual([binding(account.Id)]);
    expect(store.post.mock.calls.some(([path]) => path.includes('ShareObject'))).toBe(false);
  });

  it.each([
    ['Tracker Manager', { liveManager: true, savedRole: 'Viewer' }],
    ['saved Manager', { liveManager: true, savedRole: 'Manager' }],
    ['site owner', { sitePermissions: '33554432' }],
    ['site collection administrator', { resolved: { ...account, IsSiteAdmin: true } }],
    ['saved Manager whose site permissions the caller cannot inspect', { savedRole: 'Manager', sitePermissions: Object.assign(new Error('Forbidden'), { status: 403 }) }],
  ])('preserves a %s without granting manager or direct engineer permissions', async (_description, options) => {
    const { store, groups, acl, user } = fixture(options);
    await store.saveProject(project, account.LoginName);
    expect(user().AppRole).toBe('Manager');
    expect(groups.map(group => group.members.length)).toEqual([1, 0, options.liveManager ? 1 : 0]);
    expect(store.post.mock.calls.filter(([path]) => path.includes('/sitegroups(')).map(([path]) => path)).toEqual(['/_api/web/sitegroups(1)/users']);
    expect([...acl.values()]).toEqual([[], [], [], []]);
  });

  it.each([
    ['conflicting Manager role', { savedRole: 'Manager' }],
    ['unverified site permissions', { sitePermissions: null }],
    ['unresolved target', { resolved: { ...account, LoginName: 'claims|other', Email: 'other@example.test' } }],
    ['non-person principal', { resolved: { ...account, PrincipalType: 8 } }],
    ['unauthorized actor', { actorRole: 'Project Engineer' }],
  ])('stops before enrollment and project save for %s', async (_description, options) => {
    const { store } = fixture(options);
    await expect(store.saveProject(project, account.LoginName)).rejects.toThrow();
    expect(store.post.mock.calls.every(([path]) => path.includes('/ensureuser?'))).toBe(true);
    expect(store.update).not.toHaveBeenCalled();
    expect(store.create).not.toHaveBeenCalled();
  });
  it('lets a delegated manager enroll a Viewer with group ownership but no site-wide permission inspection', async () => {
    const { store, user, groups } = fixture({ savedRole: 'Viewer', sitePermissions: Object.assign(new Error('Forbidden'), { status: 403 }), resolved: { ...account, PrincipalType: undefined } });
    await store.saveProject(project, account.LoginName);
    expect(store.get.mock.calls.some(([path]) => path.includes('/getuserbyid(13)?'))).toBe(true);
    expect(user().AppRole).toBe('Project Engineer');
    expect(groups.map(group => group.members.length)).toEqual([1, 1, 0]);
  });
  it('blocks project persistence when enrollment cannot be confirmed and reuses Tracker Users on retry', async () => {
    const { store, groups } = fixture();
    const post = store.post.getMockImplementation();
    store.post.mockImplementation((path, options) => path === '/_api/web/sitegroups(2)/users' ? {} : post(path, options));
    await expect(store.saveProject(project, account.LoginName)).rejects.toThrow('did not verify Tracker Project Engineers');
    expect(groups.map(group => group.members.length)).toEqual([1, 0, 0]);
    expect(store.create).not.toHaveBeenCalled();
    store.post.mockImplementation(post);
    await store.saveProject(project, account.LoginName);
    expect(store.post.mock.calls.filter(([path]) => path === '/_api/web/sitegroups(1)/users')).toHaveLength(1);
    expect(store.create.mock.calls.map(([key]) => key)).toEqual(['users', 'projects']);
  });
  it('refuses to overwrite a directory row that shares a mailbox but has another claims identity', async () => {
    const { store } = fixture();
    store.directoryRows = vi.fn(async () => [{ id: 'other', spId: 19, email: account.Email, loginName: 'claims|different-account', role: 'Manager' }]);
    await expect(store.saveProject(project, account.LoginName)).rejects.toThrow('conflicting Tracker directory identities');
    expect(store.post.mock.calls.every(([path]) => path.includes('/ensureuser?'))).toBe(true);
    expect(store.update).not.toHaveBeenCalled(); expect(store.create).not.toHaveBeenCalled();
  });
  it('leaves a previously demoted Viewer alone on an unrelated save', async () => {
    const { store, groups, user } = fixture({ savedRole: 'Viewer' });
    await store.saveProject({ ...project, ownerName: account.Title, ownerKey: account.LoginName, ownerEmail: account.Email });
    expect(user().AppRole).toBe('Viewer');
    expect(groups.every(group => !group.members.length)).toBe(true);
    expect(store.post.mock.calls.some(([path]) => path.includes('ensureuser') || path.includes('/sitegroups('))).toBe(false);
  });
  it('keeps the verified sign-in email when the directory omits the Email property', async () => {
    const { store, saved, user } = fixture({ resolved: { ...account, Email: '' } });
    await store.saveProject(project, account.LoginName);
    expect(saved().OwnerEmail).toBe(account.Email);
    expect(user().Email).toBe(account.Email);
  });
  it('shares enrollment when two projects in the same session select the same new engineer', async () => {
    const { store, groups } = fixture();
    store.permissionProjects.mockResolvedValue([]);
    store.permissions.syncProject = vi.fn();
    await Promise.all([
      store.saveProject(project, account.LoginName),
      store.saveProject({ ...project, id: 'another', projectKey: 'other' }, account.LoginName),
    ]);
    expect(groups.map(group => group.members.length)).toEqual([1, 1, 0]);
    expect(store.create.mock.calls.map(([key]) => key)).toEqual(['users', 'projects', 'projects']);
  });
  it('reuses an entry added by another manager while SharePoint is confirming membership', async () => {
    const { store, user } = fixture();
    const newEntry = { spId: 7, id: 'concurrent-invite', loginName: account.LoginName, email: account.Email, role: 'Viewer' };
    store.directoryRows.mockResolvedValueOnce([]).mockResolvedValueOnce([newEntry]);
    await store.saveProject(project, account.LoginName);
    expect(store.create.mock.calls.map(([key]) => key)).toEqual(['projects']);
    expect(user()).toMatchObject({ RecordId: newEntry.id, AppRole: 'Project Engineer' });
  });
  it('preserves a concurrent manager promotion observed on the directory recheck', async () => {
    const { store, groups, acl } = fixture();
    store.directoryRows.mockResolvedValueOnce([]).mockImplementationOnce(async () => {
      groups[2].members.push(account);
      return [{ spId: 7, id: 'concurrent-manager', loginName: account.LoginName, email: account.Email, role: 'Manager' }];
    });
    await store.saveProject(project, account.LoginName);
    expect(store.update).not.toHaveBeenCalled();
    expect(store.create.mock.calls.map(([key]) => key)).toEqual(['projects']);
    expect(groups[2].members).toEqual([account]);
    expect([...acl.values()]).toEqual([[], [], [], []]);
  });
  it('restores editing on all active projects assigned to a returning engineer, including after an interruption', async () => {
    const { store, saved } = fixture({ savedRole: 'Viewer' });
    const owned = { id: 'prior', spId: 9, ownerKey: account.LoginName };
    const unrelated = { id: 'elsewhere', spId: 10, ownerKey: 'claims|someone-else', ownerEmail: account.Email };
    store.permissionProjects.mockImplementation(async () => [
      ...(saved() ? [{ id: project.id, spId: 8, ownerKey: account.LoginName }] : []),
      owned, unrelated, { ...owned, id: 'archived', spId: 11, archived: true },
    ]);
    const attempted = [];
    store.permissions.syncProject = vi.fn(async row => {
      attempted.push(row.id);
      if (row.id === owned.id && attempted.length === 2) throw new Error('interrupted');
    });
    await expect(store.saveProject(project, account.LoginName)).rejects.toThrow('Project saved, but its access setup is incomplete');
    await store.saveProject(project, account.LoginName);
    expect(attempted).toEqual([project.id, owned.id, project.id, owned.id]);
  });
});
