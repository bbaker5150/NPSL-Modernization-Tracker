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
    store.directoryRows.mockResolvedValueOnce([]).mockResolvedValueOnce([]).mockResolvedValueOnce([newEntry]);
    await store.saveProject(project, account.LoginName);
    expect(store.create.mock.calls.map(([key]) => key)).toEqual(['projects']);
    expect(user()).toMatchObject({ RecordId: newEntry.id, AppRole: 'Project Engineer' });
    expect(await store.directoryRows.mock.results.at(-1).value).toEqual([expect.objectContaining({ id: newEntry.id, role: 'Project Engineer' })]);
  });
  it('preserves a concurrent manager promotion observed on the directory recheck', async () => {
    const { store, groups, acl } = fixture();
    const promoted = { spId: 7, id: 'concurrent-manager', loginName: account.LoginName, email: account.Email, role: 'Manager' };
    store.directoryRows.mockResolvedValueOnce([]).mockResolvedValueOnce([]).mockImplementationOnce(async () => {
      groups[2].members.push(account);
      return [promoted];
    }).mockResolvedValue([promoted]);
    await store.saveProject(project, account.LoginName);
    expect(store.update).not.toHaveBeenCalled();
    expect(store.create.mock.calls.map(([key]) => key)).toEqual(['projects']);
    expect(groups[2].members).toEqual([account]);
    expect(await store.directoryRows.mock.results.at(-1).value).toEqual([promoted]);
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

describe('manager assignment of project teams', () => {
  const second = { ...account, Id: 14, Title: 'Second Engineer', LoginName: 'claims|second@example.test', Email: 'second@example.test' };
  const retained = { ...account, Id: 15, Title: 'Retained Engineer', LoginName: 'claims|retained@example.test', Email: 'retained@example.test' };
  const accounts = [account, second, retained];
  const engineer = item => ({ name: item.Title, email: item.Email, loginName: item.LoginName });
  function teamFixture({ roles = [], projects = [] } = {}) {
    const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/metsoft', scopedAccess: true });
    const groups = [TRACKER_USERS_GROUP, TRACKER_GROUPS['Project Engineer'], TRACKER_GROUPS.Manager].map((Title, i) => ({ Id: i + 1, Title, members: [] }));
    const users = new Map();
    let userId = 100;
    for (const [person, role] of roles) {
      const Id = ++userId;
      users.set(Id, { Id, RecordId: `saved-${person.Id}`, Title: person.Title, LoginKey: person.LoginName, Email: person.Email, AppRole: role });
      groups[0].members.push(person);
      if (role === 'Project Engineer') groups[1].members.push(person);
      if (role === 'Manager') groups[2].members.push(person);
    }
    const savedProjects = new Map();
    const timeline = [];
    const directory = () => [...users.values()].map(row => ({ id: row.RecordId, spId: row.Id, title: row.Title, loginName: row.LoginKey, email: row.Email, role: row.AppRole }));
    store.currentUser = vi.fn(async () => ({ loginName: 'actor' }));
    store.permissions.currentRole = vi.fn(async () => ({ role: 'Manager' }));
    store.directoryRows = vi.fn(async () => directory());
    store.permissionProjects = vi.fn(async () => [...projects, ...[...savedProjects].map(([spId, row]) => ({ spId, id: row.RecordId, projectKey: row.ProjectKey, ownerKey: row.OwnerKey, engineers: JSON.parse(row.EngineersJson || '[]') }))]);
    store.permissions.syncProject = vi.fn();
    store.get = vi.fn(async path => {
      const user = path.match(/ModernizationUsers'\)\/items\((\d+)\)/);
      if (user) return users.get(Number(user[1]));
      if (path.includes('/sitegroups/getbyname')) return groups.find(group => decodeURIComponent(path).includes(`('${group.Title}')`));
      const membership = path.match(/sitegroups\((\d+)\)\/users/);
      if (membership) return { value: groups[Number(membership[1]) - 1].members };
      if (path.includes('/getusereffectivepermissions')) return { GetUserEffectivePermissions: { Low: '0' } };
      throw new Error(`Unexpected GET ${path}`);
    });
    store.post = vi.fn(async (path, options) => {
      if (path.includes('/ensureuser?')) {
        timeline.push(`resolve:${options.body.logonName}`);
        const person = accounts.find(item => item.LoginName.toLowerCase() === options.body.logonName.toLowerCase());
        if (!person) throw new Error('Unknown account');
        return person;
      }
      const membership = path.match(/sitegroups\((\d+)\)\/users$/);
      if (membership) {
        timeline.push(`enroll:${options.body.LoginName}`);
        const person = accounts.find(item => item.LoginName === options.body.LoginName);
        if (!person) throw new Error('Unknown group member');
        groups[Number(membership[1]) - 1].members.push(person);
        return;
      }
      throw new Error(`Unexpected POST ${path}`);
    });
    store.update = vi.fn(async (key, id, fields) => {
      if (key === 'users') users.set(id, { Id: id, ...fields });
      else if (key === 'projects') savedProjects.set(id, fields);
      else throw new Error(`Unexpected update ${key}`);
    });
    store.create = vi.fn(async (key, fields) => {
      if (key === 'users') { const Id = ++userId; users.set(Id, { Id, ...fields }); return Id; }
      if (key === 'projects') { savedProjects.set(8, fields); return 8; }
      throw new Error(`Unexpected create ${key}`);
    });
    return { store, groups, users, directory, timeline, saved: () => savedProjects.get(8) };
  }

  it('enrolls every selected account, writes canonical team JSON, and mirrors the first verified engineer', async () => {
    const { store, groups, saved, directory, timeline } = teamFixture();
    const result = await store.saveProject(project, [account.LoginName, second.LoginName]);
    expect(result).toMatchObject({ spId: 8, ownerName: account.Title, ownerEmail: account.Email, ownerKey: account.LoginName, engineers: [engineer(account), engineer(second)] });
    expect(saved()).toMatchObject({ EngineersJson: JSON.stringify([engineer(account), engineer(second)]), OwnerName: account.Title, OwnerEmail: account.Email, OwnerKey: account.LoginName });
    expect(directory().map(row => [row.loginName, row.role])).toEqual([[account.LoginName, 'Project Engineer'], [second.LoginName, 'Project Engineer']]);
    expect(groups.map(group => group.members.map(row => row.LoginName))).toEqual([[account.LoginName, second.LoginName], [account.LoginName, second.LoginName], []]);
    expect(timeline.slice(0, 2)).toEqual([`resolve:${account.LoginName}`, `resolve:${second.LoginName}`]);
    expect(store.permissions.syncProject).toHaveBeenCalledExactlyOnceWith(result, directory());
  });

  it('normalizes duplicate logins before resolving or enrolling accounts', async () => {
    const { store, groups, saved } = teamFixture();
    await store.saveProject(project, [account.LoginName, ` ${account.LoginName.toUpperCase()} `, second.LoginName, second.LoginName]);
    expect(store.post.mock.calls.filter(([path]) => path.includes('/ensureuser?'))).toHaveLength(2);
    expect(JSON.parse(saved().EngineersJson)).toEqual([engineer(account), engineer(second)]);
    expect(groups[1].members).toEqual([account, second]);
    expect(store.create.mock.calls.map(([key]) => key)).toEqual(['users', 'users', 'projects']);
  });

  it('clears a team explicitly without reusing stale legacy owner fields or demoting its members', async () => {
    const { store, groups, saved, directory } = teamFixture({ roles: [[account, 'Project Engineer'], [second, 'Project Engineer']] });
    const result = await store.saveProject({ ...project, engineers: [engineer(account), engineer(second)], ownerKey: account.LoginName }, []);
    expect(saved()).toMatchObject({ EngineersJson: '[]', OwnerName: 'Unassigned', OwnerEmail: '', OwnerKey: '' });
    expect(result.engineers).toEqual([]);
    expect(store.post).not.toHaveBeenCalled();
    expect(groups[1].members).toEqual([account, second]);
    expect(directory().every(row => row.role === 'Project Engineer')).toBe(true);
    expect(store.permissions.syncProject).toHaveBeenCalledExactlyOnceWith(result, directory());
  });

  it('restores returning engineers on other team projects using the complete directory so retained teammates keep access', async () => {
    const other = { id: 'existing-team', spId: 20, projectKey: 'other', engineers: [engineer(retained), engineer(account)] };
    const unrelated = { id: 'unrelated', spId: 21, projectKey: 'elsewhere', engineers: [engineer(retained)] };
    const { store, directory } = teamFixture({ roles: [[account, 'Viewer'], [retained, 'Project Engineer']], projects: [other, unrelated, { ...other, id: 'archived', spId: 22, archived: true }] });
    await store.saveProject(project, [account.LoginName, second.LoginName]);
    expect(store.permissions.syncProject.mock.calls.map(([row]) => row.id)).toEqual([project.id, other.id]);
    const completeDirectory = directory();
    expect(completeDirectory).toHaveLength(3);
    for (const [, users] of store.permissions.syncProject.mock.calls) {
      expect(users).toEqual(completeDirectory);
      expect(users).toEqual(expect.arrayContaining([expect.objectContaining({ loginName: retained.LoginName, role: 'Project Engineer' })]));
    }
  });

  it('does not re-promote a retained Viewer when a different engineer is replaced', async () => {
    const existing = { ...project, spId: 8, engineers: [engineer(account), engineer(second)] };
    const { store, groups, directory, saved } = teamFixture({ roles: [[account, 'Project Engineer'], [second, 'Viewer']], projects: [existing] });
    const result = await store.saveProject(existing, [second.LoginName, retained.LoginName]);
    expect(result.engineers).toEqual([engineer(second), engineer(retained)]);
    expect(saved()).toMatchObject({ OwnerKey: second.LoginName, EngineersJson: JSON.stringify([engineer(second), engineer(retained)]) });
    expect(directory().find(row => row.loginName === second.LoginName).role).toBe('Viewer');
    expect(directory().find(row => row.loginName === retained.LoginName).role).toBe('Project Engineer');
    expect(groups[1].members.map(row => row.LoginName)).toEqual([account.LoginName, retained.LoginName]);
    expect(store.post.mock.calls.filter(([path]) => path.includes('/sitegroups(')).every(([, options]) => options.body.LoginName === retained.LoginName)).toBe(true);
    expect(store.update.mock.calls.filter(([key]) => key === 'users')).toEqual([]);
    expect(store.permissions.syncProject).toHaveBeenCalledExactlyOnceWith(result, directory());
  });

  it('preserves a retained engineer demoted to Viewer after the team preflight read', async () => {
    const existing = { ...project, spId: 8, engineers: [engineer(account), engineer(second)] };
    const { store, groups, users, directory, saved } = teamFixture({ roles: [[account, 'Project Engineer'], [second, 'Project Engineer']], projects: [existing] });
    let reads = 0;
    store.directoryRows.mockImplementation(async () => {
      if (++reads === 2) {
        // Another manager completes the demotion after the initial team snapshot,
        // before this retained account's enrollment starts.
        const row = [...users.values()].find(item => item.LoginKey === second.LoginName);
        users.set(row.Id, { ...row, AppRole: 'Viewer' });
        groups[1].members = groups[1].members.filter(item => item.LoginName !== second.LoginName);
      }
      return directory();
    });
    const result = await store.saveProject(existing, [second.LoginName, retained.LoginName]);
    expect(directory().find(row => row.loginName === second.LoginName).role).toBe('Viewer');
    expect(groups[1].members.map(row => row.LoginName)).toEqual([account.LoginName, retained.LoginName]);
    expect(store.post.mock.calls.filter(([path]) => path.includes('/sitegroups(')).every(([, options]) => options.body.LoginName === retained.LoginName)).toBe(true);
    expect(store.update.mock.calls.filter(([key]) => key === 'users')).toEqual([]);
    expect(saved().EngineersJson).toBe(JSON.stringify([engineer(second), engineer(retained)]));
    expect(store.permissions.syncProject).toHaveBeenCalledExactlyOnceWith(result, directory());
  });

  it('stops the project save when a retained engineer is demoted during live role verification', async () => {
    const existing = { ...project, spId: 8, engineers: [engineer(account), engineer(second)] };
    const { store, groups, users, directory } = teamFixture({ roles: [[account, 'Project Engineer'], [second, 'Project Engineer']], projects: [existing] });
    const verify = store.permissions.ensureAssignmentRole.bind(store.permissions);
    store.permissions.ensureAssignmentRole = vi.fn(async (...args) => {
      const role = await verify(...args);
      const row = [...users.values()].find(item => item.LoginKey === second.LoginName);
      users.set(row.Id, { ...row, AppRole: 'Viewer' });
      groups[1].members = groups[1].members.filter(item => item.LoginName !== second.LoginName);
      return role;
    });
    await expect(store.saveProject(existing, [second.LoginName, retained.LoginName])).rejects.toThrow('changed to Viewer during assignment');
    expect(directory().find(row => row.loginName === second.LoginName).role).toBe('Viewer');
    expect(store.create).not.toHaveBeenCalled(); expect(store.update).not.toHaveBeenCalled();
    expect(store.permissions.syncProject).not.toHaveBeenCalled();
  });

  it('rejects an unresolved second identity before any first-engineer enrollment or project write', async () => {
    const { store, groups, timeline } = teamFixture();
    await expect(store.saveProject(project, [account.LoginName, 'claims|missing@example.test'])).rejects.toThrow('Unknown account');
    expect(timeline.every(item => item.startsWith('resolve:'))).toBe(true);
    expect(groups.every(group => !group.members.length)).toBe(true);
    expect(store.create).not.toHaveBeenCalled(); expect(store.update).not.toHaveBeenCalled();
    expect(store.permissions.syncProject).not.toHaveBeenCalled();
  });

  it('checks every directory identity before enrolling the first engineer', async () => {
    const { store, groups, timeline } = teamFixture();
    store.directoryRows.mockResolvedValue([{ id: 'conflict', loginName: 'claims|different-account', email: second.Email, role: 'Viewer' }]);
    await expect(store.saveProject(project, [account.LoginName, second.LoginName])).rejects.toThrow('conflicting Tracker directory identities');
    expect(timeline.every(item => item.startsWith('resolve:'))).toBe(true);
    expect(groups.every(group => !group.members.length)).toBe(true);
    expect(store.create).not.toHaveBeenCalled(); expect(store.update).not.toHaveBeenCalled();
  });

  it('does not resolve accounts or write anything when the new team column has not been provisioned', async () => {
    const { store } = teamFixture();
    store.schemaReadiness = { checks: [{ key: 'projects', missingFields: ['EngineersJson'] }] };
    await expect(store.saveProject(project, [account.LoginName, second.LoginName])).rejects.toThrow('site owner');
    expect(store.post).not.toHaveBeenCalled(); expect(store.create).not.toHaveBeenCalled(); expect(store.update).not.toHaveBeenCalled();
    expect(store.permissions.syncProject).not.toHaveBeenCalled();
  });

  it('keeps legacy single-engineer writes compatible before the team column is added', async () => {
    const { store, saved } = teamFixture();
    store.schemaReadiness = { checks: [{ key: 'projects', missingFields: ['EngineersJson'] }] };
    await store.saveProject(project, account.LoginName);
    expect(saved()).toMatchObject({ OwnerKey: account.LoginName, OwnerName: account.Title });
    expect(saved()).not.toHaveProperty('EngineersJson');
  });
});
