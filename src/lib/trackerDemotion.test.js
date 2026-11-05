import { describe, expect, it, vi } from 'vitest';
import { SharePointStore } from './spStore';
import { TRACKER_GROUPS } from './trackerPermissions';

const person = { id: 'user-77', loginName: 'i:0#.f|membership|manager@example.com', title: 'Manager', role: 'Manager' };
const assignment = (id, roles, type = 1) => ({ PrincipalId: id, Member: { PrincipalType: type }, RoleDefinitionBindings: roles.map(Id => ({ Id })) });
function fixture(withGrants = true) {
  const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/metsoft', scopedAccess: true });
  const groups = Object.entries(TRACKER_GROUPS).map(([role, Title], index) => ({ Id: index + 1, Title, role, members: role === 'Manager' ? [{ Id: 77, LoginName: person.loginName }] : [] }));
  const trackerUsers = { Id: 99, Title: 'Tracker Users', members: [] };
  groups.push(trackerUsers);
  const rows = {
    Projects: [{ Id: 1, HasUniqueRoleAssignments: false }, { Id: 2, HasUniqueRoleAssignments: true }],
    Tasks: [10, 11, 12, 13].map(Id => ({ Id, HasUniqueRoleAssignments: Id !== 12 })),
    Updates: [{ Id: 20, HasUniqueRoleAssignments: true }], Risks: [],
  };
  const root = suffix => `/_api/web/lists/getbytitle('Modernization${suffix}')`;
  const acl = new Map([
    [`${root('Projects')}/items(2)`, [assignment(77, [99, 2]), assignment(88, [3]), assignment(50, [3], 8)]],
    ...[10, 11, 13].map(id => [`${root('Tasks')}/items(${id})`, [assignment(77, withGrants ? [3, 2] : [2]), assignment(88, [3]), assignment(50, [3], 8)]]),
    [`${root('Updates')}/items(20)`, [assignment(88, [3])]],
  ]);
  let failedRead = '', failedWrite = '';
  store.get = vi.fn(async path => {
    if (failedRead && path.includes(failedRead)) throw new Error('Cannot read ACL');
    const named = groups.find(group => path.includes(`getbyname('${encodeURIComponent(group.Title)}')`));
    if (named) return { Id: named.Id, Title: named.Title };
    const groupMatch = path.match(/sitegroups\((\d+)\)\/users/);
    if (groupMatch) return { value: groups.find(group => group.Id === Number(groupMatch[1])).members };
    if (path.includes('/siteusers?')) return { value: [{ Id: 77, LoginName: person.loginName }] };
    if (path.includes('/roledefinitions')) return { value: [{ Id: 3, RoleTypeKind: 3 }] };
    const scope = [...acl.keys()].find(scope => path.startsWith(`${scope}/roleassignments`) || path.startsWith(`${scope}?`));
    if (scope) return path.includes('/roleassignments') ? { value: structuredClone(acl.get(scope)) } : { HasUniqueRoleAssignments: true };
    const suffix = Object.keys(rows).find(suffix => path.startsWith(`${root(suffix)}/items?`));
    if (suffix) return { value: rows[suffix] };
    throw new Error(`Unexpected GET: ${path}`);
  });
  store.post = vi.fn(async (path, options) => {
    if (failedWrite && path.includes(failedWrite)) throw new Error('Write interrupted');
    const scope = [...acl.keys()].find(scope => path.startsWith(`${scope}/roleassignments/removeroleassignment`));
    if (scope) {
      expect(path).toContain('principalid=77,roledefid=3');
      acl.set(scope, acl.get(scope).map(row => Number(row.PrincipalId) === 77 ? { ...row, RoleDefinitionBindings: row.RoleDefinitionBindings.filter(binding => binding.Id !== 3) } : row));
      return;
    }
    const group = groups.find(group => path.includes(`sitegroups(${group.Id})`));
    if (group && path.includes('removebyid')) group.members = [];
    else if (group && options?.body?.LoginName === person.loginName) group.members = [{ Id: 77, LoginName: person.loginName }];
    else throw new Error(`Unexpected mutation: ${path}`);
  });
  store.permissionProjects = vi.fn(() => { throw new Error('Role change must not rebuild projects'); });
  return { store, groups, trackerUsers, rows, acl, failRead: value => { failedRead = value; }, failWrite: value => { failedWrite = value; } };
}

describe('targeted demotion cleanup', () => {
  it.each([true, false])('verifies Tracker Users before removing Manager with an existing directory entry: %s', async hasEntry => {
    const { store, groups, trackerUsers, acl } = fixture(false), before = structuredClone([...acl]);
    await store.syncUserAccess(person, 'Viewer', hasEntry ? [person] : []);
    const manager = groups.find(group => group.role === 'Manager');
    expect(store.post.mock.calls.map(([path]) => path)).toEqual([
      '/_api/web/sitegroups(99)/users', `/_api/web/sitegroups(${manager.Id})/users/removebyid(77)`,
    ]);
    const baselineReads = store.get.mock.calls.flatMap(([path], i) => path.includes('sitegroups(99)/users') ? [store.get.mock.invocationCallOrder[i]] : []);
    expect(baselineReads).toHaveLength(2);
    expect(baselineReads[1]).toBeGreaterThan(store.post.mock.invocationCallOrder[0]);
    expect(baselineReads[1]).toBeLessThan(store.post.mock.invocationCallOrder[1]);
    expect([...acl]).toEqual(before);
    expect(groups.filter(group => group !== trackerUsers).every(group => group.members.length === 0)).toBe(true);
    expect(trackerUsers.members).toEqual([{ Id: 77, LoginName: person.loginName }]);
    expect(store.permissionProjects).not.toHaveBeenCalled();
  });
  it('repairs an already-demoted directory Viewer and skips a duplicate add on retry', async () => {
    const { store, groups, trackerUsers } = fixture(false);
    groups.find(group => group.role === 'Manager').members = [];
    const viewer = { ...person, role: 'Viewer', spId: 7 };
    store.directoryRows = vi.fn(async () => [viewer]);
    const read = store.get.getMockImplementation();
    store.get.mockImplementation(path => path.includes('/items(7)?') ? { Id: 7, RecordId: viewer.id, LoginKey: viewer.loginName, AppRole: 'Viewer' } : read(path));
    store.update = vi.fn();
    await store.saveUser(viewer);
    await store.saveUser(viewer);
    expect(store.post.mock.calls).toEqual([['/_api/web/sitegroups(99)/users', { body: { LoginName: person.loginName } }]]);
    expect(trackerUsers.members).toEqual([{ Id: 77, LoginName: person.loginName }]);
    expect(store.update).toHaveBeenCalledTimes(2);
    expect(store.get.mock.calls.some(([path]) => path.includes('HasUniqueRoleAssignments'))).toBe(false);
  });
  it('retains the manager role and directory record when Tracker Users enrollment is unconfirmed', async () => {
    const { store, groups } = fixture(false);
    store.post.mockResolvedValue(undefined); // SharePoint accepts the write but never confirms membership.
    store.directoryRows = vi.fn(async () => [{ ...person, spId: 7 }]);
    store.update = vi.fn(); store.create = vi.fn();
    await expect(store.saveUser({ ...person, role: 'Viewer' })).rejects.toThrow('SharePoint did not confirm group membership');
    expect(store.post.mock.calls).toEqual([['/_api/web/sitegroups(99)/users', { body: { LoginName: person.loginName } }]]);
    expect(groups.find(group => group.role === 'Manager').members).toHaveLength(1);
    expect(store.update).not.toHaveBeenCalled(); expect(store.create).not.toHaveBeenCalled();
  });
  it('removes only existing target-user Contribute grants, preserving other engineers, owners and groups', async () => {
    const { store, acl } = fixture();
    await store.syncUserAccess(person, 'Viewer', [person]);
    const writes = store.post.mock.calls.map(([path]) => path);
    expect(writes).toHaveLength(5); // 3 real grants, basic enrollment, manager membership removal
    expect(writes.slice(0, 3).every(path => path.includes('removeroleassignment(principalid=77,roledefid=3)'))).toBe(true);
    expect(writes.join(' ')).not.toMatch(/breakroleinheritance|addroleassignment|AddValidate|resetroleinheritance/);
    for (const entries of acl.values()) {
      expect(entries.some(row => row.PrincipalId === 77 && row.RoleDefinitionBindings.some(role => role.Id === 3))).toBe(false);
      expect(entries.find(row => row.PrincipalId === 88).RoleDefinitionBindings).toEqual([{ Id: 3 }]);
    }
    const owner = [...acl.values()][0].find(row => row.PrincipalId === 77);
    expect(owner.RoleDefinitionBindings).toEqual([{ Id: 99 }, { Id: 2 }]);
  });
  it('retries interrupted revocation without repeating completed writes or changing the saved role early', async () => {
    const { store, failWrite, groups } = fixture();
    store.directoryRows = vi.fn(async () => [{ ...person, spId: 7 }]);
    store.update = vi.fn(); store.create = vi.fn();
    failWrite('/items(11)/roleassignments/');
    await expect(store.saveUser({ ...person, role: 'Viewer' })).rejects.toThrow('Write interrupted');
    expect(store.update).not.toHaveBeenCalled(); expect(store.create).not.toHaveBeenCalled();
    expect(groups.find(group => group.role === 'Manager').members).toHaveLength(1);
    failWrite(''); store.post.mockClear();
    await store.syncUserAccess(person, 'Viewer', [person]);
    expect(store.post.mock.calls).toHaveLength(4);
    expect(store.post.mock.calls.some(([path]) => path.includes('/items(10)'))).toBe(false);
  });
  it('does not write anything if any ACL preflight or scope inventory is incomplete', async () => {
    const { store, rows, failRead } = fixture();
    failRead('/items(20)/roleassignments');
    await expect(store.syncUserAccess(person, 'Viewer', [person])).rejects.toThrow('Cannot read ACL');
    expect(store.post).not.toHaveBeenCalled();
    failRead(''); rows.Risks.push({ Id: 99 });
    await expect(store.syncUserAccess(person, 'Viewer', [person])).rejects.toThrow('Cannot verify unique');
    expect(store.post).not.toHaveBeenCalled();
  });
  it('finds stale direct grants when a previous attempt already removed every tracker membership', async () => {
    const { store, groups } = fixture();
    groups.forEach(group => { group.members = []; });
    await store.syncUserAccess(person, null, [person]);
    expect(store.get.mock.calls.some(([path]) => path.includes('/siteusers?'))).toBe(true);
    expect(store.post.mock.calls).toHaveLength(3);
    expect(store.post.mock.calls.every(([path]) => path.includes('removeroleassignment'))).toBe(true);
  });
});
