import { describe, expect, it, vi } from 'vitest';
import { SharePointStore } from './spStore';
import { TRACKER_GROUPS } from './trackerPermissions';

const person = { id: 'user-77', loginName: 'i:0#.f|membership|manager@example.com', title: 'Manager', role: 'Manager' };
const assignment = (id, roles, type = 1) => ({ PrincipalId: id, Member: { PrincipalType: type }, RoleDefinitionBindings: roles.map(Id => ({ Id })) });
function fixture(withGrants = true) {
  const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/metsoft', scopedAccess: true });
  const groups = Object.entries(TRACKER_GROUPS).map(([role, Title], index) => ({ Id: index + 1, Title, role, members: role === 'Manager' ? [{ Id: 77, LoginName: person.loginName }] : [] }));
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
  return { store, groups, rows, acl, failRead: value => { failedRead = value; }, failWrite: value => { failedWrite = value; } };
}

describe('targeted demotion cleanup', () => {
  it('demotes a manager with no direct engineer grants using only the elevated-group removal', async () => {
    const { store, groups, acl } = fixture(false), before = structuredClone([...acl]);
    await store.syncUserAccess(person, 'Viewer', [person]);
    expect(store.post.mock.calls).toHaveLength(1);
    expect(store.post.mock.calls.every(([path]) => path.includes('/sitegroups('))).toBe(true);
    expect([...acl]).toEqual(before);
    expect(groups.find(group => group.role === 'Manager').members).toEqual([]);
    expect(groups.every(group => group.members.length === 0)).toBe(true);
    expect(store.permissionProjects).not.toHaveBeenCalled();
  });
  it('removes only existing target-user Contribute grants, preserving other engineers, owners and groups', async () => {
    const { store, acl } = fixture();
    await store.syncUserAccess(person, 'Viewer', [person]);
    const writes = store.post.mock.calls.map(([path]) => path);
    expect(writes).toHaveLength(4); // 3 real grants + manager membership removal
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
    expect(store.post.mock.calls).toHaveLength(3);
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
