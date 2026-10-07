import { expect, it, vi } from 'vitest';
import { TrackerPermissions } from './trackerPermissions';

const acl = id => [{ PrincipalId: id, Member: { PrincipalType: 1 }, RoleDefinitionBindings: [{ Id: 3 }] },
  { PrincipalId: 50, Member: { PrincipalType: 8 }, RoleDefinitionBindings: [{ Id: 3 }] }];
function fixture(count = 132) {
  const permissions = new TrackerPermissions({ prefix: 'Modernization' });
  permissions.baselineGrants = vi.fn(async () => []);
  const rows = Array.from({ length: count }, (_, i) => ({ Id: i + 1, FileSystemObjectType: 0, FileDirRef: '/legacy', HasUniqueRoleAssignments: true }));
  const project = { spId: 500, projectKey: 'p', title: 'Project' };
  const state = new Map(rows.map(row => [row.Id, acl(99)]));
  permissions.contributionRole = vi.fn(async () => 3);
  permissions.folder = vi.fn(async key => ({ path: `/${key}/folder`, scope: `/${key}/folderScope` }));
  permissions.pages = vi.fn(async path => path.includes('ModernizationTasks') ? rows : []);
  let active = 0, peak = 0;
  permissions.assignments = vi.fn(async scope => {
    active++; peak = Math.max(active, peak);
    await Promise.resolve(); active--;
    return state.get(Number(scope.match(/items\((\d+)\)/)[1]));
  });
  let failing = true;
  permissions.applyScope = vi.fn(async scope => {
    if (!scope.includes('ModernizationTasks')) return;
    const id = Number(scope.match(/items\((\d+)\)/)[1]);
    if (id === 131 && failing) throw new Error('Host interrupted');
    state.set(id, [acl(99)[1]]); // no engineer assigned; preserve the manager group
  });
  return { permissions, project, rows, state, peak: () => peak, resume: () => { failing = false; permissions.applyScope.mockClear(); } };
}

it('rechecks live ACLs after item 131 interruption and skips all 130 matching tasks', async () => {
  const f = fixture();
  await expect(f.permissions.syncProject(f.project, [])).rejects.toThrow('Host interrupted');
  f.resume();
  const progress = vi.fn();
  await f.permissions.syncProject(f.project, [], progress);
  const writes = f.permissions.applyScope.mock.calls.filter(([scope]) => scope.includes('ModernizationTasks'));
  expect(writes.map(([scope]) => Number(scope.match(/items\((\d+)\)/)[1]))).toEqual([131, 132]);
  expect(f.peak()).toBe(4);
  expect(progress.mock.calls.at(-1)[0]).toContain('130 already correct, 2 updated');
  expect(progress.mock.calls.some(([message]) => message.endsWith('· item 1'))).toBe(false);
});

it('detects an externally changed permission on retry instead of trusting completion history', async () => {
  const f = fixture(4); f.resume();
  await f.permissions.syncProject(f.project, []);
  f.state.set(2, acl(99)); f.permissions.applyScope.mockClear();
  await f.permissions.syncProject(f.project, []);
  const writes = f.permissions.applyScope.mock.calls.filter(([scope]) => scope.includes('ModernizationTasks'));
  expect(writes).toHaveLength(1);
  expect(writes[0][0]).toContain('/items(2)');
});

it('does not mutate any task in a batch when an ACL read fails', async () => {
  const f = fixture(4);
  f.permissions.assignments.mockRejectedValueOnce(new Error('403 denied'));
  await expect(f.permissions.syncProject(f.project, [])).rejects.toThrow('403 denied');
  expect(f.permissions.applyScope.mock.calls.some(([scope]) => scope.includes('ModernizationTasks'))).toBe(false);
});

it('does not skip inherited legacy tasks merely because their current grants match', async () => {
  const f = fixture(1); f.rows[0].HasUniqueRoleAssignments = false;
  f.state.set(1, []);
  await f.permissions.syncProject(f.project, []);
  expect(f.permissions.assignments).not.toHaveBeenCalled();
  expect(f.permissions.applyScope).toHaveBeenCalledWith(expect.stringContaining('ModernizationTasks'), null, 3, []);
});

it('compares only the assigned engineer Contribute binding while preserving group and owner grants', () => {
  const p = new TrackerPermissions({});
  expect(p.matchesEngineer(acl(11), 11, 3)).toBe(true);
  expect(p.matchesEngineer([...acl(11), ...acl(12)], 11, 3)).toBe(false);
  expect(p.matchesEngineer(acl(12), 11, 3)).toBe(false);
  expect(p.matchesEngineer([acl(11)[1], { PrincipalId: 90, Member: { PrincipalType: 1 }, RoleDefinitionBindings: [{ Id: 99 }] }], null, 3)).toBe(true);
});

it.each([{}, { value: [{ PrincipalId: 1 }] }, { value: [{ PrincipalId: 1, Member: { PrincipalType: 1 }, RoleDefinitionBindings: {} }] }])('never considers unreadable ACL data a match: %j', async response => {
  const p = new TrackerPermissions({ get: async () => response });
  await expect(p.assignments('/scope')).rejects.toThrow();
});
