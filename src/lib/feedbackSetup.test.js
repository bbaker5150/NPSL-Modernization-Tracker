import { describe, it, expect, vi } from 'vitest';
import { feedbackSetup, SUBMIT_PERMISSIONS } from './feedbackSetup';
const missing = () => Object.assign(new Error('Missing'), { status: 404 });
export function setupFixture() {
  const state = { list: null, fields: [], roles: [{ Id: 1, RoleTypeKind: 5, BasePermissions: { Low: '4294967295', High: '2147483647' } }], assignments: [], failField: false };
  const raw = { webUrl: 'https://example.test/sites/tracker', scopedAccess: true, currentUser: async () => ({ id: 8, isSiteAdmin: true }) };
  raw.get = vi.fn(async path => {
    if (path.includes('/EffectiveBasePermissions')) return { Low: '4294967295' };
    if (path.includes('/associatedownergroup')) return { Id: 10, Title: 'Site Owners' };
    if (path.includes('/sitegroups?')) return { value: [{ Id: 10, Title: 'Site Owners' }, { Id: 20, Title: 'Metrology App User' }, { Id: 21, Title: 'Tracker Project Engineers' }, { Id: 22, Title: 'Tracker Managers' }] };
    if (path.includes('/roledefinitions/getbyname')) return state.roles.find(role => role.Name);
    if (path.includes('/roledefinitions?')) return { value: structuredClone(state.roles) };
    if (path.includes('/roleassignments?')) return { value: state.assignments.map(assignment => ({ ...assignment, RoleDefinitionBindings: assignment.RoleDefinitionBindings.map(role => path.includes('BasePermissions') ? state.roles.find(entry => entry.Id === role.Id) : role) })) };
    if (path.includes('/fields?')) return { value: structuredClone(state.fields) };
    if (!state.list) throw missing();
    return structuredClone(state.list);
  });
  raw.post = vi.fn(async (path, options = {}) => {
    if (path === '/_api/web/lists') state.list = { Id: 'list', ...options.body, ItemCount: 0, HasUniqueRoleAssignments: false };
    else if (path.includes('/breakroleinheritance')) { state.list.HasUniqueRoleAssignments = true; if (!state.assignments.length) state.assignments = [{ PrincipalId: 8, RoleDefinitionBindings: [{ Id: 1 }] }]; }
    else if (path.includes('/fields/createfieldasxml')) { if (state.failField) throw new Error('Field creation failed'); state.fields.push({ InternalName: options.body.parameters.SchemaXml.match(/ Name="([^"]+)"/)[1] }); }
    else if (path === '/_api/web/roledefinitions') state.roles.push({ Id: 2, ...options.body });
    else if (path.includes('/addroleassignment(')) { const [, principal, role] = path.match(/principalid=(\d+),roledefid=(\d+)/); state.assignments.push({ PrincipalId: Number(principal), RoleDefinitionBindings: [{ Id: Number(role) }] }); }
    else if (options.headers?.['X-HTTP-Method'] === 'DELETE') { const principal = Number(path.match(/getbyprincipalid\((\d+)\)/)[1]); state.assignments = state.assignments.filter(assignment => assignment.PrincipalId !== principal); }
    else if (options.headers?.['X-HTTP-Method'] === 'MERGE') Object.assign(state.list, options.body);
    else throw new Error(`Unexpected write ${path}`);
  });
  return { raw, state, setup: feedbackSetup(raw, 'ModernizationFeedback', async () => true) };
}
describe('private feedback setup', () => {
  it('reports missing storage and discovers existing groups without writes', async () => {
    const { raw, setup } = setupFixture();
    expect(await setup.status()).toMatchObject({ ready: false, title: 'ModernizationFeedback' });
    expect((await setup.options()).selected).toEqual([20, 21, 22]);
    expect(raw.post).not.toHaveBeenCalled();
  });
  it('creates and verifies a private list, granting add-only access last and removing creator access', async () => {
    const { raw, state, setup } = setupFixture();
    expect(await setup.provision([20, 21, 22])).toMatchObject({ ready: true });
    expect(state.list).toMatchObject({ HasUniqueRoleAssignments: true, Hidden: false, NoCrawl: true, EnableAttachments: false });
    expect(state.assignments).toEqual([{ PrincipalId: 10, RoleDefinitionBindings: [{ Id: 1 }] }, ...[20, 21, 22].map(PrincipalId => ({ PrincipalId, RoleDefinitionBindings: [{ Id: 2 }] }))]);
    expect(state.roles[1].BasePermissions).toEqual(SUBMIT_PERMISSIONS);
    const calls = raw.post.mock.calls.map(([path]) => path);
    expect(calls.findIndex(path => path.includes('principalid=20'))).toBeGreaterThan(calls.findLastIndex(path => path.includes('createfieldasxml')));
    expect(await setup.status()).toMatchObject({ ready: true });
  });
  it('denies non-owners and owners lacking schema permissions before writes', async () => {
    const { raw } = setupFixture();
    await expect(feedbackSetup(raw, 'Feedback', async () => false).provision([20])).rejects.toThrow('Only site owners');
    raw.get.mockResolvedValue({ Low: '33554432' });
    await expect(feedbackSetup(raw, 'Feedback', async () => true).provision([20])).rejects.toThrow('Manage Lists');
    expect(raw.post).not.toHaveBeenCalled();
  });
  it('does not interpret access denied as a missing list', async () => {
    const { raw, setup } = setupFixture(); raw.get.mockRejectedValue(Object.assign(new Error('Access denied'), { status: 403 }));
    await expect(setup.status()).rejects.toThrow('Access denied'); expect(raw.post).not.toHaveBeenCalled();
  });
  it('rejects an unsafe existing role before list creation', async () => {
    const { raw, state, setup } = setupFixture(); state.roles.push({ Id: 3, Name: 'ModernizationFeedback Submit Only', BasePermissions: { Low: '3', High: '0' } });
    await expect(setup.provision([20])).rejects.toThrow('not add-only'); expect(raw.post).not.toHaveBeenCalled();
  });
  it('does not overwrite an existing list', async () => {
    const { raw, state, setup } = setupFixture(); state.list = { Description: 'Other list', ItemCount: 1 };
    await expect(setup.provision([20])).rejects.toThrow('already exists'); expect(raw.post).not.toHaveBeenCalled();
  });
  it('keeps an interrupted setup private and resumes an empty app-created list', async () => {
    const { raw, state, setup } = setupFixture(); state.failField = true;
    await expect(setup.provision([20])).rejects.toThrow('Field creation failed');
    expect(state.assignments.every(assignment => [8, 10].includes(assignment.PrincipalId))).toBe(true);
    expect(await setup.status()).toMatchObject({ ready: false, partial: true });
    state.failField = false; await setup.provision([20]);
    expect(raw.post.mock.calls.filter(([path]) => path === '/_api/web/lists')).toHaveLength(1);
    expect(state.assignments.map(assignment => assignment.PrincipalId)).toEqual([10, 20]);
  });
  it('detects reading access granted outside Owners on a configured list', async () => {
    const { state, setup } = setupFixture(); await setup.provision([20]);
    state.assignments.push({ PrincipalId: 999, RoleDefinitionBindings: [{ Id: 1 }] });
    await expect(setup.status()).rejects.toThrow('reading access outside');
  });
  it('refuses unknown group IDs', async () => {
    const { raw, setup } = setupFixture(); await expect(setup.provision([999])).rejects.toThrow('Select the existing'); expect(raw.post).not.toHaveBeenCalled();
  });
});
