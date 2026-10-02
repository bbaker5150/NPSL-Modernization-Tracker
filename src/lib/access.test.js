import { describe, expect, it, vi } from 'vitest';
import { authorizedStore, isManager, visibleData, validateTask } from './access';
import { createRepository } from './repository';

const user = { loginName: 'i:0#.f|membership|engineer@example.test', email: 'engineer@example.test' };
const projects = [{ id: 'p1', spId: 1, projectKey: 'own', ownerEmail: user.email, targetFinish: '2027-01-01' }, { id: 'p2', spId: 2, projectKey: 'other', ownerEmail: 'other@example.test' }];
const task = { id: 't1', spId: 11, projectKey: 'own', title: 'Review', status: 'Not Started', dueDate: '2026-09-01', ownerEmail: user.email };
const data = { projects, tasks: [task, { ...task, id: 't2', spId: 12, projectKey: 'other', ownerEmail: user.email }], updates: [{ id: 'u2', spId: 20, projectKey: 'other' }], risks: [], acronyms: [], users: [] };
const directory = (role) => [{ id: 'me', email: user.email, loginName: user.loginName, role }];
function fixture(role) {
  const state = structuredClone(data); state.users = directory(role);
  const raw = { currentUser: async () => user, load: async () => state };
  for (const method of ['saveTask', 'saveProject', 'saveUpdate', 'saveRisk', 'saveUser', 'saveAcronym', 'recycle', 'addTaskAttachment', 'deleteTaskAttachment', 'renameTaskAttachment', 'downloadTaskAttachment']) raw[method] = vi.fn(async (row) => row);
  raw.listTaskAttachments = vi.fn(async () => [{ name: 'report.pdf' }]);
  return { state, raw, store: authorizedStore(raw) };
}

describe('three-role access model', () => {
  it('recognizes claims identities and gives viewers all project reads without edit privileges', async () => {
    const { store, state, raw } = fixture('Viewer');
    expect(isManager(user, [{ email: user.email.toUpperCase(), role: 'Manager' }])).toBe(true);
    expect(isManager({ isSiteAdmin: true })).toBe(false);
    expect(visibleData(state, user).projects).toHaveLength(2);
    expect(await store.listProjectAttachments('other')).toHaveLength(1);
    await store.downloadTaskAttachment('t2', 'report.pdf');
    for (const operation of [() => store.saveTask({ ...task, status: 'Complete' }), () => store.saveTask({ ...task, id: 'new' }), () => store.saveProject(projects[0]), () => store.saveUpdate({ projectKey: 'own' }), () => store.saveRisk({ projectKey: 'own' }), () => store.recycle('tasks', 11, 't1'), () => store.addTaskAttachment('t1', new File(['a'], 'a.txt')), () => store.deleteTaskAttachment('t1', 'report.pdf'), () => store.renameTaskAttachment('t1', 'report.pdf', 'new.pdf')]) await expect(operation()).rejects.toThrow();
    expect(raw.saveTask).not.toHaveBeenCalled();
    expect(raw.addTaskAttachment).not.toHaveBeenCalled();
  });
  it('scopes engineers to project assignments, not merely task assignments', async () => {
    const { store } = fixture('Project Engineer');
    expect((await store.load()).projects.map((row) => row.id)).toEqual(['p1']);
    expect((await store.load()).tasks.map((row) => row.id)).toEqual(['t1']);
    await expect(store.saveTask(data.tasks[1])).rejects.toThrow('assigned project engineers');
    await expect(store.listProjectAttachments('other')).rejects.toThrow('cannot access');
    await expect(store.downloadTaskAttachment('t2', 'report.pdf')).rejects.toThrow('cannot access');
    await expect(store.saveProject({ id: 'new', projectKey: 'new', ownerEmail: user.email })).rejects.toThrow('assigned project engineers');
  });
  it('lets engineers edit task content and deferrals while preserving original deadlines and record identity', async () => {
    const { store, raw } = fixture('Project Engineer');
    const saved = await store.saveTask({ ...task, spId: 12, projectKey: 'other', title: 'Revised', status: 'Blocked', dueDate: '2099-01-01', deferredDate: '2026-10-01', deferredJustification: 'Parts delayed', organization: 'CHENG Team', estimatedHours: 2.5, notes: 'New notes' });
    expect(saved).toMatchObject({ spId: 11, projectKey: 'own', title: 'Revised', status: 'Blocked', dueDate: '2026-09-01', deferredDate: '2026-10-01', organization: 'CHENG Team', estimatedHours: 2.5 });
    const added = await store.saveTask({ ...task, id: 'new', spId: 12, dueDate: '2099-01-01' });
    expect(added).toMatchObject({ dueDate: '', spId: undefined });
    await store.recycle('tasks', 12, 't1');
    expect(raw.recycle).toHaveBeenCalledWith('tasks', 11, 't1');
  });
  it('protects target completion and prevents modifying another project history item', async () => {
    const { store } = fixture('Project Engineer');
    const saved = await store.saveProject({ ...projects[0], spId: 2, projectKey: 'other', title: 'New title', description: 'New description', targetFinish: '2099-01-01' });
    expect(saved).toMatchObject({ spId: 1, projectKey: 'own', title: 'New title', description: 'New description', targetFinish: '2027-01-01' });
    await expect(store.saveUpdate({ id: 'u2', spId: 20, projectKey: 'own', summary: 'Tampered' })).rejects.toThrow('assigned project engineers');
  });
  it('allows managers to change deadlines and all projects', async () => {
    const { store } = fixture('Manager');
    expect((await store.load()).projects).toHaveLength(2);
    expect(await store.saveTask({ ...data.tasks[1], dueDate: '2028-01-01' })).toMatchObject({ dueDate: '2028-01-01' });
    expect(await store.saveProject({ ...projects[1], targetFinish: '2028-01-01' })).toMatchObject({ targetFinish: '2028-01-01' });
  });
  it('honors role and project-assignment revocation on every operation', async () => {
    const { store, state } = fixture('Project Engineer');
    await store.saveProgressMode('p1', 'tasks');
    await store.addTaskAttachment('t1', new File(['a'], 'a.txt'));
    state.projects[0].ownerEmail = 'other@example.test';
    await expect(store.saveTask(task)).rejects.toThrow();
    state.projects[0].ownerEmail = user.email; state.users[0].role = 'Viewer';
    await expect(store.saveProgressMode('p1', 'phases')).rejects.toThrow();
    await expect(store.addTaskAttachment('t1', new File(['a'], 'a.txt'))).rejects.toThrow();
  });
  it('does not treat legacy SME/User roles or the retired checkbox as engineer permission', async () => {
    for (const role of ['SME', 'User', 'Viewer']) {
      const { store, state } = fixture(role); state.projects[0].ownerCanEdit = true;
      expect((await store.load()).projects).toHaveLength(2);
      await expect(store.saveTask(task)).rejects.toThrow();
    }
  });
  it('requires valid dates and justification', () => {
    expect(() => validateTask({ ...task, deferredDate: '2026-10-01' })).toThrow('justification');
    expect(() => validateTask({ ...task, deferredDate: '2026-08-01', deferredJustification: 'delay' })).toThrow('after');
    expect(() => validateTask({ ...task, status: 'Not Required' })).toThrow('justification');
    expect(() => validateTask({ ...task, assignedDate: '2026-02-30' })).toThrow();
    for (const estimatedHours of [-1, 'bad', Infinity]) expect(() => validateTask({ ...task, estimatedHours })).toThrow();
  });
  it('registers viewers without changing existing roles and supports disabling testing access', async () => {
    localStorage.clear(); window.MOD_TRACKER_CONFIG = { forceLocal: true };
    const raw = createRepository().store; const store = authorizedStore(raw);
    await store.registerCurrentUser(); await store.registerCurrentUser();
    expect((await store.load()).users).toHaveLength(1);
    expect((await store.load()).users[0].role).toBe('Viewer');
    await expect(authorizedStore(raw, { testingManagerPassword: false }).activateTestingManager('admin123')).rejects.toThrow('disabled');
    await store.activateTestingManager('admin123'); await store.registerCurrentUser();
    expect((await store.load()).users[0].role).toBe('Manager');
  });
});

describe('directory search authorization', () => {
  it('limits people search and identity resolution to managers', async () => {
    const state = structuredClone(data);
    const raw = { currentUser: async () => user, load: async () => state, searchPeople: vi.fn(async () => []), resolvePerson: vi.fn(async () => ({ loginName: 'resolved' })) };
    const store = authorizedStore(raw);
    for (const role of ['User']) {
      state.users = [{ email: user.email, role }];
      await expect(store.searchPeople('Engineer')).rejects.toThrow('Only managers');
      await expect(store.resolvePerson('engineer')).rejects.toThrow('Only managers');
    }
    state.users[0].role = 'Manager';
    await expect(store.searchPeople('Engineer')).resolves.toEqual([]);
    await expect(store.resolvePerson('engineer')).resolves.toEqual({ loginName: 'resolved' });
  });
});


it('requires a manager and a saved matching role before granting site access', async () => {
  const state = structuredClone(data);
  const raw = { currentUser: async () => user, load: async () => state, shareSiteAccess: vi.fn(async () => ({ access: 'Read' })) };
  const store = authorizedStore(raw);
  await expect(store.shareSiteAccess({ loginName: 'other' }, 'User', 'https://site')).rejects.toThrow('Only managers');
  state.users = [{ email: user.email, role: 'Manager' }, { loginName: 'other', role: 'User' }];
  await expect(store.shareSiteAccess({ loginName: 'other' }, 'Manager', 'https://site')).rejects.toThrow('Save the person');
  await expect(store.shareSiteAccess({ loginName: 'other' }, 'User', 'https://site')).resolves.toEqual({ access: 'Read' });
  expect(raw.shareSiteAccess).toHaveBeenCalledTimes(1);
});


it('allows managers to delete persisted directory users but never themselves', async () => {
  const state = structuredClone(data);
  state.users = [{ id: 'me', spId: 1, email: user.email, role: 'Manager' }, { id: 'other', spId: 2, loginName: 'other', role: 'User' }];
  const raw = { currentUser: async () => user, load: async () => state, recycle: vi.fn() };
  const store = authorizedStore(raw);
  await expect(store.recycle('users', 2, 'me')).rejects.toThrow('own account');
  await store.recycle('users', 1, 'other');
  expect(raw.recycle).toHaveBeenCalledWith('users', 2, 'other');
  await expect(store.recycle('users', 99, 'missing')).rejects.toThrow('no longer');
  state.users[0].role = 'User';
  await expect(store.recycle('users', 2, 'other')).rejects.toThrow('Only managers');
  expect(raw.recycle).toHaveBeenCalledTimes(1);
});
