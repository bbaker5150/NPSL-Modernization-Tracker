import { beforeEach, describe, expect, it, vi } from 'vitest';
import { authorizedStore, isManager, visibleData, validateTask } from './access';
import { createRepository } from './repository';

const user = { loginName: 'i:0#.f|membership|engineer@example.test', email: 'engineer@example.test' };
const projects = [{ id: 'p1', projectKey: 'own', ownerEmail: user.email }, { id: 'p2', projectKey: 'other', ownerEmail: 'other@example.test' }];
const task = { id: 't1', projectKey: 'own', title: 'Review', status: 'Not Started', dueDate: '2026-09-01', ownerEmail: user.email };
const data = { projects, tasks: [task, { ...task, id: 't2', projectKey: 'other', ownerEmail: 'other@example.test' }], updates: [{ projectKey: 'other', summary: 'private' }], risks: [], acronyms: [], users: [] };

describe('role and task authorization', () => {
  it('defaults unknown identities to standard users and recognizes claims identities', () => {
    expect(isManager(user)).toBe(false);
    expect(isManager(user, [{ email: user.email.toUpperCase(), role: 'Manager' }])).toBe(true);
    expect(isManager({ isSiteAdmin: true })).toBe(false);
    const scoped = visibleData(data, user);
    expect(scoped.projects.map((row) => row.id)).toEqual(['p1']);
    expect(scoped.tasks.map((row) => row.id)).toEqual(['t1']);
    expect(scoped.updates).toEqual([]);
  });
  it('requires both kinds of justification and keeps original dates', () => {
    expect(() => validateTask({ ...task, deferredDate: '2026-10-01' })).toThrow('justification');
    expect(() => validateTask({ ...task, deferredDate: '2026-08-01', deferredJustification: 'Vendor delay' })).toThrow('after');
    expect(() => validateTask({ ...task, status: 'Not Required' })).toThrow('justification');
    expect(() => validateTask({ ...task, status: 'Not Required', notRequiredJustification: 'Outside scope' })).not.toThrow();
    expect(() => validateTask({ ...task, assignedDate: '2026-02-30' })).toThrow();
    expect(() => validateTask({ ...task, deferredDate: '2026-02-30', deferredJustification: 'Delay' })).toThrow();
  });
  it('restricts every standard-user write to permitted task fields', async () => {
    const raw = { currentUser: async () => user, load: async () => structuredClone(data), saveTask: vi.fn(async (row) => row), saveProject: vi.fn(), recycle: vi.fn(), saveUser: vi.fn() };
    const store = authorizedStore(raw);
    const saved = await store.saveTask({ ...task, title: 'Tampered', assignedDate: '2099-01-01', dueDate: '2099-01-01', ownerEmail: 'attacker', status: 'In Progress – At Program Office', deferredDate: '2026-10-01', deferredJustification: 'Vendor delay' });
    expect(saved).toMatchObject({ title: 'Review', dueDate: '2026-09-01', ownerEmail: user.email, status: 'Not Started', deferredDate: '2026-10-01' });
    await expect(store.saveTask({ ...data.tasks[1], status: 'Complete' })).rejects.toThrow('assigned');
    await expect(store.saveTask({ ...task, id: 'new', projectKey: 'other' })).rejects.toThrow('authorized project owners');
    await expect(store.saveTask({ ...task, id: 'new', spId: 99, ownerKey: 'other', dueDate: '2099-01-01', status: 'Complete' })).rejects.toThrow('authorized project owners');
    expect(saved.assignedDate).toBeUndefined();
    await expect(store.saveProject({})).rejects.toThrow('authorized project owners');
    await expect(store.saveUser({})).rejects.toThrow('Only managers');
    await expect(store.recycle({})).rejects.toThrow();
    expect(raw.saveProject).not.toHaveBeenCalled();
    expect((await store.load()).projects).toHaveLength(1);
  });
  it('checks current roles again on each mutation', async () => {
    const changing = structuredClone(data);
    changing.users = [{ email: user.email, role: 'Manager' }];
    const raw = { currentUser: async () => user, load: async () => changing, saveProject: vi.fn(async (row) => row) };
    const store = authorizedStore(raw);
    await store.saveProject(projects[0]);
    changing.users[0].role = 'User';
    await expect(store.saveProject(projects[0])).rejects.toThrow('Only managers');
    expect(raw.saveProject).toHaveBeenCalledTimes(1);
  });
  it('recycles only the selected local task and persists directory entries', async () => {
    localStorage.clear(); window.MOD_TRACKER_CONFIG = { forceLocal: true };
    const store = createRepository().store;
    await store.saveTask(task); await store.saveTask(data.tasks[1]);
    await store.recycle('tasks', undefined, task.id);
    expect((await store.load()).tasks.map((row) => row.id)).toEqual(['t2']);
    await store.saveUser({ id: 'u1', title: 'Engineer', email: user.email, role: 'User' });
    expect((await createRepository().store.load()).users[0].title).toBe('Engineer');
  });
  it('registers only the current identity, preserves roles, and supports disabling testing access', async () => {
    localStorage.clear(); window.MOD_TRACKER_CONFIG = { forceLocal: true };
    const raw = createRepository().store;
    const store = authorizedStore(raw);
    await store.registerCurrentUser();
    await store.registerCurrentUser();
    expect((await store.load()).users).toHaveLength(1);
    expect((await store.load()).users[0].role).toBe('User');
    await expect(store.activateTestingManager('wrong')).rejects.toThrow('incorrect');
    await expect(authorizedStore(raw, { testingManagerPassword: false }).activateTestingManager('admin123')).rejects.toThrow('disabled');
    await store.activateTestingManager('admin123');
    await store.registerCurrentUser();
    const loaded = await createRepository().store.load();
    expect(loaded.users).toHaveLength(1);
    expect(loaded.users[0]).toMatchObject({ role: 'Manager', loginName: 'local' });
  });

  it('lets only manager-authorized owners change project progress mode', async () => {
    const ownerData = structuredClone(data); ownerData.projects[0].ownerCanEdit = true;
    const raw = { currentUser: async () => user, load: async () => ownerData, saveProject: vi.fn(async (row) => row) };
    const store = authorizedStore(raw);
    expect(await store.saveProgressMode('p1', 'tasks')).toEqual({ ...ownerData.projects[0], progressMode: 'tasks' });
    await expect(store.saveProgressMode('p2', 'tasks')).rejects.toThrow('owner');
    await expect(store.saveProgressMode('p1', 'invalid')).rejects.toThrow('valid');
  });

});

describe('standard user task documents', () => {
  it('treats legacy SME entries as standard scoped users', async () => {
    const smeData = { ...structuredClone(data), users: [{ id: 'sme', email: user.email, role: 'SME' }] };
    const raw = { currentUser: async () => user, load: async () => smeData, listTaskAttachments: vi.fn(async () => [{ name: 'review.pdf', url: '/review.pdf' }]), addTaskAttachment: vi.fn(), saveTask: vi.fn(), saveProject: vi.fn(), saveUser: vi.fn(), recycle: vi.fn() };
    const store = authorizedStore(raw);
    expect((await store.load()).projects).toHaveLength(1);
    await expect(store.listTaskAttachments('t2')).rejects.toThrow('cannot access');
    await store.saveTask({ ...task, status: 'Complete' });
    expect(raw.saveTask).toHaveBeenCalledWith(expect.objectContaining({ status: 'Not Started' }));
    await expect(store.saveProgressMode('p1', 'tasks')).rejects.toThrow('authorized project owners');
    await expect(store.addTaskAttachment('t1', new File(['x'], 'a.txt'))).resolves.toBeUndefined();
    await expect(store.saveUser({ role: 'Manager' })).rejects.toThrow('Only managers');
    await store.activateTestingManager('admin123');
    expect(raw.addTaskAttachment).toHaveBeenCalled();
  });
  it('allows a standard user to attach only to authorized saved tasks without saving task fields', async () => {
    const raw = { currentUser: async () => user, load: async () => structuredClone(data), addTaskAttachment: vi.fn(async () => []), listTaskAttachments: vi.fn(async () => []), saveTask: vi.fn() };
    const store = authorizedStore(raw);
    const file = new File(['report'], 'report.txt');
    await store.addTaskAttachment('t1', file);
    expect(raw.addTaskAttachment).toHaveBeenCalledWith(task, file);
    expect(raw.saveTask).not.toHaveBeenCalled();
    await expect(store.addTaskAttachment('t2', file)).rejects.toThrow('cannot access');
    await expect(store.listTaskAttachments('unknown')).rejects.toThrow('cannot access');
  });
  it('accepts optional estimated hours and rejects negative or nonnumeric values', () => {
    for (const estimatedHours of [null, '', 0, 2.5]) expect(() => validateTask({ ...task, estimatedHours })).not.toThrow();
    for (const estimatedHours of [-1, 'bad', Infinity]) expect(() => validateTask({ ...task, estimatedHours })).toThrow('Estimated hours');
  });
});


describe('project document access', () => {
  it('does not expose other task documents to a task-only assignee', async () => {
    const state = structuredClone(data);
    state.projects[0].ownerEmail = 'other@example.test';
    state.tasks.push({ ...task, id: 'private-task', ownerEmail: 'other@example.test' });
    const raw = { currentUser: async () => user, load: async () => state, listTaskAttachments: vi.fn(async () => [{ name: 'report.pdf' }]) };
    const store = authorizedStore(raw);
    expect(await store.listProjectAttachments('own')).toEqual([{ name: 'report.pdf', taskId: 't1', taskTitle: 'Review' }]);
    expect(raw.listTaskAttachments).toHaveBeenCalledTimes(1);
  });
  it('allows checked owners to manage tasking and immediately honors revocation', async () => {
    const state = structuredClone(data);
    state.projects[0].ownerCanEdit = true;
    const raw = { currentUser: async () => user, load: async () => state, saveTask: vi.fn(async (row) => row), saveProject: vi.fn(async (row) => row), recycle: vi.fn() };
    const store = authorizedStore(raw);
    expect(await store.saveTask({ ...task, title: 'Owner edit', status: 'Complete' })).toMatchObject({ title: 'Owner edit', status: 'Complete' });
    expect(await store.saveTask({ ...task, id: 'new' })).toMatchObject({ id: 'new' });
    await store.recycle('tasks', undefined, 't1');
    expect(await store.saveProject({ ...state.projects[0], title: 'Updated', ownerEmail: 'other@example.test', ownerCanEdit: false })).toMatchObject({ title: 'Updated', ownerEmail: user.email, ownerCanEdit: true });
    state.projects[0].ownerCanEdit = false;
    expect(await store.saveTask({ ...task, title: 'Forbidden', status: 'Complete' })).toMatchObject({ title: 'Review', status: 'Not Started' });
    await expect(store.saveTask({ ...task, id: 'new' })).rejects.toThrow('authorized project owners');
    await expect(store.recycle('tasks', undefined, 't1')).rejects.toThrow('authorized project owners');
    await expect(store.saveProject({ ...state.projects[0], ownerCanEdit: true })).rejects.toThrow('authorized project owners');
  });
  it('aggregates project tasks and permits only authorized owners or managers to change documents', async () => {
    const state = structuredClone(data);
    state.projects[0].ownerCanEdit = true;
    state.tasks.push({ ...task, id: 't3', ownerEmail: 'someone@example.test' });
    const raw = { currentUser: async () => user, load: async () => state, listTaskAttachments: vi.fn(async () => [{ name: 'report.pdf', url: '/report.pdf' }]), deleteTaskAttachment: vi.fn(), renameTaskAttachment: vi.fn() };
    const store = authorizedStore(raw);
    expect(await store.listProjectAttachments('own')).toMatchObject([{ taskId: 't1', name: 'report.pdf' }, { taskId: 't3', name: 'report.pdf' }]);
    await store.deleteTaskAttachment('t3', 'report.pdf');
    await store.renameTaskAttachment('t3', 'report.pdf', 'renamed.pdf');
    expect(raw.deleteTaskAttachment).toHaveBeenCalledTimes(1);
    expect(raw.renameTaskAttachment).toHaveBeenCalledWith(state.tasks.find((row) => row.id === 't3'), 'report.pdf', 'renamed.pdf');
    state.projects[0].ownerCanEdit = false;
    await expect(store.deleteTaskAttachment('t1', 'report.pdf')).rejects.toThrow('authorized project owners');
    await expect(store.listProjectAttachments('other')).rejects.toThrow('cannot access');
    state.users = [{ email: user.email, role: 'Manager' }];
    await store.deleteTaskAttachment('t2', 'report.pdf');
    expect(raw.deleteTaskAttachment).toHaveBeenCalledTimes(2);
    state.users[0].role = 'User';
    await expect(store.deleteTaskAttachment('t2', 'report.pdf')).rejects.toThrow('cannot access');
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
