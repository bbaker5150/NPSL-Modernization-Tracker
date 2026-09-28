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
    expect(saved).toMatchObject({ title: 'Review', dueDate: '2026-09-01', ownerEmail: user.email, status: 'In Progress – At Program Office', deferredDate: '2026-10-01' });
    await expect(store.saveTask({ ...data.tasks[1], status: 'Complete' })).rejects.toThrow('assigned');
    await expect(store.saveTask({ ...task, id: 'new', projectKey: 'other' })).rejects.toThrow('own project');
    const added = await store.saveTask({ ...task, id: 'new', spId: 99, ownerKey: 'other', dueDate: '2099-01-01', status: 'Complete' });
    expect(added).toMatchObject({ status: 'Not Started', dueDate: '', ownerEmail: user.email });
    expect(added.spId).toBeUndefined();
    expect(added.assignedDate).toBe(new Date().toISOString().slice(0, 10));
    expect(saved.assignedDate).toBeUndefined();
    for (const method of ['saveProject', 'saveUser', 'recycle']) await expect(store[method]({})).rejects.toThrow('Only managers');
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

  it('lets owners change only the progress mode on their projects', async () => {
    const raw = { currentUser: async () => user, load: async () => structuredClone(data), saveProject: vi.fn(async (row) => row) };
    const store = authorizedStore(raw);
    expect(await store.saveProgressMode('p1', 'tasks')).toEqual({ ...projects[0], progressMode: 'tasks' });
    await expect(store.saveProgressMode('p2', 'tasks')).rejects.toThrow('owner');
    await expect(store.saveProgressMode('p1', 'invalid')).rejects.toThrow('valid');
  });

});

describe('SME viewing and task documents', () => {
  it('allows SME portfolio and attachment reads but denies mutations even on owned tasks', async () => {
    const smeData = { ...structuredClone(data), users: [{ id: 'sme', email: user.email, role: 'SME' }] };
    const raw = { currentUser: async () => user, load: async () => smeData, listTaskAttachments: vi.fn(async () => [{ name: 'review.pdf', url: '/review.pdf' }]), addTaskAttachment: vi.fn(), saveTask: vi.fn(), saveProject: vi.fn(), saveUser: vi.fn(), recycle: vi.fn() };
    const store = authorizedStore(raw);
    expect((await store.load()).projects).toHaveLength(2);
    await expect(store.listTaskAttachments('t2')).resolves.toHaveLength(1);
    await expect(store.saveTask({ ...task, status: 'Complete' })).rejects.toThrow('read-only');
    await expect(store.saveProgressMode('p1', 'tasks')).rejects.toThrow('read-only');
    await expect(store.addTaskAttachment('t1', new File(['x'], 'a.txt'))).rejects.toThrow('cannot attach');
    await expect(store.saveUser({ role: 'Manager' })).rejects.toThrow('Only managers');
    await expect(store.activateTestingManager('admin123')).rejects.toThrow('Ask a manager');
    expect(raw.addTaskAttachment).not.toHaveBeenCalled();
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
  it('aggregates all project tasks and permits only owners or managers to delete', async () => {
    const state = structuredClone(data);
    state.tasks.push({ ...task, id: 't3', ownerEmail: 'someone@example.test' });
    const raw = { currentUser: async () => user, load: async () => state, listTaskAttachments: vi.fn(async () => [{ name: 'report.pdf', url: '/report.pdf' }]), deleteTaskAttachment: vi.fn() };
    const store = authorizedStore(raw);
    expect(await store.listProjectAttachments('own')).toMatchObject([{ taskId: 't1', name: 'report.pdf' }, { taskId: 't3', name: 'report.pdf' }]);
    await store.deleteTaskAttachment('t3', 'report.pdf');
    expect(raw.deleteTaskAttachment).toHaveBeenCalledTimes(1);
    state.projects[0].ownerEmail = 'different@example.test';
    await expect(store.deleteTaskAttachment('t1', 'report.pdf')).rejects.toThrow('Only project owners');
    await expect(store.listProjectAttachments('other')).rejects.toThrow('cannot access');
    state.users = [{ email: user.email, role: 'Manager' }];
    await store.deleteTaskAttachment('t2', 'report.pdf');
    expect(raw.deleteTaskAttachment).toHaveBeenCalledTimes(2);
    state.users[0].role = 'SME';
    await expect(store.deleteTaskAttachment('t2', 'report.pdf')).rejects.toThrow('Only project owners');
  });
});
