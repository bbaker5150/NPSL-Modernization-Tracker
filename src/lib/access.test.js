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
    expect(isManager({ isSiteAdmin: true })).toBe(true);
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
    expect(() => validateTask({ ...task, deferredDate: '2026-02-30', deferredJustification: 'Delay' })).toThrow();
  });
  it('restricts every standard-user write to permitted task fields', async () => {
    const raw = { currentUser: async () => user, load: async () => structuredClone(data), saveTask: vi.fn(async (row) => row), saveProject: vi.fn(), recycle: vi.fn(), saveUser: vi.fn() };
    const store = authorizedStore(raw);
    const saved = await store.saveTask({ ...task, title: 'Tampered', dueDate: '2099-01-01', ownerEmail: 'attacker', status: 'In Progress', deferredDate: '2026-10-01', deferredJustification: 'Vendor delay' });
    expect(saved).toMatchObject({ title: 'Review', dueDate: '2026-09-01', ownerEmail: user.email, status: 'In Progress', deferredDate: '2026-10-01' });
    await expect(store.saveTask({ ...data.tasks[1], status: 'Complete' })).rejects.toThrow('assigned');
    await expect(store.saveTask({ ...task, id: 'new' })).rejects.toThrow('assigned');
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
});
