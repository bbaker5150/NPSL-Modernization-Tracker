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
  it('keeps distinct claims accounts with a shared email separate for roles and project edits', async () => {
    const { store, state, raw } = fixture('Project Engineer');
    const colleagueLogin = 'i:0#.f|membership|other-account@example.test';
    state.users.unshift({ id: 'colleague', loginName: colleagueLogin, email: user.email, role: 'Manager' });
    Object.assign(state.projects[0], { ownerKey: user.loginName, ownerEmail: user.email });
    Object.assign(state.projects[1], { ownerKey: colleagueLogin, ownerEmail: user.email });
    expect(isManager(user, state.users)).toBe(false);
    expect((await store.load()).projects.map(row => row.id)).toEqual(['p1']);
    await expect(store.saveProject({ ...state.projects[1], ownerKey: user.loginName })).rejects.toThrow('assigned project engineers');
    await expect(store.saveTask({ ...state.tasks[1], projectKey: 'own' })).rejects.toThrow('assigned project engineers');
    expect(raw.saveProject).not.toHaveBeenCalled();
    expect(raw.saveTask).not.toHaveBeenCalled();
  });
  it('lets engineers update tracking on existing assigned tasks and preserves the saved task definition', async () => {
    const { store, raw, state } = fixture('Project Engineer');
    Object.assign(state.tasks[0], { phaseKey: 'acquisition', order: 3 });
    const saved = await store.saveTask({ ...task, spId: 12, projectKey: 'other', phaseKey: 'deployment', order: 72, title: 'Revised', status: 'Blocked', dueDate: '2099-01-01', deferredDate: '2026-10-01', deferredJustification: 'Parts delayed', organization: 'CHENG Team', estimatedHours: 2.5, notes: 'New notes' });
    expect(saved).toMatchObject({ spId: 11, id: 't1', projectKey: 'own', phaseKey: 'acquisition', order: 3, title: 'Review', status: 'Blocked', dueDate: '2026-09-01', deferredDate: '2026-10-01', deferredJustification: 'Parts delayed', organization: 'NAVAIR METCAL 260B', estimatedHours: 2.5, notes: 'New notes' });
    expect(raw.saveTask).toHaveBeenCalledExactlyOnceWith(saved, { trackingOnly: true });
    await expect(store.saveTask({ ...task, id: 'new', spId: 11 })).rejects.toThrow('Only managers can create tasks');
    await expect(store.recycle('tasks', 12, 't1')).rejects.toThrow('Only managers');
    await expect(store.recycle('projects', 1, 'p1')).rejects.toThrow('Only managers');
    expect(raw.saveTask).toHaveBeenCalledTimes(1);
    expect(raw.recycle).not.toHaveBeenCalled();
    raw.saveTasks = vi.fn();
    expect(store.saveTasks).toBeUndefined();
  });
  it('saves only engineer project tracking fields and never forwards a role assignment', async () => {
    const { store, raw, state } = fixture('Project Engineer');
    Object.assign(state.projects[0], { title: 'Original', description: 'Definition', measurementArea: 'Voltage', priority: 'High', importedBaseline: true, progressMode: 'phases' });
    raw.saveProjectTracking = vi.fn(async row => row);
    const saved = await store.saveProject({ ...projects[0], spId: 2, projectKey: 'other', title: 'New title', description: 'New description', targetFinish: '2099-01-01', ownerName: 'Forged', ownerKey: 'other', measurementArea: '', priority: 'Low', importedBaseline: false, progressMode: 'tasks', health: 'On Track', status: 'In Progress', currentStageKey: 'procurement' }, 'forged-assignee');
    expect(saved).toEqual({ ...state.projects[0], health: 'On Track', status: 'In Progress', currentStageKey: 'procurement' });
    expect(raw.saveProjectTracking).toHaveBeenCalledExactlyOnceWith(saved);
    expect(store.saveProjectTracking).toBeUndefined();
    expect(raw.saveProject).not.toHaveBeenCalled();
    await store.saveProgressMode('p1', 'tasks');
    expect(raw.saveProjectTracking).toHaveBeenLastCalledWith({ ...state.projects[0], progressMode: 'tasks' }, { progressModeOnly: true });
    await expect(store.saveProject({ ...state.projects[1], ownerEmail: user.email })).rejects.toThrow('assigned project engineers');
    await expect(store.saveUpdate({ id: 'u2', spId: 20, projectKey: 'own', summary: 'Tampered' })).rejects.toThrow('assigned project engineers');
  });
  it('allows managers to change deadlines and all projects', async () => {
    const { store, raw } = fixture('Manager');
    expect((await store.load()).projects).toHaveLength(2);
    expect(await store.saveTask({ ...data.tasks[1], dueDate: '2028-01-01' })).toMatchObject({ dueDate: '2028-01-01' });
    expect(await store.saveProject({ ...projects[1], targetFinish: '2028-01-01' }, 'chosen-assignee')).toMatchObject({ targetFinish: '2028-01-01' });
    expect(raw.saveProject).toHaveBeenCalledWith(expect.objectContaining({ id: 'p2' }), 'chosen-assignee');
    expect(await store.saveTask({ ...task, id: 'new', title: 'New task' })).toMatchObject({ id: 'new', title: 'New task', spId: undefined });
    await store.recycle('tasks', 11, 't1');
    expect(raw.recycle).toHaveBeenCalledWith('tasks', 11, 't1');
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
  it('does not let the retired checkbox override an explicit Viewer role', async () => {
    for (const role of ['Viewer']) {
      const { store, state } = fixture(role); state.projects[0].ownerCanEdit = true;
      expect((await store.load()).projects).toHaveLength(2);
      await expect(store.saveTask(task)).rejects.toThrow();
    }
  });
  it('migrates existing project owners by identity and persists roles when a manager opens the app', async () => {
    localStorage.clear(); window.MOD_TRACKER_CONFIG = { forceLocal: true };
    const raw = createRepository().store;
    await raw.saveUser({ id: 'manager', title: 'Local Engineer', loginName: 'local', role: 'Manager' });
    await raw.saveUser({ id: 'owner', loginName: 'i:0#.f|membership|owner@example.test', role: 'User' });
    await raw.saveUser({ id: 'outsider', loginName: 'other', role: 'SME' });
    await raw.saveUser({ id: 'viewer', loginName: 'viewer', role: 'Viewer' });
    await raw.saveProject({ id: 'old', projectKey: 'old', ownerEmail: 'OWNER@EXAMPLE.TEST', ownerCanEdit: false });
    await raw.saveProject({ id: 'view', projectKey: 'view', ownerKey: 'viewer', ownerCanEdit: true });
    await authorizedStore(raw).registerCurrentUser();
    expect(raw.data.users.map(({ id, role }) => [id, role])).toEqual([['manager', 'Manager'], ['owner', 'Project Engineer'], ['outsider', 'Viewer'], ['viewer', 'Viewer']]);
    await raw.saveProject({ id: 'old', projectKey: 'old', ownerKey: 'new' });
    expect((await raw.load()).users.find(row => row.id === 'owner').role).toBe('Project Engineer');
    expect(visibleData(await raw.load(), { email: 'owner@example.test' }).projects).toHaveLength(0);
  });
  it('registers an existing project owner missing from the directory as an engineer', async () => {
    localStorage.clear(); window.MOD_TRACKER_CONFIG = { forceLocal: true };
    const raw = createRepository().store;
    await raw.saveProject({ id: 'old', projectKey: 'old', ownerKey: 'local' });
    await authorizedStore(raw).registerCurrentUser();
    expect(raw.data.users[0].role).toBe('Project Engineer');
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
  it('checks live group membership for scoped email lookups without loading the portfolio', async () => {
    const { raw } = fixture('Manager');
    raw.load = vi.fn(raw.load);
    raw.permissions = { currentRole: vi.fn().mockResolvedValueOnce({ role: 'Viewer' }).mockResolvedValue({ role: 'Manager' }) };
    raw.searchPeople = vi.fn(async () => [user]);
    const store = authorizedStore(raw);
    await expect(store.searchPeople(user.email)).rejects.toThrow('Only managers');
    await expect(store.searchPeople(user.email)).resolves.toEqual([user]);
    expect(raw.permissions.currentRole).toHaveBeenCalledWith(user);
    expect(raw.searchPeople).toHaveBeenCalledTimes(1);
    expect(raw.load).not.toHaveBeenCalled();
  });

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

it('preserves assignment and original due dates while allowing engineers to defer', async () => {
  const { store, state } = fixture('Project Engineer');
  Object.assign(state.tasks[0], { assignedDate: '2026-08-01', deferredDate: '2026-10-01', deferredJustification: 'Approved' });
  const result = await store.saveTask({ ...state.tasks[0], assignedDate: '2099-01-01', deferredDate: '2026-11-01', deferredJustification: 'Updated reason' });
  expect(result).toMatchObject({ assignedDate: '2026-08-01', dueDate: '2026-09-01', deferredDate: '2026-11-01', deferredJustification: 'Updated reason' });
  await expect(store.saveTask({ ...task, id: 'new', assignedDate: '2099-01-01' })).rejects.toThrow('Only managers can create tasks');
});
it('validates METCAL responsibilities and clears them when changing organization', async () => {
  const { store } = fixture('Manager');
  expect(await store.saveTask({ ...task, organization: 'PMA 260B', subAssignment: 'Team 2 Lead' })).toMatchObject({ organization: 'NAVAIR METCAL 260B', subAssignment: 'Team 2 Lead' });
  await expect(store.saveTask({ ...task, organization: 'PMA 260B', subAssignment: 'Unknown' })).rejects.toThrow('METCAL');
  expect(await store.saveTask({ ...task, organization: 'NPSL', subAssignment: 'Team 2 Lead' })).toMatchObject({ subAssignment: '' });
});
it('stamps updates with the signed-in identity and rejects task references from another project', async () => {
  const { store, raw } = fixture('Project Engineer');
  const saved = await store.saveUpdate({ id: 'new', projectKey: 'own', summary: ' Decision ', taskId: 't1', authorName: 'Forged', postedAt: '1900-01-01' });
  expect(saved).toMatchObject({ summary: 'Decision', taskId: 't1', taskTitle: 'Review', authorName: user.email, authorKey: user.loginName });
  expect(saved.postedAt).toMatch(/^\d{4}-\d{2}-\d{2}T/);
  await expect(store.saveUpdate({ id: 'new', projectKey: 'own', summary: 'Decision', taskId: 't2' })).rejects.toThrow('this project');
  await store.saveUpdate({ id: 'none', projectKey: 'own', summary: 'No task required' });
  expect(raw.saveUpdate).toHaveBeenLastCalledWith(expect.objectContaining({ taskId: '' }));
});
it('checks update attachment scope and permissions independently of task documents', async () => {
  const { store, state, raw } = fixture('Project Engineer');
  state.updates.push({ id: 'u1', spId: 21, projectKey: 'own' });
  raw.addUpdateAttachment = vi.fn(); raw.listUpdateAttachments = vi.fn(); raw.downloadUpdateAttachment = vi.fn();
  const file = new File(['content'], 'decision.txt');
  await store.addUpdateAttachment('u1', file);
  expect(raw.addUpdateAttachment).toHaveBeenCalledWith(state.updates[1], file);
  expect(raw.addTaskAttachment).not.toHaveBeenCalled();
  await expect(store.addUpdateAttachment('u2', file)).rejects.toThrow('cannot access');
  state.users = directory('Viewer');
  await expect(store.addUpdateAttachment('u1', file)).rejects.toThrow('Only managers');
  await store.downloadUpdateAttachment('u2', 'decision.txt');
  expect(raw.downloadUpdateAttachment).toHaveBeenCalledWith(state.updates[0], 'decision.txt');
});

function updateFixture(role) {
  const f = fixture(role);
  f.state.updates.push({ id: 'u1', spId: 21, projectKey: 'own', summary: 'Original', authorName: 'Original author', authorKey: 'original', postedAt: '2026-11-01T12:00:00Z', version: '"3"' });
  f.state.updates[0].version = '"1"';
  f.raw.editUpdate = vi.fn(async row => row); f.raw.deleteUpdate = vi.fn();
  return { ...f, update: f.state.updates[1] };
}
it.each(['Manager', 'Project Engineer'])('allows %s to edit one update, preserving identity and stamping the editor', async role => {
  const { store, raw, update } = updateFixture(role);
  const saved = await store.editUpdate({ ...update, summary: ' Revised ', taskId: 't1', projectKey: 'other', spId: 99, authorName: 'Forged', postedAt: '1900-01-01', editorName: 'Forged' });
  expect(saved).toMatchObject({ summary: 'Revised', taskTitle: 'Review', projectKey: 'own', spId: 21, authorName: 'Original author', postedAt: update.postedAt, editorName: user.email, editorKey: user.loginName });
  expect(saved.editedAt).toMatch(/^\d{4}-\d{2}-\d{2}T/);
  await store.deleteUpdate({ ...update, spId: 99, projectKey: 'other' });
  expect(raw.deleteUpdate).toHaveBeenCalledExactlyOnceWith(update);
});
it('rechecks roles, ownership, task scope and persisted versions for every update change', async () => {
  const { store, state, raw, update } = updateFixture('Project Engineer');
  await expect(store.editUpdate({ ...update, taskId: 't2' })).rejects.toThrow('this project');
  await expect(store.editUpdate({ ...update, summary: ' ' })).rejects.toThrow('Enter an update');
  for (const method of ['editUpdate', 'deleteUpdate']) {
    for (const version of [undefined, '*', '"2"']) await expect(store[method]({ ...update, version })).rejects.toThrow('changed');
    await expect(store[method]({ ...state.updates[0], projectKey: 'own' })).rejects.toThrow('Only managers');
    await expect(store[method]({ ...update, id: 'missing' })).rejects.toThrow('existing project update');
  }
  state.users = directory('Viewer');
  await expect(store.editUpdate(update)).rejects.toThrow('Only managers');
  await expect(store.deleteUpdate(update)).rejects.toThrow('Only managers');
  expect(raw.editUpdate).not.toHaveBeenCalled(); expect(raw.deleteUpdate).not.toHaveBeenCalled();
});
it('retains an unchanged historical task reference and prevents saveUpdate bypassing edit checks', async () => {
  const { store, update } = updateFixture('Manager');
  Object.assign(update, { taskId: 'removed', taskTitle: 'Historical task' });
  expect(await store.editUpdate(update)).toMatchObject({ taskId: 'removed', taskTitle: 'Historical task' });
  expect(await store.editUpdate({ ...update, taskId: '' })).toMatchObject({ taskId: '', taskTitle: '' });
  await expect(store.saveUpdate(update)).rejects.toThrow('Use Edit update');
});
