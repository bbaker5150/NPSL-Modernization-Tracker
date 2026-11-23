import { expect, it, vi } from 'vitest';
import { authorizedStore } from './access';

function fixture(role = 'Manager') {
  const user = { loginName: 'claims|engineer', email: 'engineer@example.test', title: 'Engineer' };
  const project = { id: 'p', spId: 1, projectKey: 'project', ownerKey: user.loginName, organization: 'NAVAIR METCAL 260B',
    health: 'On Track', status: 'Active', currentStageKey: '', progressMode: 'phases' };
  const task = { id: 't', spId: 2, projectKey: project.projectKey, title: 'Review', phaseKey: 'requirement', order: 1,
    organization: project.organization, subAssignment: '', status: 'In Progress', ownerName: user.title,
    ownerEmail: user.email, ownerKey: user.loginName, estimatedHours: 2, startDate: '', finishDate: '',
    assignedDate: '2026-11-01', dueDate: '2027-01-01', deferredDate: '', deferredJustification: '',
    notRequiredJustification: '', blockedReason: '', notes: '', sourceStartLabel: '', dataIssue: '' };
  const update = { id: 'u', spId: 3, projectKey: project.projectKey, summary: 'Progress', taskId: task.id,
    taskTitle: task.title, version: '"3"', editedAt: '2026-11-01T12:00:00Z', authorName: 'Original author' };
  const risk = { id: 'r', spId: 4, projectKey: project.projectKey, title: 'Risk', mitigation: '', severity: 'Low', status: 'Open', version: '"4"' };
  const references = [{ id: 'ref', spId: 5, kind: 'file', name: 'Guide.pdf', fileName: 'Guide.pdf', parentId: '', size: 123 },
    { id: 'folder', spId: 6, kind: 'folder', name: 'Folder', parentId: '' }];
  const state = { projects: [project], tasks: [task], updates: [update], risks: [risk], references,
    users: [{ id: 'me', loginName: user.loginName, email: user.email, role }],
    acronyms: [{ id: 'term', spId: 7, acronym: 'CSS', term: 'Specification' }] };
  const raw = {
    currentUser: vi.fn(async () => user),
    load: vi.fn(async options => ({ ...state, acronyms: options?.includeAcronyms === false ? [] : state.acronyms })),
    listReferenceEntries: vi.fn(async () => references),
  };
  for (const name of ['saveTask', 'saveProject', 'saveProjectTracking', 'editUpdate', 'editRisk', 'saveReferenceEntry', 'saveAcronym', 'recycle', 'addTaskAttachment']) {
    raw[name] = vi.fn(async value => value);
  }
  return { raw, store: authorizedStore(raw), state, user, project, task, update, risk, references };
}

it('keeps default/export loads complete and explicitly omits glossary reads for authorization', async () => {
  const { store, raw, state, task } = fixture();
  expect((await store.load()).acronyms).toEqual(state.acronyms);
  expect(raw.load).toHaveBeenLastCalledWith({ includeAcronyms: true });
  expect((await store.load({ includeAcronyms: false })).acronyms).toEqual([]);
  expect(raw.load).toHaveBeenLastCalledWith({ includeAcronyms: false });
  await store.saveTask({ ...task, notes: 'Changed' });
  expect(raw.load).toHaveBeenLastCalledWith({ includeAcronyms: false });
});

it('checks current glossary rows for duplicate names and existing-item rights', async () => {
  const { store, raw, state } = fixture('Project Engineer');
  await expect(store.saveAcronym({ id: 'new', acronym: ' css ', term: 'Duplicate' })).rejects.toThrow('already in the glossary');
  expect(raw.load).toHaveBeenLastCalledWith({ includeAcronyms: true });
  await expect(store.saveAcronym({ ...state.acronyms[0], term: 'Changed' })).rejects.toThrow('Only managers');
  await store.saveAcronym({ id: 'new', acronym: ' NEW ', term: ' New term ' });
  expect(raw.saveAcronym).toHaveBeenCalledExactlyOnceWith({ id: 'new', acronym: 'NEW', term: 'New term', definition: '' });
  state.users[0].role = 'Manager';
  await store.recycle('acronyms', 7, 'term');
  expect(raw.load).toHaveBeenLastCalledWith({ includeAcronyms: true });
});

it.each(['Manager', 'Project Engineer'])('returns the current task for normalized unchanged %s drafts and writes real changes', async role => {
  const { store, raw, task } = fixture(role);
  expect(await store.saveTask({ ...task, estimatedHours: '2.00', organization: 'PMA 260B', notes: null, uiExpanded: true })).toBe(task);
  expect(raw.saveTask).not.toHaveBeenCalled();
  await store.saveTask({ ...task, notes: 'New tracking detail' });
  expect(raw.saveTask).toHaveBeenCalledTimes(1);
  expect(raw.saveTask.mock.calls[0][0].notes).toBe('New tracking detail');
  expect(raw.currentUser).toHaveBeenCalledTimes(2);
  expect(raw.load.mock.calls.every(([options]) => options.includeAcronyms === false)).toBe(true);
});

it('retains manager-only task definition changes while ignoring forbidden engineer fields', async () => {
  const { store, raw, task, state } = fixture('Project Engineer');
  expect(await store.saveTask({ ...task, title: 'Forged definition', dueDate: '2028-01-01', spId: 999 })).toBe(task);
  expect(raw.saveTask).not.toHaveBeenCalled();
  state.users[0].role = 'Manager';
  await store.saveTask({ ...task, title: 'Updated definition' });
  expect(raw.saveTask).toHaveBeenCalledExactlyOnceWith(expect.objectContaining({ title: 'Updated definition', spId: task.spId }));
});

it('still permits pending attachment retries after an unchanged task save', async () => {
  const { store, raw, task } = fixture('Project Engineer');
  const saved = await store.saveTask({ ...task });
  const file = new File(['retry'], 'retry.txt');
  await store.addTaskAttachment(saved.id, file);
  expect(raw.saveTask).not.toHaveBeenCalled();
  expect(raw.addTaskAttachment).toHaveBeenCalledExactlyOnceWith(task, file);
});

it('checks revoked roles and project assignments before unchanged task, progress or project returns', async () => {
  const { store, raw, task, project, state } = fixture('Project Engineer');
  const operations = [() => store.saveTask({ ...task }), () => store.saveProgressMode(project.id, 'phases'), () => store.saveProject({ ...project })];
  for (const deniedBy of ['assignment', 'role']) {
    project.ownerKey = deniedBy === 'assignment' ? 'claims|other' : 'claims|engineer';
    state.users[0].role = deniedBy === 'role' ? 'Viewer' : 'Project Engineer';
    for (const operation of operations) await expect(operation()).rejects.toThrow(/Only managers/);
  }
  expect(raw.saveTask).not.toHaveBeenCalled(); expect(raw.saveProjectTracking).not.toHaveBeenCalled();
});

it('skips unchanged progress/tracking but keeps manager project permission repair callable', async () => {
  const { store, raw, project, state } = fixture('Project Engineer');
  expect(await store.saveProgressMode(project.id, 'phases')).toBe(project);
  expect(await store.saveProject({ ...project, title: 'Ignored definition' })).toBe(project);
  expect(raw.saveProjectTracking).not.toHaveBeenCalled();
  await expect(store.saveProgressMode(project.id, 'invalid')).rejects.toThrow('valid progress');
  await store.saveProgressMode(project.id, 'tasks');
  await store.saveProject({ ...project, health: 'At Risk' });
  expect(raw.saveProjectTracking).toHaveBeenCalledTimes(2);
  state.users[0].role = 'Manager';
  await store.saveProject({ ...project });
  expect(raw.saveProject).toHaveBeenCalledTimes(1);
});

it.each(['editUpdate', 'editRisk'])('%s preserves current identity/version for unchanged content and still writes changed content', async method => {
  const { store, raw, update, risk } = fixture('Project Engineer');
  const row = method === 'editUpdate' ? update : risk;
  const draft = method === 'editUpdate' ? { ...row, summary: ' Progress ', editedAt: 'forged' } : { ...row, title: ' Risk ', mitigation: null };
  expect(await store[method](draft)).toBe(row);
  expect(raw[method]).not.toHaveBeenCalled();
  const change = method === 'editUpdate' ? { summary: 'New progress' } : { mitigation: 'New mitigation' };
  await store[method]({ ...row, ...change });
  expect(raw[method]).toHaveBeenCalledTimes(1);
  if (method === 'editUpdate') expect(raw[method].mock.calls[0][0]).toMatchObject({ summary: 'New progress', authorName: row.authorName, editorName: 'Engineer' });
});

it.each(['editUpdate', 'editRisk'])('%s never lets a no-op bypass stale versions, revoked roles or assignments', async method => {
  const { store, raw, update, risk, state, project } = fixture('Project Engineer');
  const row = method === 'editUpdate' ? update : risk;
  for (const version of [undefined, '*', '"stale"']) await expect(store[method]({ ...row, version })).rejects.toThrow('changed');
  project.ownerKey = 'claims|other';
  await expect(store[method]({ ...row })).rejects.toThrow('Only managers');
  project.ownerKey = 'claims|engineer'; state.users[0].role = 'Viewer';
  await expect(store[method]({ ...row })).rejects.toThrow('Only managers');
  expect(raw[method]).not.toHaveBeenCalled();
});

it('normalizes reference edits, validates current destinations, and authorizes before skipping unchanged saves', async () => {
  const { store, raw, references, state } = fixture();
  const current = references[0];
  expect(await store.saveReferenceEntry({ ...current, name: ' Guide.pdf ', parentId: null })).toBe(current);
  expect(raw.saveReferenceEntry).not.toHaveBeenCalled();
  await expect(store.saveReferenceEntry({ ...current, parentId: 'missing' })).rejects.toThrow('existing folder');
  await store.saveReferenceEntry({ ...current, parentId: 'folder' });
  expect(raw.saveReferenceEntry).toHaveBeenCalledExactlyOnceWith({ ...current, parentId: 'folder' }, undefined);
  state.users[0].role = 'Viewer';
  await expect(store.saveReferenceEntry({ ...current })).rejects.toThrow('Only managers');
  expect(raw.saveReferenceEntry).toHaveBeenCalledTimes(1);
});
