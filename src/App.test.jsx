import { PeopleInvite } from './components/PeopleInvite';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { App, ProjectEditor, userInitials } from './App';
import { createRepository } from './lib/repository';

globalThis.IS_REACT_ACT_ENVIRONMENT = true;

describe('application shell', () => {
  let root;

  beforeEach(() => {
    localStorage.clear();
    document.documentElement.removeAttribute('data-theme');
    document.body.innerHTML = '<div id="root"></div>';
    window.MOD_TRACKER_CONFIG = { forceLocal: true };
    window.matchMedia = () => ({ matches: false, addEventListener() {}, removeEventListener() {} });
    window.confirm = vi.fn(() => true);
    window.prompt = vi.fn(() => 'Outside this project scope');
  });

  afterEach(async () => {
    if (root) await act(async () => root.unmount());
    vi.restoreAllMocks();
  });

  async function renderApp(seedManager = true) {
    if (seedManager) await createRepository().store.saveUser({ id: 'test-manager', title: 'Local Engineer', loginName: 'local', email: 'local.engineer@example.invalid', role: 'Manager' });
    await act(async () => {
      root = createRoot(document.getElementById('root'));
      root.render(<App />);
      await new Promise((resolve) => setTimeout(resolve, 30));
    });
  }

  function changeValue(element, value) {
    const prototype = element instanceof HTMLSelectElement ? HTMLSelectElement.prototype : element instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(prototype, 'value').set.call(element, value);
    element.dispatchEvent(new Event(element instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true }));
  }

  it('formats SharePoint display-name initials without rank or organization suffixes', () => {
    expect(userInitials({ title: 'Doe, Jordan T CIV (USA)' })).toBe('DJ');
    expect(userInitials({ title: 'Leila Engineer' })).toBe('LE');
  });

  it('hides the empty register and duplicate first-project action', async () => {
    await renderApp();

    expect(document.body.textContent).toContain('Modernization at a glance');
    expect(document.body.textContent).toContain('0 total measurement areas');
    expect(document.body.textContent).not.toContain('Your modernization portfolio is ready');
    expect(document.querySelector('.portfolio-register')).toBeNull();
    expect(document.querySelector('.topbar .search-box')).toBeNull();
    expect(document.querySelector('.brand-logo')?.getAttribute('alt')).toBe('NAVAIR');
    expect(document.querySelector('.sidebar-user .avatar')?.textContent).toBe('LE');
    expect(document.querySelector('.sync-card')).toBeNull();
    expect(document.querySelectorAll('.empty-actions button')).toHaveLength(0);
    expect(document.body.textContent).not.toContain('Demo workspace');
    expect(document.body.textContent).not.toContain('SharePoint workspace');
    expect(document.body.textContent).not.toContain('historical baseline');
    expect(document.body.textContent).not.toContain('Open sample portfolio');
    expect([...document.querySelectorAll('.phase-bar span')].every((bar) => bar.style.width === '0%')).toBe(true);

    const boardButton = [...document.querySelectorAll('button')].find((button) => button.textContent.includes('Pipeline board'));
    await act(async () => boardButton.click());
    expect(document.body.textContent).toContain('Scan where every project sits');
    expect(document.querySelectorAll('.kanban-column')).toHaveLength(4);

    const glossaryButton = [...document.querySelectorAll('.sidebar nav button')].find((button) => button.textContent.includes('Acronym glossary'));
    await act(async () => glossaryButton.click());
    expect(document.body.textContent).toContain('Calibration Standard Specification');
  });

  it('adds and removes shared glossary acronyms', async () => {
    await renderApp();
    const glossaryButton = [...document.querySelectorAll('.sidebar nav button')].find((button) => button.textContent.includes('Acronym glossary'));
    await act(async () => glossaryButton.click());

    await act(async () => {
      changeValue(document.querySelector('input[aria-label="Acronym"]'), 'abc');
      changeValue(document.querySelector('input[aria-label="Full term"]'), 'Added By Collaborator');
      changeValue(document.querySelector('input[aria-label="Definition"]'), 'A shared glossary entry.');
    });
    await act(async () => {
      document.querySelector('.glossary-composer').dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
      await Promise.resolve();
    });
    expect(document.body.textContent).toContain('Added By Collaborator');
    expect(document.querySelector('input[aria-label="Acronym"]').value).toBe('');

    await act(async () => {
      document.querySelector('button[aria-label="Remove ABC"]').click();
      await Promise.resolve();
    });
    expect(window.confirm).not.toHaveBeenCalled();
    expect(document.body.textContent).not.toContain('Added By Collaborator');
    expect(document.body.textContent).toContain('Calibration Standard Specification');
  });

  it.each(['click', 'Enter'])('lets a viewer add an acronym via %s in an embedded host that blocks form submission', async method => {
    await renderApp(false);
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find(button => button.textContent.includes('Acronym glossary')).click());
    await act(async () => {
      changeValue(document.querySelector('[aria-label="Acronym"]'), 'vt');
      changeValue(document.querySelector('[aria-label="Full term"]'), 'Viewer term');
    });
    const form = document.querySelector('.glossary-composer');
    form.addEventListener('submit', event => { event.preventDefault(); event.stopPropagation(); });
    await act(async () => {
      if (method === 'click') form.querySelector('button').click();
      else form.querySelector('input').dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
    });
    expect(document.querySelector('.glossary-table').textContent).toContain('Viewer term');
    expect(document.querySelector('.glossary-delete')).toBeNull();
    expect((await createRepository().store.load()).acronyms).toContainEqual(expect.objectContaining({ acronym: 'VT', term: 'Viewer term' }));
  });

  it('counts owned projects in My Work and assigns their starter tasks to the project owner', async () => {
    await renderApp();
    const newProject = [...document.querySelectorAll('button')].find((button) => button.textContent.includes('New project'));
    await act(async () => newProject.click());
    const nameInput = document.querySelector('.modal .field input');
    const assignMe = [...document.querySelectorAll('.modal button')].find((button) => button.textContent === 'Assign me');
    expect(document.querySelector('.modal input[type="checkbox"]')).toBeNull();
    await act(async () => { changeValue(nameInput, 'Owned modernization project'); assignMe.click(); });
    const saveProject = [...document.querySelectorAll('.modal button')].find((button) => button.textContent.includes('Save project'));
    await act(async () => {
      saveProject.click();
      await new Promise((resolve) => setTimeout(resolve, 50));
    });

    const myWork = [...document.querySelectorAll('.sidebar nav button')].find((button) => button.textContent.includes('My work'));
    expect(myWork.querySelector('.badge').textContent).toBe('1');
    await act(async () => myWork.click());
    expect(document.body.textContent).toContain('Owned modernization project');
    expect(document.body.textContent).toContain('Work breakdown (4)');
    expect((await createRepository().store.load()).projects[0]).toMatchObject({ ownerKey: 'local' });
  });

  it('retries only missing starter tasks after a partial project creation', async () => {
    const prototype = Object.getPrototypeOf(createRepository().store);
    const saveTask = prototype.saveTask;
    const writes = vi.spyOn(prototype, 'saveTask').mockRejectedValueOnce(new Error('Starter task could not be saved')).mockImplementation(saveTask);
    await renderApp();
    await act(async () => [...document.querySelectorAll('button')].find(button => button.textContent.includes('New project')).click());
    const name = document.querySelector('.modal .field input');
    await act(async () => changeValue(name, 'Partial project'));
    const save = () => [...document.querySelectorAll('.modal button')].find(button => button.textContent === 'Save project').click();
    await act(async () => save());
    expect(document.querySelector('.modal').textContent).toContain('Starter task could not be saved');
    expect(name.value).toBe('Partial project');
    const partial = await createRepository().store.load();
    expect(partial.projects).toHaveLength(1);
    expect(partial.tasks).toHaveLength(3);
    await act(async () => save());
    const complete = await createRepository().store.load();
    expect(complete.projects.map(row => row.id)).toEqual([partial.projects[0].id]);
    expect(new Set(complete.tasks.map(row => row.id)).size).toBe(4);
    expect(writes).toHaveBeenCalledTimes(5);
  });

  it('assigns a verified SharePoint user from partial-name suggestions when editing a project', async () => {
    const raw = createRepository().store;
    await raw.saveUser({ id: 'hank', title: 'Chi, Hank CIV', email: 'hank@example.test', loginName: 'claims|hank@example.test', role: 'Project Engineer' });
    await raw.saveProject({ id: 'owner-edit', projectKey: 'owner-edit', title: 'Owner selection', ownerName: 'Old Owner', ownerEmail: 'old@example.test', ownerKey: 'old' });
    await renderApp();
    await act(async () => document.querySelector('.project-table-row:not(.table-header)').click());
    await act(async () => document.querySelector('[aria-label="Project settings"]').click());
    await act(async () => [...document.querySelectorAll('button')].find((button) => button.textContent === 'Edit project').click());
    expect(document.querySelector('.modal').textContent).not.toContain('Assign project owner');
    const label = [...document.querySelectorAll('.modal label')].find(row => row.textContent === 'Project Engineer’s email');
    const owner = label.control;
    const save = () => [...document.querySelectorAll('.modal button')].find((button) => button.textContent === 'Save project');
    await act(async () => changeValue(owner, 'Hank'));
    expect(save().disabled).toBe(true);
    expect(document.querySelector('.modal output').textContent).not.toContain('Old Owner');
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 360)); });
    const option = document.querySelector('.modal [role="option"]');
    expect(option.textContent).toContain('Chi, Hank CIV');
    expect(option.textContent).toContain('hank@example.test');
    expect(save().disabled).toBe(true);
    await act(async () => option.click());
    expect(document.querySelector('.modal output').textContent).toBe('Chi, Hank CIV');
    expect(owner.value).toBe('hank@example.test');
    expect(save().disabled).toBe(false);
    expect(document.querySelector('.owner-edit-option')).toBeNull();
    await act(async () => save().click());
    expect((await createRepository().store.load()).projects[0]).toMatchObject({ ownerKey: 'claims|hank@example.test', ownerEmail: 'hank@example.test', ownerName: 'Chi, Hank CIV' });
  });

  it('searches partial email addresses and resolves only the keyboard-selected account before saving', async () => {
    vi.useFakeTimers();
    try {
      const first = { title: 'Alex River', loginName: 'claims|one@example.test', email: 'alex.river@example.test' };
      const second = { title: 'Alex Rivers', loginName: 'claims|two@example.test', email: 'alex.rivers@example.test' };
      let confirm;
      const store = {
        searchPeople: vi.fn().mockResolvedValue([first, second]),
        resolvePerson: vi.fn(() => new Promise(resolve => { confirm = resolve; })),
      };
      const onSave = vi.fn();
      const onClose = vi.fn();
      await act(async () => {
        root = createRoot(document.getElementById('root'));
        root.render(<ProjectEditor manager store={store} user={{}} onSave={onSave} onClose={onClose} />);
      });
      const email = document.querySelector('#project-engineer-email');
      const options = () => [...document.querySelectorAll('.modal [role="option"]')];
      const key = key => act(async () => email.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })));
      const save = [...document.querySelectorAll('.modal button')].find(button => button.textContent === 'Save project');
      await act(async () => {
        email.focus();
        changeValue(email, 'a');
        changeValue(document.querySelector('.modal .field input'), 'New assignment');
      });
      await act(async () => vi.advanceTimersByTimeAsync(300));
      expect(store.searchPeople).not.toHaveBeenCalled();
      await act(async () => changeValue(email, 'alex.riv'));
      await act(async () => vi.advanceTimersByTimeAsync(300));
      expect(store.searchPeople).toHaveBeenCalledWith('alex.riv');
      expect(options().map(option => option.textContent)).toEqual([expect.stringContaining(first.email), expect.stringContaining(second.email)]);
      await key('Escape');
      expect(options()).toHaveLength(0);
      expect(onClose).not.toHaveBeenCalled();
      expect(save.disabled).toBe(true);
      await act(async () => changeValue(email, 'alex.rive'));
      await act(async () => vi.advanceTimersByTimeAsync(300));
      await key('ArrowDown');
      await key('ArrowDown');
      await key('Enter');
      expect(store.resolvePerson).toHaveBeenCalledExactlyOnceWith(second.loginName);
      expect(save.disabled).toBe(true);
      const confirmed = { ...second, title: 'Rivers, Alex' };
      await act(async () => confirm(confirmed));
      expect(document.querySelector('.modal output').textContent).toBe(confirmed.title);
      expect(email.value).toBe(confirmed.email);
      expect(options()).toHaveLength(0);
      await act(async () => save.click());
      expect(onSave).toHaveBeenCalledExactlyOnceWith(expect.objectContaining({ ownerName: confirmed.title, ownerEmail: confirmed.email, ownerKey: confirmed.loginName }), confirmed.loginName);
    } finally { vi.useRealTimers(); }
  });

  it('shows suggestions for a partially typed domain and rejects a different resolved login with the same email', async () => {
    vi.useFakeTimers();
    try {
      const suggestion = { title: 'River Smith', email: 'river@example.com', loginName: 'claims|river-primary' };
      const store = {
        searchPeople: vi.fn().mockResolvedValue([suggestion]),
        resolvePerson: vi.fn().mockResolvedValue({ ...suggestion, loginName: 'claims|river-alternate' }),
      };
      const onSave = vi.fn();
      await act(async () => {
        root = createRoot(document.getElementById('root'));
        root.render(<ProjectEditor manager store={store} user={{}} project={{ id: 'p', title: 'Existing project' }} onClose={() => {}} onSave={onSave} />);
      });
      const email = document.querySelector('#project-engineer-email');
      await act(async () => { email.focus(); changeValue(email, 'river@example.co'); });
      await act(async () => vi.advanceTimersByTimeAsync(300));
      const option = document.querySelector('.modal [role="option"]');
      expect(option.textContent).toContain(suggestion.email);
      expect(store.resolvePerson).not.toHaveBeenCalled();
      await act(async () => option.click());
      expect(store.resolvePerson).toHaveBeenCalledExactlyOnceWith(suggestion.loginName);
      expect(document.querySelector('.modal [role="alert"]').textContent).toContain('does not match');
      const save = [...document.querySelectorAll('.modal button')].find(button => button.textContent === 'Save project');
      expect(save.disabled).toBe(true);
      await act(async () => save.click());
      expect(onSave).not.toHaveBeenCalled();
    } finally { vi.useRealTimers(); }
  });

  it('discards a stale engineer lookup and retries a new project with the same identity and IDs', async () => {
    vi.useFakeTimers();
    try {
      let finishFirstLookup;
      const firstLookup = new Promise(resolve => { finishFirstLookup = resolve; });
      const engineer = { title: 'Engineer Two', email: 'two@example.test', loginName: 'claims|two@example.test' };
      const store = {
        searchPeople: vi.fn(async email => email === engineer.email ? [] : firstLookup),
        resolvePerson: vi.fn(async login => login === engineer.email ? engineer : { title: 'Engineer One', email: 'one@example.test', loginName: 'claims|one@example.test' }),
      };
      const onSave = vi.fn().mockRejectedValueOnce(new Error('Permission update interrupted')).mockResolvedValueOnce({});
      await act(async () => {
        root = createRoot(document.getElementById('root'));
        root.render(<ProjectEditor manager store={store} user={{}} onClose={() => {}} onSave={onSave} />);
      });
      const email = document.querySelector('#project-engineer-email');
      const save = [...document.querySelectorAll('.modal button')].find(button => button.textContent === 'Save project');
      await act(async () => {
        changeValue(document.querySelector('.modal .field input'), 'Equipment upgrade');
        changeValue(email, 'one@example.test');
      });
      await act(async () => vi.advanceTimersByTimeAsync(300));
      await act(async () => changeValue(email, engineer.email));
      expect(save.disabled).toBe(true);
      await act(async () => vi.advanceTimersByTimeAsync(300));
      await act(async () => finishFirstLookup([{ title: 'Engineer One', email: 'one@example.test', loginName: 'claims|one@example.test' }]));
      expect(document.querySelector('.modal output').textContent).toBe(engineer.title);
      expect(onSave).not.toHaveBeenCalled();
      await act(async () => save.click());
      expect(document.querySelector('.modal [role="alert"]').textContent).toBe('Permission update interrupted');
      await act(async () => save.click());
      expect(onSave).toHaveBeenCalledTimes(2);
      expect(onSave.mock.calls[1]).toEqual(onSave.mock.calls[0]);
      expect(onSave.mock.calls[1]).toEqual([expect.objectContaining({
        id: expect.any(String), projectKey: expect.any(String), ownerKey: engineer.loginName, ownerName: engineer.title, ownerEmail: engineer.email,
      }), engineer.loginName]);
    } finally { vi.useRealTimers(); }
  });

  it('blocks ambiguous or mismatched full email addresses before saving an assignment', async () => {
    vi.useFakeTimers();
    try {
      const different = { title: 'Another account', email: 'other@example.test', loginName: 'claims|other@example.test' };
      const shared = { title: 'Shared', email: 'shared@example.test', loginName: 'one' };
      const store = {
        searchPeople: vi.fn(async email => email === shared.email ? [shared, { ...shared, loginName: 'two' }] : []),
        resolvePerson: vi.fn(async () => different),
      };
      const onSave = vi.fn();
      await act(async () => {
        root = createRoot(document.getElementById('root'));
        root.render(<ProjectEditor manager store={store} user={{}} project={{ id: 'p', title: 'Existing', ownerKey: 'original', ownerName: 'Original' }} onClose={() => {}} onSave={onSave} />);
      });
      const email = document.querySelector('#project-engineer-email');
      const save = [...document.querySelectorAll('.modal button')].find(button => button.textContent === 'Save project');
      await act(async () => changeValue(email, shared.email));
      await act(async () => vi.advanceTimersByTimeAsync(300));
      expect(document.querySelector('.modal [role="alert"]').textContent).toContain('More than one account');
      expect(store.resolvePerson).not.toHaveBeenCalled();
      expect(save.disabled).toBe(true);
      await act(async () => changeValue(email, 'wrong@example.test'));
      await act(async () => vi.advanceTimersByTimeAsync(300));
      expect(document.querySelector('.modal [role="alert"]').textContent).toContain('does not match');
      expect(save.disabled).toBe(true);
      expect(onSave).not.toHaveBeenCalled();
    } finally { vi.useRealTimers(); }
  });

  it('lets an assigned engineer save tracking fields while preserving the assignment and project definition', async () => {
    const raw = createRepository().store;
    await raw.saveUser({ id: 'engineer', title: 'Local Engineer', loginName: 'local', email: 'local.engineer@example.invalid', role: 'Project Engineer' });
    await raw.saveProject({ id: 'p', projectKey: 'p', title: 'Engineer tracking', ownerName: 'Legacy engineer', ownerKey: 'local', health: 'On Track', status: 'Active' });
    await raw.saveTask({ id: 't', projectKey: 'p', title: 'Buy standard', phaseKey: 'procurement', status: 'Not Started' });
    await renderApp(false);
    await act(async () => document.querySelector('.project-table-row:not(.table-header)').click());
    await act(async () => document.querySelector('[aria-label="Project settings"]').click());
    await act(async () => [...document.querySelectorAll('button')].find(button => button.textContent === 'Edit project').click());
    const field = label => [...document.querySelectorAll('.modal label')].find(el => el.firstElementChild?.textContent === label).querySelector('input,select,textarea');
    for (const label of ['Project name', 'Measurement area', 'Target completion', 'Priority', 'Description']) expect(field(label).disabled, label).toBe(true);
    expect(document.querySelector('#project-engineer-email').disabled).toBe(true);
    expect(document.querySelector('.modal output').textContent).toBe('Legacy engineer');
    expect(field('Current stage').value).toBe('');
    expect(field('Current stage').selectedOptions[0].textContent).toMatch(/Follow work breakdown.*Procurement/i);
    expect(field('Project status').value).toBe('Active');
    await act(async () => {
      changeValue(field('Health'), 'At Risk');
      changeValue(field('Current stage'), 'deployment');
      changeValue(field('Project status'), 'In Progress');
    });
    await act(async () => [...document.querySelectorAll('.modal button')].find(button => button.textContent === 'Save project').click());
    expect((await createRepository().store.load()).projects[0]).toMatchObject({
      title: 'Engineer tracking', ownerKey: 'local', ownerName: 'Legacy engineer', health: 'At Risk', currentStageKey: 'deployment', status: 'In Progress',
    });
    expect(document.querySelector('.project-drawer .progress-caption').textContent).toContain('Deployment');
    expect(document.querySelector('.drawer-badges').textContent).toContain('At Risk');
    await act(async () => document.querySelector('[aria-label="Project settings"]').click());
    await act(async () => [...document.querySelectorAll('button')].find(button => button.textContent === 'Edit project').click());
    await act(async () => changeValue(field('Current stage'), ''));
    await act(async () => [...document.querySelectorAll('.modal button')].find(button => button.textContent === 'Save project').click());
    expect(document.querySelector('.project-drawer .progress-caption').textContent).toMatch(/procurement/i);
  });

  it('lets managers edit tasks and delete projects without sample data', async () => {
    await renderApp();
    const newProject = [...document.querySelectorAll('button')].find((button) => button.textContent.includes('New project'));
    await act(async () => newProject.click());
    await act(async () => changeValue(document.querySelector('.modal .field input'), 'Task workflow project'));
    const saveProject = [...document.querySelectorAll('.modal button')].find((button) => button.textContent.includes('Save project'));
    await act(async () => {
      saveProject.click();
      await new Promise((resolve) => setTimeout(resolve, 50));
    });

    const tasksTab = [...document.querySelectorAll('.drawer-tabs button')].find((button) => button.textContent.includes('Work breakdown'));
    await act(async () => tasksTab.click());
    const phaseButton = document.querySelector('.phase-required-button:not(:disabled)');
    const phaseSection = phaseButton.closest('.task-stage');
    await act(async () => { phaseButton.click(); await Promise.resolve(); });
    await act(async () => changeValue(document.querySelector('.modal textarea'), 'Outside this project scope'));
    await act(async () => [...document.querySelectorAll('.modal button')].find((button) => button.textContent === 'Apply to phase').click());
    expect(window.prompt).not.toHaveBeenCalled();
    expect(phaseSection.classList.contains('phase-not-required')).toBe(true);
    expect(phaseSection.textContent).toContain('Restore phase');
    await act(async () => { phaseSection.querySelector('.phase-required-button').click(); await Promise.resolve(); });
    expect(phaseSection.classList.contains('phase-not-required')).toBe(false);
    const checkbox = document.querySelector('.task-check:not(.checked)');
    const wbs = checkbox.getAttribute('aria-label').replace('Complete ', '');
    await act(async () => { checkbox.click(); await Promise.resolve(); });
    expect(document.querySelector('.modal')).toBeNull();
    expect(document.querySelector(`[aria-label="Reopen ${wbs}"]`)).not.toBeNull();

    const taskRow = document.querySelector(`[aria-label="Reopen ${wbs}"]`).closest('.phase-task');
    await act(async () => taskRow.click());
    const statusSelect = [...document.querySelectorAll('.modal select')].find((select) => [...select.options].some((option) => option.textContent === 'Not Required'));
    expect(statusSelect).not.toBeUndefined();
    expect([...statusSelect.options].map((option) => option.textContent)).toContain('Not Required');
    const progressBeforeNotRequired = document.querySelector('.drawer-progress span').textContent;
    const dueDate = document.querySelector('.modal input[type="date"]');
    await act(async () => {
      changeValue(statusSelect, 'Not Required');
    });
    expect(dueDate.disabled).toBe(false);
    const justification = [...document.querySelectorAll('.modal .field')].find((field) => field.textContent.includes('Not required justification')).querySelector('textarea');
    await act(async () => changeValue(justification, 'Outside this project scope'));
    expect(dueDate.value).toBe('');
    const saveTaskButton = [...document.querySelectorAll('.modal .button.primary')].find((button) => button.textContent.includes('Save task'));
    await act(async () => { saveTaskButton.click(); await Promise.resolve(); });
    expect(document.querySelector(`[aria-label="Reopen ${wbs}"]`)).not.toBeNull();
    expect(document.querySelector('.drawer-progress span').textContent).toBe(progressBeforeNotRequired);

    await act(async () => document.querySelector('.project-drawer button[aria-label="Close"]').click());
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((button) => button.textContent.includes('Pipeline board')).click());
    const menuButton = document.querySelector('.project-card .project-menu');
    await act(async () => menuButton.click());
    const deleteButton = [...document.querySelectorAll('.project-menu-popover button')].find((button) => button.textContent.includes('Delete project'));
    await act(async () => { deleteButton.click(); await Promise.resolve(); });
    expect(window.confirm).not.toHaveBeenCalled();
    expect(document.querySelectorAll('.project-card')).toHaveLength(0);
    expect(document.querySelector('[role="status"]')?.textContent).toContain('deleted');
  });
  it('shows all portfolio rows and detailed overdue and exception records', async () => {
    const projects = Array.from({ length: 8 }, (_, i) => ({ id: `p${i}`, projectKey: `p${i}`, title: `Project ${i}`, ownerName: 'Engineer', health: 'On Track', status: i === 7 ? 'Complete' : 'In Progress', currentStageKey: 'requirement', percentComplete: 0 }));
    localStorage.setItem('modernization-project-tracker:v2', JSON.stringify({ projects, tasks: [
      { id: 't1', projectKey: 'p0', title: 'Late task', status: 'In Progress', dueDate: '2020-01-01', deferredDate: '2030-01-01', deferredJustification: 'Parts delayed' },
      { id: 't2', projectKey: 'p0', title: 'Omitted task', status: 'Not Required', notRequiredJustification: 'Scope approved' },
    ], updates: [], risks: [], acronyms: [] }));
    await renderApp();
    expect(document.querySelectorAll('.project-table-row:not(.table-header)')).toHaveLength(8);
    expect(document.querySelector('.attention-panel')).toBeNull();
    expect(document.body.textContent).not.toContain('All projects');
    await act(async () => [...document.querySelectorAll('.kpi-card')].find((card) => card.textContent.includes('NPSL (Calibration Lab)')).click());
    expect(document.querySelectorAll('.attention-section')).toHaveLength(1);
    expect(document.body.textContent).toContain('Overdue');
    expect(document.body.textContent).toContain('Parts delayed');
    expect(document.body.textContent).not.toContain('Scope approved');
    expect(document.body.textContent).not.toContain('Omitted task');
  });

  it('lets an explicitly assigned Viewer read all projects without editing', async () => {
    const raw = createRepository().store;
    vi.spyOn(Object.getPrototypeOf(raw), 'currentUser').mockResolvedValue({ loginName: 'engineer', title: 'Engineer' });
    await raw.saveUser({ id: 'viewer', loginName: 'engineer', title: 'Engineer', role: 'Viewer' });
    await raw.saveProject({ id: 'mine', projectKey: 'mine', title: 'Assigned project', ownerKey: 'engineer', ownerName: 'Engineer', health: 'On Track', status: 'In Progress' });
    await raw.saveProject({ id: 'other', projectKey: 'other', title: 'Private other project', ownerKey: 'other', ownerName: 'Other' });
    await raw.saveTask({ id: 'task', projectKey: 'mine', title: 'Assigned task', phaseKey: 'requirement', status: 'Not Started', dueDate: '2026-09-01', ownerKey: 'engineer' });
    await renderApp();
    expect(document.body.textContent).toContain('Private other project');
    expect(document.body.textContent).not.toContain('New project');
    expect(document.body.textContent).toContain('Users and managers');
    await act(async () => document.querySelector('.project-table-row:not(.table-header)').click());
    expect(document.querySelector('[aria-label="Project settings"]')).toBeNull();
    await act(async () => document.querySelector('.upcoming-list button').click());
    expect(document.querySelector('.modal input[type="date"]').disabled).toBe(true);
    expect(document.querySelector('.modal input[aria-label="Task name"]').matches(':disabled')).toBe(true);
    expect([...document.querySelectorAll('.modal select')].find((select) => [...select.options].some((option) => option.value === 'Complete')).matches(':disabled')).toBe(true);
    expect([...document.querySelectorAll('.modal button')].some((button) => button.textContent.includes('Delete task'))).toBe(false);
  });

  it('adds and deletes custom phase tasking without removing other tasks', async () => {
    const raw = createRepository().store;
    await raw.saveProject({ id: 'p', projectKey: 'p', title: 'Custom project', ownerName: 'Engineer', status: 'In Progress' });
    await raw.saveTask({ id: 'existing', projectKey: 'p', title: 'Keep this task', phaseKey: 'requirement', status: 'Not Started' });
    await renderApp();
    await act(async () => document.querySelector('.project-table-row:not(.table-header)').click());
    await act(async () => [...document.querySelectorAll('.drawer-tabs button')].find((button) => button.textContent.includes('Work breakdown')).click());
    await act(async () => [...document.querySelectorAll('.task-stage button')].find((button) => button.getAttribute('aria-label')?.startsWith('Add task to ')).click());
    await act(async () => changeValue(document.querySelector('[aria-label="Task name"]'), 'Custom task'));
    await act(async () => [...document.querySelectorAll('.modal button')].find((button) => button.textContent === 'Save task').click());
    expect(document.querySelectorAll('.phase-task')).toHaveLength(2);
    await act(async () => [...document.querySelectorAll('.phase-task')].find((row) => row.textContent.includes('Custom task')).click());
    await act(async () => [...document.querySelectorAll('.modal button')].find((button) => button.textContent === 'Delete task').click());
    expect(document.querySelectorAll('.phase-task')).toHaveLength(1);
    expect(document.querySelector('.phase-task').textContent).toContain('Keep this task');
    expect(JSON.parse(localStorage.getItem('modernization-project-tracker:v2')).tasks).toHaveLength(1);
  });

  it('queues documents before task creation and retries only failed files without creating a duplicate task', async () => {
    const raw = createRepository().store;
    await raw.saveProject({ id: 'p', projectKey: 'p', title: 'Upload project' });
    const prototype = Object.getPrototypeOf(raw);
    vi.spyOn(prototype, 'listTaskAttachments').mockResolvedValue([]);
    const upload = vi.spyOn(prototype, 'addTaskAttachment').mockResolvedValueOnce([]).mockRejectedValueOnce(new Error('Connection interrupted')).mockResolvedValue([]);
    await renderApp();
    await act(async () => document.querySelector('.project-table-row:not(.table-header)').click());
    await act(async () => [...document.querySelectorAll('.drawer-tabs button')].find(button => button.textContent.includes('Work breakdown')).click());
    await act(async () => document.querySelector('.add-task-button').click());
    await act(async () => changeValue(document.querySelector('[aria-label="Task name"]'), 'Queued task'));
    const input = document.querySelector('[aria-label="Attach documents"]');
    Object.defineProperty(input, 'files', { value: [new File(['one'], 'one.txt'), new File(['two'], 'two.txt')] });
    await act(async () => input.dispatchEvent(new Event('change', { bubbles: true })));
    expect(upload).not.toHaveBeenCalled();
    const save = () => [...document.querySelectorAll('.modal button')].find(button => button.textContent === 'Save task').click();
    await act(async () => save());
    expect(document.querySelector('.modal [role="alert"]').textContent).toContain('Connection interrupted');
    expect(document.querySelector('.queued-documents').textContent).toContain('two.txt');
    expect(document.querySelector('.queued-documents').textContent).not.toContain('one.txt');
    expect(JSON.parse(localStorage.getItem('modernization-project-tracker:v2')).tasks).toHaveLength(1);
    await act(async () => save());
    expect(document.querySelector('.modal')).toBeNull();
    expect(upload.mock.calls.map(([, file]) => file.name)).toEqual(['one.txt', 'two.txt', 'two.txt']);
    expect(new Set(upload.mock.calls.map(([task]) => task.id)).size).toBe(1);
    expect(JSON.parse(localStorage.getItem('modernization-project-tracker:v2')).tasks).toHaveLength(1);
  });

  it('registers standard users and grants testing manager access only after the right password', async () => {
    await renderApp(false);
    expect([...document.querySelectorAll('.sidebar nav button')].map((row) => row.textContent)).toEqual(['Portfolio', 'Pipeline board', 'My work0', 'Reference Documents', 'Acronym glossary', 'Users and managers']);
    expect(JSON.parse(localStorage.getItem('modernization-project-tracker:v2')).users[0]).toMatchObject({ title: 'Local Engineer', role: 'Viewer' });
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((row) => row.textContent === 'Users and managers').click());
    expect(document.querySelectorAll('.directory-row')).toHaveLength(1);
    expect(document.querySelector('.directory-row button')).toBeNull();
    expect(document.querySelector('input[type="password"]')).toBeNull();
    await act(async () => window.dispatchEvent(new KeyboardEvent('keydown', { key: 'm', ctrlKey: true, bubbles: true, cancelable: true })));
    await act(async () => changeValue(document.querySelector('input[type="password"]'), 'wrong'));
    await act(async () => [...document.querySelectorAll('.directory-form button')].find((button) => button.textContent === 'Enable manager access').click());
    expect(document.querySelector('[role="alert"]').textContent).toContain('incorrect');
    expect(document.body.textContent).not.toContain('New project');
    await act(async () => changeValue(document.querySelector('input[type="password"]'), 'admin123'));
    await act(async () => [...document.querySelectorAll('.directory-form button')].find((button) => button.textContent === 'Enable manager access').click());
    expect(document.body.textContent).toContain('New project');
    expect(document.querySelector('input[type="password"]')).toBeNull();
    expect(JSON.parse(localStorage.getItem('modernization-project-tracker:v2')).users[0].role).toBe('Manager');
  });

  it('combines stage multi-selection with local portfolio search and health filters', async () => {
    const raw = createRepository().store;
    for (const [id, stage, health] of [['Alpha', 'requirement', 'On Track'], ['Beta', 'acquisition', 'At Risk'], ['Gamma', 'procurement', 'On Track']]) await raw.saveProject({ id, projectKey: id, title: id, ownerName: 'Engineer', currentStageKey: stage, health, status: 'Planned' });
    await renderApp();
    const rows = () => [...document.querySelectorAll('.project-table-row:not(.table-header)')].map((row) => row.textContent);
    const stages = document.querySelectorAll('.phase-node');
    await act(async () => stages[0].click());
    expect(rows()).toHaveLength(1);
    await act(async () => stages[1].dispatchEvent(new MouseEvent('click', { bubbles: true, ctrlKey: true })));
    expect(rows()).toHaveLength(2);
    await act(async () => changeValue(document.querySelector('[aria-label="Search portfolio projects"]'), 'Beta'));
    expect(rows()).toHaveLength(1);
    expect(rows()[0]).toContain('Beta');
    expect(document.querySelectorAll('.phase-node.selected')).toHaveLength(2);
    await act(async () => changeValue(document.querySelector('[aria-label="Filter by health"]'), 'On Track'));
    expect(rows()).toHaveLength(0);
    await act(async () => { changeValue(document.querySelector('[aria-label="Search portfolio projects"]'), ''); changeValue(document.querySelector('[aria-label="Filter by health"]'), ''); });
    await act(async () => [...document.querySelectorAll('button')].find((row) => row.textContent === 'Clear filter').click());
    expect(rows()).toHaveLength(3);
    expect(document.querySelector('.topbar input')).toBeNull();
  });

  it('filters by organization and sorts project names', async () => {
    const raw = createRepository().store;
    await raw.saveProject({ id: 'zulu', projectKey: 'zulu', title: 'Zulu project', organization: 'NPSL', ownerName: 'Zulu Owner', currentStageKey: 'procurement', health: 'At Risk' });
    await raw.saveProject({ id: 'alpha', projectKey: 'alpha', title: 'Alpha project', organization: 'ISE', ownerName: 'Alpha Owner', currentStageKey: 'requirement', health: 'On Track' });
    await raw.saveTask({ id: 'z-task', projectKey: 'zulu', title: 'Z task', organization: 'NPSL', status: 'Not Started' });
    await raw.saveTask({ id: 'a-task', projectKey: 'alpha', title: 'A task', organization: 'ISE', status: 'Not Started' });
    await renderApp();
    const rows = () => [...document.querySelectorAll('.project-table-row:not(.table-header)')].map((row) => row.textContent);
    expect(rows()[0]).toContain('Alpha project');
    await act(async () => [...document.querySelectorAll('.table-sort')].find((button) => button.textContent.includes('Project')).click());
    expect(rows()[0]).toContain('Zulu project');
    await act(async () => changeValue(document.querySelector('[aria-label="Filter by organization"]'), 'NPSL Metrology Engineering Division'));
    expect(rows()).toHaveLength(1);
    expect(rows()[0]).toContain('Alpha project');
    expect(rows()[0]).toContain('NPSL Metrology Engineering');
  });

  it('sorts all portfolio data columns in both directions', async () => {
    const raw = createRepository().store;
    await raw.saveProject({ id: 'a', projectKey: 'a', title: 'Alpha', ownerName: 'Zulu', organization: 'Program Office', currentStageKey: 'procurement', health: 'On Track' });
    await raw.saveProject({ id: 'z', projectKey: 'z', title: 'Zulu', ownerName: 'Alpha', organization: 'ISE', currentStageKey: 'requirement', health: 'Blocked' });
    await raw.saveTask({ id: 'ad', projectKey: 'a', title: 'Done', phaseKey: 'requirement', status: 'Complete' });
    await raw.saveTask({ id: 'an', projectKey: 'a', title: 'Alpha milestone', phaseKey: 'procurement', status: 'Not Started' });
    await raw.saveTask({ id: 'zn', projectKey: 'z', title: 'Zulu milestone', phaseKey: 'requirement', status: 'Not Started' });
    await renderApp();
    const first = () => document.querySelector('.project-table-row:not(.table-header) > span > strong').textContent;
    for (const [label, ascending] of [['Project engineer', 'Zulu'], ['Organization', 'Alpha'], ['Stage', 'Zulu'], ['Health', 'Alpha'], ['Progress', 'Zulu'], ['Next milestone', 'Alpha'], ['Project', 'Alpha']]) {
      const heading = () => [...document.querySelectorAll('.table-sort')].find((button) => button.textContent.startsWith(label));
      await act(async () => heading().click());
      expect(first(), `${label} ascending`).toBe(ascending);
      await act(async () => heading().click());
      expect(first(), `${label} descending`).toBe(ascending === 'Alpha' ? 'Zulu' : 'Alpha');
    }
  });

  it('scopes each attention card to its organization and sorts board cards', async () => {
    const raw = createRepository().store;
    for (const [id, organization] of [['p', 'CHENG Team'], ['n', 'NPSL'], ['i', 'NPSL Metrology Engineering']]) {
      await raw.saveProject({ id, projectKey: id, title: `${organization} project`, organization, currentStageKey: 'requirement' });
      await raw.saveTask({ id: `${id}-t`, projectKey: id, title: `${organization} task`, phaseKey: 'requirement', status: 'Not Started', ...(id === 'n' ? { organization: 'NPSL Metrology Calibration' } : {}) });
    }
    await renderApp();
    expect(document.querySelector('.kpi-grid').textContent).not.toContain('Portfolio Progress');
    for (const [organization, label] of [['CHENG Team', 'NAVAIR METCAL 260B Needs Attention'], ['NPSL', 'NPSL (Calibration Lab)'], ['NPSL Metrology Engineering', 'NPSL Metrology Engineering Division']]) {
      await act(async () => [...document.querySelectorAll('.kpi-card')].find((card) => card.textContent.includes(label)).click());
      expect(document.querySelector('h1').textContent).toBe(label);
      expect(document.querySelector('.attention-groups').textContent).toContain(`${organization} project`);
      for (const other of ['CHENG Team', 'NPSL', 'NPSL Metrology Engineering'].filter((value) => value !== organization)) expect(document.querySelector('.attention-groups').textContent).not.toContain(`${other} project`);
      await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((button) => button.textContent === 'Portfolio').click());
    }
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((button) => button.textContent === 'Pipeline board').click());
    await act(async () => changeValue(document.querySelector('[aria-label="Sort pipeline projects"]'), 'organization'));
    expect([...document.querySelectorAll('.kanban-board .project-card h3')].map((el) => el.textContent)).toEqual(['CHENG Team project', 'NPSL project', 'NPSL Metrology Engineering project']);
  });

  it('places project documents after upcoming work and names files in the phase indicator', async () => {
    const raw = createRepository().store;
    await raw.saveProject({ id: 'p', projectKey: 'p', title: 'Document project', currentStageKey: 'requirement' });
    await raw.saveTask({ id: 't', projectKey: 'p', title: 'Review plan', phaseKey: 'requirement', status: 'Not Started' });
    const attached = [{ name: 'Plan.pdf' }, { name: 'Evidence.docx' }];
    vi.spyOn(Object.getPrototypeOf(raw), 'listTaskAttachments').mockImplementation(async () => attached);
    vi.spyOn(Object.getPrototypeOf(raw), 'renameTaskAttachment').mockImplementation(async (_task, name, nextName) => { attached.find((file) => file.name === name).name = nextName; return attached; });
    await renderApp();
    await act(async () => document.querySelector('.project-table-row:not(.table-header)').click());
    const headings = [...document.querySelectorAll('.drawer-section-stack h3')].map((el) => el.textContent);
    expect(headings.indexOf('Documents')).toBeGreaterThan(headings.indexOf('Upcoming work'));
    await act(async () => document.querySelector('[aria-label="Edit Plan.pdf"]').click());
    expect([...document.querySelectorAll('.reference-action-menu button')].map((el) => el.textContent)).toEqual(['Download', 'Rename', 'Delete']);
    await act(async () => [...document.querySelectorAll('.reference-action-menu button')].find((button) => button.textContent === 'Rename').click());
    await act(async () => changeValue(document.querySelector('.document-rename input'), 'Updated Plan.pdf'));
    await act(async () => [...document.querySelectorAll('.document-rename button')].find((button) => button.textContent === 'Save').click());
    await act(async () => [...document.querySelectorAll('.drawer-tabs button')].find((button) => button.textContent.includes('Work breakdown')).click());
    const indicator = document.querySelector('.phase-document-indicator');
    expect(indicator.title).toContain('Review plan: Updated Plan.pdf');
    expect(indicator.title).toContain('Review plan: Evidence.docx');
    expect(indicator.nextElementSibling.classList.contains('badge')).toBe(true);
  });

  it('shows a persistent activation error when the directory write fails', async () => {
    await renderApp(false);
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((row) => row.textContent === 'Users and managers').click());
    await act(async () => window.dispatchEvent(new KeyboardEvent('keydown', { key: 'm', ctrlKey: true, bubbles: true, cancelable: true })));
    vi.spyOn(Object.getPrototypeOf(createRepository().store), 'saveUser').mockRejectedValue(new Error('SharePoint denied the directory update (403).'));
    await act(async () => changeValue(document.querySelector('input[type="password"]'), 'admin123'));
    await act(async () => [...document.querySelectorAll('.directory-form button')].find((row) => row.textContent === 'Enable manager access').click());
    expect(document.querySelector('[role="alert"]').textContent).toContain('403');
    expect(document.body.textContent).not.toContain('New project');
    expect(JSON.parse(localStorage.getItem('modernization-project-tracker:v2')).users[0].role).toBe('Viewer');
    expect(document.querySelector('.directory-form button').disabled).toBe(false);
  });

  it('lets an assigned engineer update existing work and progress while managers control task definitions', async () => {
    const raw = createRepository().store;
    await raw.saveUser({ id: 'engineer', loginName: 'local', title: 'Engineer', role: 'Project Engineer' });
    await raw.saveProject({ id: 'owner-project', projectKey: 'owner-project', title: 'Owner project', ownerKey: 'local', ownerName: 'Local Engineer', ownerCanEdit: true, status: 'In Progress' });
    await raw.saveTask({ id: 'done', projectKey: 'owner-project', title: 'Done task', phaseKey: 'requirement', status: 'Complete' });
    await raw.saveTask({ id: 'review', projectKey: 'owner-project', title: 'Review requirements', phaseKey: 'requirement', status: 'Not Started', dueDate: '2026-12-01', assignedDate: '2026-11-01' });
    await renderApp(false);
    await act(async () => document.querySelector('.project-table-row:not(.table-header)').click());
    await act(async () => changeValue(document.querySelector('[aria-label="Progress calculation"]'), 'tasks'));
    expect(document.querySelector('.drawer-progress').textContent).toContain('1 of 2 tasks completed');
    await act(async () => [...document.querySelectorAll('.drawer-tabs button')].find((row) => row.textContent.includes('Work breakdown')).click());
    expect(document.querySelector('.add-task-button')).toBeNull();
    await act(async () => [...document.querySelectorAll('.phase-task')].find(row => row.textContent.includes('Review requirements')).click());
    for (const label of ['Task name', 'Original due date', 'Assigned / creation date']) {
      expect(document.querySelector(`[aria-label="${label}"]`).disabled).toBe(true);
    }
    expect(document.querySelector('.modal').textContent).not.toContain('Delete task');
    await act(async () => changeValue(document.querySelector('[aria-label="Notes"]'), 'Reviewed with calibration lab'));
    await act(async () => [...document.querySelectorAll('.modal button')].find((row) => row.textContent === 'Save task').click());
    await act(async () => document.querySelector('[aria-label="Complete Review requirements"]').click());
    expect(document.querySelector('.drawer-progress').textContent).toContain('2 of 2 tasks completed');
    const loaded = await createRepository().store.load();
    expect(loaded.projects[0].progressMode).toBe('tasks');
    expect(loaded.tasks.find(row => row.id === 'review')).toMatchObject({ title: 'Review requirements', status: 'Complete', notes: 'Reviewed with calibration lab', dueDate: '2026-12-01', assignedDate: '2026-11-01' });
  });

  it('normalizes legacy program-office work, preserves project health, and switches task markers and portfolio progress', async () => {
    const raw = createRepository().store;
    await raw.saveProject({ id: 'p', projectKey: 'p', title: 'Office handoff', ownerName: 'Engineer', health: 'On Track', status: 'In Progress' });
    for (const [id, status] of [['done', 'Complete'], ['office', 'In Progress – At Program Office'], ['skip', 'Not Required']]) await raw.saveTask({ id, projectKey: 'p', title: `${id} task`, phaseKey: 'requirement', order: id === 'done' ? 1 : 2, status });
    await renderApp();
    expect(document.querySelector('.project-table-row:not(.table-header)').textContent).toContain('On Track');
    await act(async () => changeValue(document.querySelector('[aria-label="Portfolio progress calculation"]'), 'tasks'));
    expect(document.querySelector('.project-table-row:not(.table-header)').textContent).toContain('33%');
    await act(async () => document.querySelector('.project-table-row:not(.table-header)').click());
    await act(async () => changeValue(document.querySelector('[aria-label="Progress calculation"]'), 'tasks'));
    expect(document.querySelectorAll('.task-markers > span')).toHaveLength(3);
    expect(document.querySelectorAll('.task-markers .done')).toHaveLength(1);
    expect(document.querySelector('.progress-caption strong').textContent).toBe('office task');
    expect(document.querySelector('.task-markers [aria-current="step"]').title).toContain('office task');
    expect(document.body.textContent).not.toContain('Claim project');
    await act(async () => document.querySelector('[aria-label="Project settings"]').click());
    expect(document.querySelector('.drawer-header-actions .project-menu-popover').textContent).toBe('Edit projectDelete project');
    await act(async () => document.querySelector('.project-drawer [aria-label="Close"]').click());
    await act(async () => [...document.querySelectorAll('.kpi-card')].find((card) => card.textContent.includes('NPSL (Calibration Lab)')).click());
    expect(document.querySelectorAll('.attention-section')).toHaveLength(1);
    expect(document.querySelector('.attention-section summary').textContent).toContain('In Progress');
    expect(document.querySelector('.attention-section').textContent).not.toContain('skip task');
    expect(document.querySelector('.attention-section').open).toBe(true);
  });

  it('navigates from the active card and brand, and dismisses both project menus', async () => {
    await createRepository().store.saveProject({ id: 'p', projectKey: 'p', title: 'Navigation project', ownerName: 'Engineer' });
    await renderApp();
    await act(async () => [...document.querySelectorAll('.kpi-card')].find((row) => row.textContent.includes('Active projects')).click());
    expect(document.querySelector('.kanban-board')).not.toBeNull();
    await act(async () => document.querySelector('.project-menu').click());
    expect(document.querySelector('.project-menu-popover')).not.toBeNull();
    await act(async () => document.querySelector('.page-heading').dispatchEvent(new Event('pointerdown', { bubbles: true })));
    expect(document.querySelector('.project-menu-popover')).toBeNull();
    await act(async () => document.querySelector('.project-card').click());
    await act(async () => document.querySelector('[aria-label="Project settings"]').click());
    await act(async () => document.querySelector('.drawer-progress').dispatchEvent(new Event('pointerdown', { bubbles: true })));
    expect(document.querySelector('.project-menu-popover')).toBeNull();
    await act(async () => document.querySelector('[aria-label="Project settings"]').click());
    await act(async () => document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })));
    expect(document.querySelector('.project-menu-popover')).toBeNull();
    await act(async () => document.querySelector('.project-drawer [aria-label="Close"]').click());
    await act(async () => document.querySelector('.brand').click());
    expect(document.body.textContent).toContain('Modernization at a glance');
  });

  it.each(['Manager', 'Project Engineer', 'Viewer'])('saves %s through Edit for a user without email and preserves SharePoint identity', async (role) => {
    await createRepository().store.saveUser({ id: 'other', title: 'Other Engineer', email: '', loginName: 'i:0#.w|domain\\engineer', role: role === 'Viewer' ? 'Manager' : 'Viewer' });
    await renderApp();
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((row) => row.textContent === 'Users and managers').click());
    expect(document.body.textContent).not.toContain('Make manager');
    const card = [...document.querySelectorAll('.directory-row')].find((row) => row.textContent.includes('Other Engineer'));
    await act(async () => card.querySelector('button').click());
    expect(document.querySelector('.directory-form')).toBeNull();
    await act(async () => [...card.querySelectorAll('button')].find((button) => button.textContent === 'Update User').click());
    await act(async () => changeValue(document.querySelector('.directory-form select'), role));
    const form = document.querySelector('.directory-form');
    expect(form.checkValidity()).toBe(true);
    expect(form.querySelector('button').disabled).toBe(false);
    const blockNativeSubmit = (event) => event.preventDefault();
    form.addEventListener('click', blockNativeSubmit, true);
    await act(async () => form.querySelector('button').click());
    form.removeEventListener('click', blockNativeSubmit, true);
    const saved = (await createRepository().store.load()).users.find((row) => row.id === 'other');
    expect(saved).toMatchObject({ role, email: '', loginName: 'i:0#.w|domain\\engineer' });
    expect(card.textContent).toContain(role);
    expect(document.querySelector('.page-stack > [role="status"]').textContent).toContain(`saved as ${role}`);
  });

  it('keeps role edit errors visible and preserves the draft for retry', async () => {
    await createRepository().store.saveUser({ id: 'other', title: 'Other Engineer', email: 'other@example.test', loginName: 'other', role: 'Viewer' });
    await renderApp();
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((row) => row.textContent === 'Users and managers').click());
    const card = [...document.querySelectorAll('.directory-row')].find((row) => row.textContent.includes('Other Engineer'));
    await act(async () => card.querySelector('button').click());
    expect(document.querySelector('.directory-form')).toBeNull();
    await act(async () => [...card.querySelectorAll('button')].find((button) => button.textContent === 'Update User').click());
    await act(async () => changeValue(document.querySelector('.directory-form select'), 'Manager'));
    vi.spyOn(Object.getPrototypeOf(createRepository().store), 'saveUser').mockRejectedValue(new Error('SharePoint denied the directory update (403).'));
    await act(async () => document.querySelector('.directory-form button').click());
    expect(document.querySelector('.page-stack > [role="alert"]').textContent).toContain('403');
    expect(document.querySelector('.directory-form select').value).toBe('Manager');
    expect(card.textContent).toContain('Viewer');
    expect(document.querySelector('.directory-form button').disabled).toBe(false);
  });

  it('collapses project groups independently and saves the assigned date', async () => {
    const raw = createRepository().store;
    for (const id of ['alpha', 'beta']) {
      await raw.saveProject({ id, projectKey: id, title: id, ownerName: 'Engineer' });
      await raw.saveTask({ id: `${id}-task`, projectKey: id, title: `${id} work`, phaseKey: 'requirement', status: 'In Progress' });
    }
    await renderApp();
    await act(async () => [...document.querySelectorAll('.kpi-card')].find((row) => row.textContent.includes('NPSL (Calibration Lab)')).click());
    await act(async () => [...document.querySelectorAll('.attention-group-toggle button')].find((button) => button.textContent === 'Projects').click());
    const groups = document.querySelectorAll('.attention-section');
    expect(groups).toHaveLength(2);
    groups[0].open = false;
    expect(groups[1].open).toBe(true);
    await act(async () => groups[1].querySelector('.attention-task-title').click());
    const field = [...document.querySelectorAll('.modal .field')].find((field) => field.textContent.includes('Assigned / creation date')).querySelector('input');
    expect(field.value).toBe('');
    await act(async () => changeValue(field, '2026-09-15'));
    await act(async () => [...document.querySelectorAll('.modal button')].find((row) => row.textContent === 'Save task').click());
    expect((await createRepository().store.load()).tasks.find((row) => row.id === 'beta-task').assignedDate).toBe('2026-09-15');
  });

  it('converts legacy SME roles to scoped standard users', async () => {
    const raw = createRepository().store;
    await raw.saveUser({ id: 'sme', title: 'Local Engineer', loginName: 'local', role: 'SME' });
    await raw.saveProject({ id: 'outside', projectKey: 'outside', title: 'Other project', ownerKey: 'other', ownerName: 'Other Engineer' });
    await raw.saveTask({ id: 'outside-task', projectKey: 'outside', title: 'Other task', status: 'Not Started', phaseKey: 'requirement', estimatedHours: 4.5 });
    await renderApp(false);
    expect(document.body.textContent).toContain('My work');
    expect(document.body.textContent).toContain('Other project');
    expect(document.body.textContent).not.toContain('New project');
    expect((await raw.load()).users[0].role).toBe('Viewer');
  });

  it('reinvites the selected identity and clears the previous directory removal message', async () => {
    const prototype = Object.getPrototypeOf(createRepository().store);
    const candidate = { title: 'New Engineer', email: 'new@example.test', loginName: 'i:0#.f|membership|new@example.test' };
    vi.spyOn(prototype, 'searchPeople').mockResolvedValue([candidate]);
    const resolve = vi.spyOn(prototype, 'resolvePerson').mockResolvedValue(candidate);
    window.MOD_TRACKER_CONFIG.appUrl = 'https://tenant.sharepoint-mil.us/sites/mod/SitePages/Tracker.aspx';
    await createRepository().store.saveUser({ ...candidate, id: 'former-manager', role: 'Manager' });
    await renderApp();
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((row) => row.textContent === 'Users and managers').click());
    const previous = [...document.querySelectorAll('.directory-row')].find(row => row.textContent.includes(candidate.email));
    await act(async () => previous.querySelector('button').click());
    await act(async () => [...previous.querySelectorAll('button')].find(button => button.textContent === 'Delete User').click());
    expect(document.body.textContent).toContain('removed from the tracker directory');
    await act(async () => changeValue(document.querySelector('.people-invite input[type="search"]'), 'New Engineer'));
    await act(async () => new Promise((done) => setTimeout(done, 400)));
    await act(async () => document.querySelector('.people-results button').click());
    await act(async () => changeValue(document.querySelector('.people-invite select'), 'Viewer'));
    await act(async () => [...document.querySelectorAll('.people-invite button')].find((button) => button.textContent === 'Add to tracker').click());
    expect(resolve).toHaveBeenCalledWith(candidate.loginName);
    const saved = (await createRepository().store.load()).users.find((row) => row.loginName === candidate.loginName);
    expect(saved).toMatchObject({ ...candidate, role: 'Viewer' });
    expect(document.body.textContent).not.toContain('removed from the tracker directory');
    expect(document.querySelector('.people-invite input[type="url"]')).toBeNull();
    expect(document.querySelector('.invitation-result')).toBeNull();
  });

  it('retries failed site sharing without duplicating a saved user', async () => {
    const person = { title: 'Invite Engineer', email: 'invite@example.test', loginName: 'claims|invite' };
    const store = { searchPeople: vi.fn(async () => [person]), resolvePerson: vi.fn(async () => person), shareSiteAccess: vi.fn().mockRejectedValueOnce(new Error('Site sharing denied')).mockResolvedValueOnce({ access: 'Read', emailRequested: true }) };
    const save = vi.fn(async (row) => ({ ...row, id: 'invited' }));
    await act(async () => { root = createRoot(document.getElementById('root')); root.render(<PeopleInvite store={store} config={{ appUrl: 'https://tenant.sharepoint.com/sites/mod/app.aspx' }} onSave={save} localPreview={false} />); });
    await act(async () => changeValue(document.querySelector('input[type="search"]'), 'Invite'));
    await act(async () => new Promise((done) => setTimeout(done, 400)));
    await act(async () => document.querySelector('.people-results button').click());
    await act(async () => changeValue(document.querySelector('select'), 'Viewer'));
    await act(async () => [...document.querySelectorAll('button')].find((button) => button.textContent === 'Invite and grant access').click());
    expect(document.querySelector('[role="alert"]').textContent).toContain('Site sharing denied');
    await act(async () => [...document.querySelectorAll('button')].find((button) => button.textContent === 'Retry invitation').click());
    expect(save).toHaveBeenCalledTimes(1);
    expect(store.shareSiteAccess).toHaveBeenCalledTimes(2);
    expect(document.querySelector('.invitation-result')).toBeNull();
    expect(document.querySelector('[role="alert"]')).toBeNull();
    expect([...document.querySelectorAll('button')].find((button) => button.textContent === 'Invitation requested').disabled).toBe(true);
    expect(document.querySelector('[role="status"]').textContent).toContain('SharePoint accepted the invitation email request for Invite Engineer');
    expect(store.shareSiteAccess).toHaveBeenLastCalledWith(expect.objectContaining({ loginName: person.loginName }), 'Viewer', 'https://tenant.sharepoint.com/sites/mod/app.aspx');
  });

  it('creates nested references and moves a folder with persistence', async () => {
    await renderApp();
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((row) => row.textContent === 'Reference Documents').click());
    const click = async (text) => act(async () => [...document.querySelectorAll('.reference-library button')].find((row) => row.textContent === text).click());
    await click('New folder');
    expect(document.querySelector('.reference-editor select')).toBeNull();
    await act(async () => changeValue(document.querySelector('.reference-editor input'), 'Templates'));
    await click('Create folder');
    await click('Templates');
    await click('New folder');
    expect(document.querySelector('.reference-editor select')).toBeNull();
    await act(async () => changeValue(document.querySelector('.reference-editor input'), 'Acquisition'));
    await click('Create folder');
    let rows = await createRepository().store.listReferenceEntries();
    const parent = rows.find((row) => row.name === 'Templates');
    expect(rows.find((row) => row.name === 'Acquisition').parentId).toBe(parent.id);
    await act(async () => document.querySelector('[aria-label="Edit Acquisition"]').click());
    await click('Move');
    await act(async () => changeValue(document.querySelector('.reference-editor select'), ''));
    await click('Save reference');
    rows = await createRepository().store.listReferenceEntries();
    expect(rows.find((row) => row.name === 'Acquisition').parentId).toBe('');
  });

  it('deletes a user only from Edit, preserves assignments, and hides self-deletion', async () => {
    const raw = createRepository().store;
    await raw.saveUser({ id: 'other', title: 'Other Engineer', email: 'other@example.test', loginName: 'other', role: 'Viewer' });
    await raw.saveTask({ id: 'assigned', projectKey: 'project', title: 'Assigned task', ownerKey: 'other' });
    await renderApp();
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((button) => button.textContent === 'Users and managers').click());
    expect(document.body.textContent).not.toContain('Delete user');
    const card = [...document.querySelectorAll('.directory-row')].find((row) => row.textContent.includes('Other Engineer'));
    await act(async () => card.querySelector('button').click());
    await act(async () => [...card.querySelectorAll('button')].find((button) => button.textContent === 'Delete User').click());
    const loaded = await createRepository().store.load();
    expect(loaded.users.some((row) => row.id === 'other')).toBe(false);
    expect(loaded.tasks.find((row) => row.id === 'assigned').ownerKey).toBe('other');
    expect(document.querySelector('.page-stack > [role="status"]').textContent).toContain('removed');
    expect(window.confirm).not.toHaveBeenCalled();
    await act(async () => document.querySelector('.directory-row button').click());
    expect([...document.querySelectorAll('.directory-row button')].find((button) => button.textContent === 'Delete User').disabled).toBe(true);
    expect(document.querySelector('.directory-form')).toBeNull();
  });

});
