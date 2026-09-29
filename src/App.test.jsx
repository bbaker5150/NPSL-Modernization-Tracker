import { PeopleInvite } from './components/PeopleInvite';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { App, userInitials } from './App';
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

  it('counts owned projects in My Work and assigns their starter tasks to the project owner', async () => {
    await renderApp();
    const newProject = [...document.querySelectorAll('button')].find((button) => button.textContent.includes('New project'));
    await act(async () => newProject.click());
    const nameInput = document.querySelector('.modal .field input');
    const assignMe = [...document.querySelectorAll('.modal button')].find((button) => button.textContent === 'Assign me');
    await act(async () => {
      changeValue(nameInput, 'Owned modernization project');
      assignMe.click();
    });
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
    await act(async () => [...document.querySelectorAll('.kpi-card')].find((card) => card.textContent.includes('Needs attention')).click());
    expect(document.querySelectorAll('.attention-group')).toHaveLength(1);
    await act(async () => document.querySelector('.attention-collapse').click());
    expect(document.body.textContent).toContain('Overdue');
    expect(document.body.textContent).toContain('Parts delayed');
    expect(document.body.textContent).not.toContain('Scope approved');
    expect(document.body.textContent).not.toContain('Omitted task');
  });

  it('limits standard users to assigned work and keeps deadline controls read-only', async () => {
    const raw = createRepository().store;
    vi.spyOn(Object.getPrototypeOf(raw), 'currentUser').mockResolvedValue({ loginName: 'engineer', title: 'Engineer' });
    await raw.saveProject({ id: 'mine', projectKey: 'mine', title: 'Assigned project', ownerKey: 'engineer', ownerName: 'Engineer', health: 'On Track', status: 'In Progress' });
    await raw.saveProject({ id: 'other', projectKey: 'other', title: 'Private other project', ownerKey: 'other', ownerName: 'Other' });
    await raw.saveTask({ id: 'task', projectKey: 'mine', title: 'Assigned task', phaseKey: 'requirement', status: 'Not Started', dueDate: '2026-09-01', ownerKey: 'engineer' });
    await renderApp();
    expect(document.body.textContent).not.toContain('Private other project');
    expect(document.body.textContent).not.toContain('New project');
    expect(document.body.textContent).toContain('Users and managers');
    await act(async () => document.querySelector('.project-card').click());
    expect(document.body.textContent).not.toContain('Edit project');
    await act(async () => document.querySelector('.upcoming-list button').click());
    expect(document.querySelector('.modal input[type="date"]').disabled).toBe(true);
    expect(document.querySelector('.modal input[aria-label="Task name"]').disabled).toBe(true);
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

  it('registers standard users and grants testing manager access only after the right password', async () => {
    await renderApp(false);
    expect([...document.querySelectorAll('.sidebar nav button')].map((row) => row.textContent)).toEqual(['My work0', 'Reference Documents', 'Acronym glossary', 'Users and managers']);
    expect(JSON.parse(localStorage.getItem('modernization-project-tracker:v2')).users[0]).toMatchObject({ title: 'Local Engineer', role: 'User' });
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((row) => row.textContent === 'Users and managers').click());
    expect(document.querySelectorAll('.directory-row')).toHaveLength(1);
    expect(document.querySelector('.directory-row button')).toBeNull();
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

  it('shows a persistent activation error when the directory write fails', async () => {
    await renderApp(false);
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((row) => row.textContent === 'Users and managers').click());
    vi.spyOn(Object.getPrototypeOf(createRepository().store), 'saveUser').mockRejectedValue(new Error('SharePoint denied the directory update (403).'));
    await act(async () => changeValue(document.querySelector('input[type="password"]'), 'admin123'));
    await act(async () => [...document.querySelectorAll('.directory-form button')].find((row) => row.textContent === 'Enable manager access').click());
    expect(document.querySelector('[role="alert"]').textContent).toContain('403');
    expect(document.body.textContent).not.toContain('New project');
    expect(JSON.parse(localStorage.getItem('modernization-project-tracker:v2')).users[0].role).toBe('User');
    expect(document.querySelector('.directory-form button').disabled).toBe(false);
  });

  it('lets a standard project owner add tasks and persist task-based progress', async () => {
    const raw = createRepository().store;
    await raw.saveProject({ id: 'owner-project', projectKey: 'owner-project', title: 'Owner project', ownerKey: 'local', ownerName: 'Local Engineer', status: 'In Progress' });
    await raw.saveTask({ id: 'done', projectKey: 'owner-project', title: 'Done task', phaseKey: 'requirement', status: 'Complete' });
    await renderApp(false);
    await act(async () => document.querySelector('.project-card').click());
    await act(async () => changeValue(document.querySelector('[aria-label="Progress calculation"]'), 'tasks'));
    expect(document.querySelector('.drawer-progress').textContent).toContain('1 of 1 tasks completed');
    await act(async () => [...document.querySelectorAll('.drawer-tabs button')].find((row) => row.textContent.includes('Work breakdown')).click());
    await act(async () => [...document.querySelectorAll('.task-stage button')].find((row) => row.getAttribute('aria-label')?.startsWith('Add task to ')).click());
    expect(document.querySelector('[aria-label="Task name"]').disabled).toBe(false);
    expect(document.querySelector('.modal input[type="date"]').disabled).toBe(true);
    await act(async () => changeValue(document.querySelector('[aria-label="Task name"]'), 'Owner entered task'));
    await act(async () => [...document.querySelectorAll('.modal button')].find((row) => row.textContent === 'Save task').click());
    expect(document.querySelector('.drawer-progress').textContent).toContain('1 of 2 tasks completed');
    expect(document.querySelector('.drawer-progress').textContent).toContain('50%');
    expect(JSON.parse(localStorage.getItem('modernization-project-tracker:v2')).projects[0].progressMode).toBe('tasks');
  });

  it('groups program office work, preserves project health, and switches task markers and portfolio progress', async () => {
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
    await act(async () => [...document.querySelectorAll('.kpi-card')].find((card) => card.textContent.includes('Needs attention')).click());
    expect(document.querySelectorAll('.attention-group')).toHaveLength(1);
    expect(document.querySelector('.attention-group header').textContent).toContain('In Progress – At Program Office');
    expect(document.querySelector('.attention-group').textContent).not.toContain('skip task');
    await act(async () => document.querySelector('.attention-collapse').click());
    expect(document.querySelector('.attention-collapse').getAttribute('aria-expanded')).toBe('true');
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

  it.each(['SME', 'Manager', 'User'])('saves %s through Edit for a user without email and preserves SharePoint identity', async (role) => {
    await createRepository().store.saveUser({ id: 'other', title: 'Other Engineer', email: '', loginName: 'i:0#.w|domain\\engineer', role: role === 'User' ? 'SME' : 'User' });
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
    await createRepository().store.saveUser({ id: 'other', title: 'Other Engineer', email: 'other@example.test', loginName: 'other', role: 'User' });
    await renderApp();
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((row) => row.textContent === 'Users and managers').click());
    const card = [...document.querySelectorAll('.directory-row')].find((row) => row.textContent.includes('Other Engineer'));
    await act(async () => card.querySelector('button').click());
    expect(document.querySelector('.directory-form')).toBeNull();
    await act(async () => [...card.querySelectorAll('button')].find((button) => button.textContent === 'Update User').click());
    await act(async () => changeValue(document.querySelector('.directory-form select'), 'SME'));
    vi.spyOn(Object.getPrototypeOf(createRepository().store), 'saveUser').mockRejectedValue(new Error('SharePoint denied the directory update (403).'));
    await act(async () => document.querySelector('.directory-form button').click());
    expect(document.querySelector('.page-stack > [role="alert"]').textContent).toContain('403');
    expect(document.querySelector('.directory-form select').value).toBe('SME');
    expect(card.textContent).toContain('User');
    expect(document.querySelector('.directory-form button').disabled).toBe(false);
  });

  it('collapses projects independently within a status and saves the assigned date', async () => {
    const raw = createRepository().store;
    for (const id of ['alpha', 'beta']) {
      await raw.saveProject({ id, projectKey: id, title: id, ownerName: 'Engineer' });
      await raw.saveTask({ id: `${id}-task`, projectKey: id, title: `${id} work`, phaseKey: 'requirement', status: 'In Progress' });
    }
    await renderApp();
    await act(async () => [...document.querySelectorAll('.kpi-card')].find((row) => row.textContent.includes('Needs attention')).click());
    expect(document.querySelector('.attention-group').tagName).toBe('SECTION');
    const groups = document.querySelectorAll('.attention-project-group');
    expect(groups).toHaveLength(2);
    expect([...groups].every((group) => group.querySelector('.attention-collapse').getAttribute('aria-expanded') === 'false')).toBe(true);
    await act(async () => groups[0].querySelector('.attention-collapse').click());
    expect(groups[0].querySelector('.attention-collapse').getAttribute('aria-expanded')).toBe('true');
    expect(groups[1].querySelector('.attention-collapse').getAttribute('aria-expanded')).toBe('false');
    await act(async () => groups[0].click());
    await act(async () => document.querySelector('.upcoming-list button').click());
    const field = [...document.querySelectorAll('.modal .field')].find((field) => field.textContent.includes('Assigned / creation date')).querySelector('input');
    expect(field.value).toBe('');
    await act(async () => changeValue(field, '2026-09-15'));
    await act(async () => [...document.querySelectorAll('.modal button')].find((row) => row.textContent === 'Save task').click());
    expect((await createRepository().store.load()).tasks.find((row) => row.id === 'alpha-task').assignedDate).toBe('2026-09-15');
  });

  it('gives SMEs full portfolio visibility with read-only task details', async () => {
    const raw = createRepository().store;
    await raw.saveUser({ id: 'sme', title: 'Local Engineer', loginName: 'local', role: 'SME' });
    await raw.saveProject({ id: 'outside', projectKey: 'outside', title: 'Other project', ownerKey: 'other', ownerName: 'Other Engineer' });
    await raw.saveTask({ id: 'outside-task', projectKey: 'outside', title: 'Other task', status: 'Not Started', phaseKey: 'requirement', estimatedHours: 4.5 });
    await renderApp(false);
    expect(document.body.textContent).toContain('Modernization at a glance');
    expect(document.body.textContent).toContain('Other project');
    expect(document.body.textContent).not.toContain('New project');
    await act(async () => document.querySelector('.project-table-row:not(.table-header)').click());
    expect(document.querySelector('[aria-label="Project settings"]')).toBeNull();
    await act(async () => document.querySelector('.upcoming-list button').click());
    expect(document.querySelector('.modal h2').textContent).toBe('Task details');
    expect(document.querySelector('.task-fields').disabled).toBe(true);
    expect(document.querySelector('[aria-label="Est. Hours"]').value).toBe('4.5');
    expect(document.querySelector('[aria-label="Attach documents"]')).toBeNull();
    expect([...document.querySelectorAll('.modal button')].some((row) => row.textContent === 'Save task')).toBe(false);
  });

  it('selects a directory person, saves the verified SME identity, and prepares an invitation', async () => {
    const prototype = Object.getPrototypeOf(createRepository().store);
    const candidate = { title: 'New Engineer', email: 'new@example.test', loginName: 'i:0#.f|membership|new@example.test' };
    vi.spyOn(prototype, 'searchPeople').mockResolvedValue([candidate]);
    const resolve = vi.spyOn(prototype, 'resolvePerson').mockResolvedValue(candidate);
    window.MOD_TRACKER_CONFIG.appUrl = 'https://tenant.sharepoint-mil.us/sites/mod/SitePages/Tracker.aspx';
    await renderApp();
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((row) => row.textContent === 'Users and managers').click());
    await act(async () => changeValue(document.querySelector('.people-invite input[type="search"]'), 'New Engineer'));
    await act(async () => new Promise((done) => setTimeout(done, 400)));
    await act(async () => document.querySelector('.people-results button').click());
    await act(async () => changeValue(document.querySelector('.people-invite select'), 'SME'));
    await act(async () => [...document.querySelectorAll('.people-invite button')].find((button) => button.textContent === 'Add to tracker').click());
    expect(resolve).toHaveBeenCalledWith(candidate.loginName);
    const saved = (await createRepository().store.load()).users.find((row) => row.loginName === candidate.loginName);
    expect(saved).toMatchObject({ ...candidate, role: 'SME' });
    expect(document.querySelector('.people-invite input[type="url"]')).toBeNull();
    expect(document.querySelector('.invitation-result a')).toBeNull();
    expect(document.querySelector('.invitation-result').textContent).toContain('Preview only: no site access granted or email sent.');
  });

  it('retries failed site sharing without duplicating a saved user', async () => {
    const person = { title: 'Invite Engineer', email: 'invite@example.test', loginName: 'claims|invite' };
    const store = { searchPeople: vi.fn(async () => [person]), resolvePerson: vi.fn(async () => person), shareSiteAccess: vi.fn().mockRejectedValueOnce(new Error('Site sharing denied')).mockResolvedValueOnce({ access: 'Read', emailRequested: true }) };
    const save = vi.fn(async (row) => ({ ...row, id: 'invited' }));
    await act(async () => { root = createRoot(document.getElementById('root')); root.render(<PeopleInvite store={store} config={{ appUrl: 'https://tenant.sharepoint.com/sites/mod/app.aspx' }} onSave={save} localPreview={false} />); });
    await act(async () => changeValue(document.querySelector('input[type="search"]'), 'Invite'));
    await act(async () => new Promise((done) => setTimeout(done, 400)));
    await act(async () => document.querySelector('.people-results button').click());
    await act(async () => changeValue(document.querySelector('select'), 'SME'));
    await act(async () => [...document.querySelectorAll('button')].find((button) => button.textContent === 'Invite and grant access').click());
    expect(document.querySelector('[role="alert"]').textContent).toContain('Site sharing denied');
    expect(document.querySelector('.invitation-result').textContent).toContain('not yet confirmed');
    await act(async () => [...document.querySelectorAll('button')].find((button) => button.textContent === 'Retry site invitation').click());
    expect(save).toHaveBeenCalledTimes(1);
    expect(store.shareSiteAccess).toHaveBeenCalledTimes(2);
    expect(document.querySelector('.invitation-result').textContent).toContain('Invitation sent to SharePoint for delivery.');
    expect(store.shareSiteAccess).toHaveBeenLastCalledWith(expect.objectContaining({ loginName: person.loginName }), 'SME', 'https://tenant.sharepoint.com/sites/mod/app.aspx');
  });

  it('creates nested references and moves a folder with persistence', async () => {
    await renderApp(false);
    await act(async () => [...document.querySelectorAll('.sidebar nav button')].find((row) => row.textContent === 'Reference Documents').click());
    const click = async (text) => act(async () => [...document.querySelectorAll('.reference-library button')].find((row) => row.textContent === text).click());
    await click('New folder');
    await act(async () => changeValue(document.querySelector('.reference-editor input'), 'Templates'));
    await click('Save reference');
    await click('Templates');
    await click('New folder');
    await act(async () => changeValue(document.querySelector('.reference-editor input'), 'Acquisition'));
    await click('Save reference');
    let rows = await createRepository().store.listReferenceEntries();
    const parent = rows.find((row) => row.name === 'Templates');
    expect(rows.find((row) => row.name === 'Acquisition').parentId).toBe(parent.id);
    await act(async () => document.querySelector('[aria-label="Organize Acquisition"]').click());
    await act(async () => changeValue(document.querySelector('.reference-editor select'), ''));
    await click('Save reference');
    rows = await createRepository().store.listReferenceEntries();
    expect(rows.find((row) => row.name === 'Acquisition').parentId).toBe('');
  });

  it('deletes a user only from Edit, preserves assignments, and hides self-deletion', async () => {
    const raw = createRepository().store;
    await raw.saveUser({ id: 'other', title: 'Other Engineer', email: 'other@example.test', loginName: 'other', role: 'SME' });
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
