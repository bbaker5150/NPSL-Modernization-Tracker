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

  async function renderApp() {
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

  it('boots a clean workspace with one centered first-project action and no workspace status card', async () => {
    await renderApp();

    expect(document.body.textContent).toContain('Modernization at a glance');
    expect(document.body.textContent).toContain('0 total measurement areas');
    expect(document.body.textContent).toContain('Your modernization portfolio is ready');
    expect(document.querySelector('.brand-logo')?.getAttribute('alt')).toBe('NAVAIR');
    expect(document.querySelector('.sidebar-user .avatar')?.textContent).toBe('LE');
    expect(document.querySelector('.sync-card')).toBeNull();
    expect(document.querySelectorAll('.empty-actions button')).toHaveLength(1);
    expect(document.querySelector('.empty-actions button')?.textContent).toContain('Add first project');
    expect(document.body.textContent).not.toContain('Demo workspace');
    expect(document.body.textContent).not.toContain('SharePoint workspace');
    expect(document.body.textContent).not.toContain('historical baseline');
    expect(document.body.textContent).not.toContain('Open sample portfolio');
    expect([...document.querySelectorAll('.phase-bar span')].every((bar) => bar.style.width === '0%')).toBe(true);

    const boardButton = [...document.querySelectorAll('button')].find((button) => button.textContent.includes('Pipeline board'));
    await act(async () => boardButton.click());
    expect(document.body.textContent).toContain('Scan where every project sits');
    expect(document.querySelectorAll('.kanban-column')).toHaveLength(5);

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
    expect(window.confirm).toHaveBeenCalledWith('Remove ABC from the shared glossary?');
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
    expect(document.body.textContent).toContain('5');
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
    expect(document.querySelectorAll('.attention-project')).toHaveLength(7);
    expect(document.body.textContent).toContain('Overdue');
    expect(document.body.textContent).toContain('Parts delayed');
    expect(document.body.textContent).toContain('Scope approved');
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
    expect(document.body.textContent).not.toContain('Users and managers');
    await act(async () => document.querySelector('.project-table-row:not(.table-header)').click());
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
    await act(async () => [...document.querySelectorAll('.task-stage button')].find((button) => button.textContent === 'Add task').click());
    await act(async () => changeValue(document.querySelector('[aria-label="Task name"]'), 'Custom task'));
    await act(async () => [...document.querySelectorAll('.modal button')].find((button) => button.textContent === 'Save task').click());
    expect(document.querySelectorAll('.phase-task')).toHaveLength(2);
    await act(async () => [...document.querySelectorAll('.phase-task')].find((row) => row.textContent.includes('Custom task')).click());
    await act(async () => [...document.querySelectorAll('.modal button')].find((button) => button.textContent === 'Delete task').click());
    expect(document.querySelectorAll('.phase-task')).toHaveLength(1);
    expect(document.querySelector('.phase-task').textContent).toContain('Keep this task');
    expect(JSON.parse(localStorage.getItem('modernization-project-tracker:v2')).tasks).toHaveLength(1);
  });

});
