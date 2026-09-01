import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { App, userInitials } from './App';

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
  });

  afterEach(async () => {
    if (root) await act(async () => root.unmount());
  });

  async function renderApp() {
    await act(async () => {
      root = createRoot(document.getElementById('root'));
      root.render(<App />);
      await new Promise((resolve) => setTimeout(resolve, 30));
    });
  }

  function changeValue(element, value) {
    const prototype = element instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
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
    expect(document.body.textContent).toContain('No acronyms yet');
  });

  it('adds and removes shared glossary acronyms', async () => {
    await renderApp();
    const glossaryButton = [...document.querySelectorAll('.sidebar nav button')].find((button) => button.textContent.includes('Acronym glossary'));
    await act(async () => glossaryButton.click());

    await act(async () => {
      changeValue(document.querySelector('input[aria-label="Acronym"]'), 'css');
      changeValue(document.querySelector('input[aria-label="Full term"]'), 'Calibration Standard Specification');
      changeValue(document.querySelector('input[aria-label="Definition"]'), 'Technical requirements for a calibration standard.');
    });
    await act(async () => {
      document.querySelector('.glossary-composer').dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
      await Promise.resolve();
    });
    expect(document.body.textContent).toContain('Calibration Standard Specification');
    expect(document.querySelector('input[aria-label="Acronym"]').value).toBe('');

    await act(async () => {
      document.querySelector('button[aria-label="Remove CSS"]').click();
      await Promise.resolve();
    });
    expect(window.confirm).toHaveBeenCalledWith('Remove CSS from the shared glossary?');
    expect(document.body.textContent).toContain('No acronyms yet');
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

  it('lets users edit tasks and delete projects without sample data', async () => {
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
    const phaseSection = phaseButton.closest('.wbs-stage');
    await act(async () => { phaseButton.click(); await Promise.resolve(); });
    expect(phaseSection.classList.contains('phase-not-required')).toBe(true);
    expect(phaseSection.textContent).toContain('Restore phase');
    await act(async () => { phaseSection.querySelector('.phase-required-button').click(); await Promise.resolve(); });
    expect(phaseSection.classList.contains('phase-not-required')).toBe(false);
    const checkbox = document.querySelector('.task-check:not(.checked)');
    const wbs = checkbox.getAttribute('aria-label').replace('Complete WBS ', '');
    await act(async () => { checkbox.click(); await Promise.resolve(); });
    expect(document.querySelector('.modal')).toBeNull();
    expect(document.querySelector(`[aria-label="Reopen WBS ${wbs}"]`)).not.toBeNull();

    const taskRow = document.querySelector(`[aria-label="Reopen WBS ${wbs}"]`).closest('.wbs-task');
    await act(async () => taskRow.click());
    const statusSelect = [...document.querySelectorAll('.modal select')].find((select) => [...select.options].some((option) => option.textContent === 'Not Required'));
    expect(statusSelect).not.toBeUndefined();
    expect([...statusSelect.options].map((option) => option.textContent)).toContain('Not Required');
    const progressBeforeNotRequired = document.querySelector('.drawer-progress span').textContent;
    const wbsInput = document.querySelector('.modal input[aria-label="WBS code"]');
    const dueDate = document.querySelector('.modal input[type="date"]');
    await act(async () => {
      changeValue(wbsInput, '9.9.9');
      changeValue(statusSelect, 'Not Required');
    });
    expect(dueDate.disabled).toBe(true);
    expect(dueDate.value).toBe('');
    const saveTaskButton = [...document.querySelectorAll('.modal .button.primary')].find((button) => button.textContent.includes('Save task'));
    await act(async () => { saveTaskButton.click(); await Promise.resolve(); });
    expect(document.querySelector('[aria-label="Reopen WBS 9.9.9"]')).not.toBeNull();
    expect(document.querySelector('.drawer-progress span').textContent).toBe(progressBeforeNotRequired);

    await act(async () => document.querySelector('.project-drawer button[aria-label="Close"]').click());
    const menuButton = document.querySelector('.project-card .project-menu');
    await act(async () => menuButton.click());
    const deleteButton = [...document.querySelectorAll('.project-menu-popover button')].find((button) => button.textContent.includes('Delete project'));
    await act(async () => { deleteButton.click(); await Promise.resolve(); });
    expect(window.confirm).not.toHaveBeenCalled();
    expect(document.body.textContent).toContain('0 total measurement areas');
    expect(document.querySelector('[role="status"]')?.textContent).toContain('deleted');
  });
});
