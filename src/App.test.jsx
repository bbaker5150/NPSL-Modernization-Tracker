import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { App } from './App';

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

    const mockButton = [...document.querySelectorAll('button')].find((button) => button.textContent.includes('Open sample portfolio'));
    await act(async () => mockButton.click());
    expect(document.body.textContent).toContain('14 total measurement areas');
    expect(document.body.textContent).toContain('Return to live SharePoint data');
    expect(document.body.textContent).not.toContain('Read-only');

    const boardButton = [...document.querySelectorAll('button')].find((button) => button.textContent.includes('Pipeline board'));
    await act(async () => boardButton.click());
    expect(document.body.textContent).toContain('Scan where every project sits');
    expect(document.querySelectorAll('.kanban-column')).toHaveLength(7);
  });

  it('lets users practice task edits and project deletion in the sample portfolio', async () => {
    await renderApp();
    const sampleButton = [...document.querySelectorAll('button')].find((button) => button.textContent.includes('Open sample portfolio'));
    await act(async () => sampleButton.click());

    const firstCard = document.querySelector('.project-card');
    expect(firstCard.textContent).not.toContain('No target date');
    expect(firstCard.querySelector('[title^="Target completion:"]')).not.toBeNull();
    await act(async () => firstCard.click());

    const tasksTab = [...document.querySelectorAll('.drawer-tabs button')].find((button) => button.textContent.includes('Work breakdown'));
    await act(async () => tasksTab.click());
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

    await act(async () => document.querySelector('.modal button[aria-label="Close"]').click());
    await act(async () => document.querySelector('.project-drawer button[aria-label="Close"]').click());
    const menuButton = document.querySelector('.project-card .project-menu');
    await act(async () => menuButton.click());
    const deleteButton = [...document.querySelectorAll('.project-menu-popover button')].find((button) => button.textContent.includes('Delete project'));
    await act(async () => { deleteButton.click(); await Promise.resolve(); });
    expect(window.confirm).toHaveBeenCalledOnce();
    expect(document.body.textContent).toContain('13 total measurement areas');
  });
});
