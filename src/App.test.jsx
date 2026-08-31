import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
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
  });

  afterEach(async () => {
    if (root) await act(async () => root.unmount());
  });

  it('boots a clean workspace and opens the read-only mock portfolio on demand', async () => {
    await act(async () => {
      root = createRoot(document.getElementById('root'));
      root.render(<App />);
      await new Promise((resolve) => setTimeout(resolve, 30));
    });

    expect(document.body.textContent).toContain('Modernization at a glance');
    expect(document.body.textContent).toContain('0 total measurement areas');
    expect(document.body.textContent).toContain('Your modernization portfolio is ready');
    expect(document.querySelector('.brand-logo')?.getAttribute('alt')).toBe('NAVAIR');
    expect(document.querySelector('.sidebar-user .avatar')?.textContent).toBe('LE');
    expect(document.body.textContent).not.toContain('Demo workspace');
    expect(document.body.textContent).not.toContain('historical baseline');

    const mockButton = [...document.querySelectorAll('button')].find((button) => button.textContent.includes('Preview mock portfolio'));
    await act(async () => mockButton.click());
    expect(document.body.textContent).toContain('14 total measurement areas');
    expect(document.body.textContent).toContain('Return to live SharePoint data');

    const boardButton = [...document.querySelectorAll('button')].find((button) => button.textContent.includes('Pipeline board'));
    await act(async () => boardButton.click());
    expect(document.body.textContent).toContain('Scan where every project sits');
    expect(document.querySelectorAll('.kanban-column')).toHaveLength(7);
  });
});
