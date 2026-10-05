import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, it, vi } from 'vitest';
import { SiteMigration } from './SiteMigration';
import { migrationSite, MIGRATION_TARGET, previewMigration, runMigration } from '../lib/siteMigration';

vi.mock('../lib/siteMigration', async importOriginal => ({
  ...await importOriginal(), migrationSite: vi.fn(), previewMigration: vi.fn(), runMigration: vi.fn(),
}));
globalThis.IS_REACT_ACT_ENVIRONMENT = true;
let root;
afterEach(async () => { if (root) await act(async () => root.unmount()); vi.resetAllMocks(); });
async function render(webUrl = MIGRATION_TARGET) {
  document.body.innerHTML = '<div id="root"></div>';
  root = createRoot(document.getElementById('root'));
  await act(async () => root.render(<SiteMigration store={{ webUrl }} />));
}
const button = label => [...document.querySelectorAll('button')].find(node => node.textContent === label);
it('does not expose migration on the source site or to a member', async () => {
  await render('https://flankspeed.sharepoint-mil.us/sites/ISEAMETENG');
  expect(migrationSite).not.toHaveBeenCalled();
  expect(document.querySelector('details')).toBeNull();
  migrationSite.mockReturnValue({ canMigrate: async () => false });
  await act(async () => root.render(<SiteMigration store={{ webUrl: MIGRATION_TARGET }} />));
  expect(document.querySelector('details')).toBeNull();
});
it('requires paused edits and explicit conflict approval, then exposes a verified report', async () => {
  migrationSite.mockReturnValue({ canMigrate: async () => true });
  const plan = { lists: [{ key: 'users', title: 'Users', fields: ['AppRole'], entries: [{ source: { RecordId: 'u', Title: 'Owner', AppRole: 'Manager' }, destination: { AppRole: 'Viewer' }, conflict: true, names: [] }] }] };
  previewMigration.mockResolvedValue(plan);
  runMigration.mockResolvedValue({ verified: true });
  await render();
  expect(document.querySelector('details').open).toBe(false);
  await act(async () => button('Preview migration').click());
  expect(button('Copy and verify').disabled).toBe(true);
  const checks = document.querySelectorAll('input[type="checkbox"]');
  await act(async () => checks[0].click());
  expect(button('Copy and verify').disabled).toBe(true);
  await act(async () => checks[1].click());
  expect(button('Copy and verify').disabled).toBe(false);
  await act(async () => button('Copy and verify').click());
  expect(runMigration.mock.calls[0][3].replaceConflicts).toBe(true);
  expect(button('Download verification report')).toBeTruthy();
  expect(button('Reload tracker')).toBeTruthy();
});
it('invalidates the preview after failure so a retry checks partial destination state', async () => {
  migrationSite.mockReturnValue({ canMigrate: async () => true });
  previewMigration.mockResolvedValue({ lists: [] });
  runMigration.mockRejectedValue(new Error('Network interrupted'));
  await render();
  await act(async () => button('Preview migration').click());
  await act(async () => document.querySelector('input').click());
  await act(async () => button('Copy and verify').click());
  expect(document.querySelector('[role="alert"]').textContent).toContain('preview again to resume');
  expect(button('Copy and verify')).toBeUndefined();
  expect(button('Download verification report')).toBeUndefined();
});
