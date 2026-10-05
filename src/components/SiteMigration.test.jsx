import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, it, vi } from 'vitest';
import { SiteMigration } from './SiteMigration';
import { migrationSite, MIGRATION_SOURCE, MIGRATION_TARGET } from '../lib/siteMigration';
import { exportTrackerBackup, validateTrackerBackup, previewTrackerImport, importTrackerBackup } from '../lib/trackerBackup';

vi.mock('../lib/siteMigration', async importOriginal => ({ ...await importOriginal(), migrationSite: vi.fn() }));
vi.mock('../lib/trackerBackup', async importOriginal => ({ ...await importOriginal(), exportTrackerBackup: vi.fn(), validateTrackerBackup: vi.fn(), previewTrackerImport: vi.fn(), importTrackerBackup: vi.fn() }));
globalThis.IS_REACT_ACT_ENVIRONMENT = true;
let root;
afterEach(async () => { if (root) await act(async () => root.unmount()); vi.restoreAllMocks(); vi.resetAllMocks(); });
async function render(webUrl = MIGRATION_TARGET) {
  document.body.innerHTML = '<div id="root"></div>';
  root = createRoot(document.getElementById('root'));
  await act(async () => root.render(<SiteMigration store={{ webUrl }} />));
}
const button = label => [...document.querySelectorAll('button')].find(node => node.textContent === label);
async function choose() {
  const input = document.querySelector('input[type="file"]');
  Object.defineProperty(input, 'files', { configurable: true, value: [{ size: 100, text: async () => JSON.stringify({ version: 1, payload: { lists: [], files: [] } }) }] });
  await act(async () => input.dispatchEvent(new Event('change', { bubbles: true })));
}
it('hides the tools on unrelated sites and from non-owners', async () => {
  await render('https://example.invalid/sites/elsewhere');
  expect(migrationSite).not.toHaveBeenCalled();
  expect(document.querySelector('details')).toBeNull();
  migrationSite.mockReturnValue({ canMigrate: async () => false });
  await act(async () => root.render(<SiteMigration store={{ webUrl: MIGRATION_TARGET }} />));
  expect(document.querySelector('details')).toBeNull();
});
it('offers only export on the source and requires paused edits', async () => {
  migrationSite.mockReturnValue({ canMigrate: async () => true });
  exportTrackerBackup.mockRejectedValue(new Error('Source changed during export'));
  await render(MIGRATION_SOURCE);
  expect(migrationSite.mock.calls.map(args => args[0])).toEqual([MIGRATION_SOURCE]);
  expect(button('Choose backup file')).toBeUndefined();
  expect(button('Download backup').disabled).toBe(true);
  await act(async () => document.querySelector('input[type="checkbox"]').click());
  await act(async () => button('Download backup').click());
  expect(exportTrackerBackup).toHaveBeenCalledTimes(1);
  expect(document.querySelector('[role="alert"]').textContent).toContain('Source changed');
});
it('requires a valid package, paused edits and conflict approval before import', async () => {
  migrationSite.mockReturnValue({ canMigrate: async () => true });
  previewTrackerImport.mockResolvedValue({ lists: [{ key: 'users', title: 'Users', fields: ['AppRole'], entries: [{ source: { RecordId: 'u', Title: 'Owner', AppRole: 'Manager' }, destination: { AppRole: 'Viewer' }, conflict: true, names: [] }] }] });
  importTrackerBackup.mockResolvedValue({ verified: true });
  await render();
  expect(document.querySelector('details').open).toBe(false);
  expect(button('Preview import')).toBeUndefined();
  const picker = vi.spyOn(document.querySelector('input[type="file"]'), 'click');
  await act(async () => button('Choose backup file').click());
  expect(picker).toHaveBeenCalledOnce();
  await choose();
  expect(validateTrackerBackup).toHaveBeenCalledOnce();
  await act(async () => button('Preview import').click());
  expect(button('Import and verify').disabled).toBe(true);
  const checks = document.querySelectorAll('input[type="checkbox"]');
  await act(async () => checks[0].click());
  expect(button('Import and verify').disabled).toBe(true);
  await act(async () => checks[1].click());
  await act(async () => button('Import and verify').click());
  expect(importTrackerBackup.mock.calls[0][3].replaceConflicts).toBe(true);
  expect(button('Download verification report')).toBeTruthy();
  expect(button('Reload tracker')).toBeTruthy();
  expect(migrationSite.mock.calls.map(args => args[0])).toEqual([MIGRATION_TARGET]);
});
it('invalidates failed imports while retaining the package and last operation for resuming', async () => {
  migrationSite.mockReturnValue({ canMigrate: async () => true });
  previewTrackerImport.mockResolvedValue({ lists: [] });
  importTrackerBackup.mockImplementation(async (_backup, _target, _plan, options) => { options.onProgress('Importing Tasks: 25/40…'); throw new Error('Network interrupted'); });
  await render(); await choose();
  await act(async () => button('Preview import').click());
  await act(async () => document.querySelector('input[type="checkbox"]').click());
  await act(async () => button('Import and verify').click());
  expect(document.querySelector('[role="alert"]').textContent).toContain('preview the same package again');
  expect(document.querySelector('[role="status"]').textContent).toContain('Importing Tasks: 25/40');
  expect(button('Import and verify')).toBeUndefined();
  expect(button('Download verification report')).toBeUndefined();
  expect(button('Preview import')).toBeTruthy();
});
it('does not offer preview after package validation fails', async () => {
  migrationSite.mockReturnValue({ canMigrate: async () => true });
  validateTrackerBackup.mockRejectedValue(new Error('Backup checksum failed'));
  await render(); await choose();
  expect(document.querySelector('[role="alert"]').textContent).toContain('checksum');
  expect(button('Preview import')).toBeUndefined();
  expect(importTrackerBackup).not.toHaveBeenCalled();
});

it('exports without references and lets the owner explicitly turn off task file downloads', async () => {
  migrationSite.mockReturnValue({ canMigrate: async () => true });
  exportTrackerBackup.mockRejectedValue(new Error('Stopped for test'));
  await render(MIGRATION_SOURCE);
  expect(document.body.textContent).toContain('Reference Documents and their folders are excluded');
  const checks = document.querySelectorAll('input[type="checkbox"]');
  expect(checks[1].checked).toBe(true);
  await act(async () => checks[0].click());
  await act(async () => button('Download backup').click());
  expect(exportTrackerBackup.mock.calls[0][2]).toEqual({ includeReferences: false, includeTaskAttachments: true });
  await act(async () => checks[1].click());
  await act(async () => button('Download backup').click());
  expect(exportTrackerBackup.mock.calls[1][2]).toEqual({ includeReferences: false, includeTaskAttachments: false });
});
it('shows excluded content in the import preview and completion message', async () => {
  migrationSite.mockReturnValue({ canMigrate: async () => true });
  previewTrackerImport.mockResolvedValue({ lists: [] });
  importTrackerBackup.mockResolvedValue({ verified: true, manualTransfer: ['Reference Documents and folders', 'task attachments'] });
  await render();
  const input = document.querySelector('input[type="file"]');
  Object.defineProperty(input, 'files', { configurable: true, value: [{ size: 100, text: async () => JSON.stringify({ version: 2, payload: { scope: { includeReferences: false, includeTaskAttachments: false }, lists: [], files: [] } }) }] });
  await act(async () => input.dispatchEvent(new Event('change', { bubbles: true })));
  expect(document.querySelector('[role="note"]').textContent).toContain('Reference Documents and folders; task attachments');
  await act(async () => button('Preview import').click());
  await act(async () => document.querySelector('input[type="checkbox"]').click());
  await act(async () => button('Import and verify').click());
  expect(document.querySelector('[role="status"]').textContent).toContain('Manual transfer still required');
});
