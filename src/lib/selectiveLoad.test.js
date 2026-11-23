import { beforeEach, expect, it, vi } from 'vitest';
import { SharePointStore } from './spStore';
import { createRepository } from './repository';

beforeEach(() => { localStorage.clear(); window.MOD_TRACKER_CONFIG = { forceLocal: true }; });

it('omits glossary rows only when requested, preserving fresh core reads and full-load exports', async () => {
  const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/tracker' });
  let glossary = [{ id: 'first', acronym: 'FIRST' }];
  store.listItems = vi.fn(async key => key === 'acronyms' ? glossary : []);
  expect((await store.load({ includeAcronyms: false })).acronyms).toEqual([]);
  expect(store.listItems.mock.calls.map(([key]) => key)).toEqual(['projects', 'tasks', 'updates', 'risks', 'users']);
  expect((await store.load()).acronyms).toEqual(glossary);
  glossary = [{ id: 'latest', acronym: 'LATEST' }];
  expect(await store.listAcronyms()).toEqual(glossary);
});

it('keeps local preview full-load compatibility and returns an independent glossary snapshot', async () => {
  const store = createRepository().store;
  expect((await store.load()).acronyms.length).toBeGreaterThan(0);
  expect((await store.load({ includeAcronyms: false })).acronyms).toEqual([]);
  const entries = await store.listAcronyms();
  const count = entries.length;
  entries.pop();
  expect((await store.listAcronyms()).length).toBe(count);
});

it('skips an unchanged project record write while preserving permission repair and its retry', async () => {
  const store = new SharePointStore({ webUrl: 'https://tenant.sharepoint.com/sites/tracker', scopedAccess: true });
  const project = { id: 'project', spId: 7, projectKey: 'project', title: 'Project', engineers: [], progressMode: 'phases' };
  store.currentUser = vi.fn(async () => ({ id: 1 }));
  store.permissions.currentRole = vi.fn(async () => ({ role: 'Manager' }));
  store.permissionProjects = vi.fn(async () => [project]);
  store.directoryRows = vi.fn(async () => []);
  store.permissions.syncProject = vi.fn().mockRejectedValueOnce(new Error('ACL interrupted')).mockResolvedValue(undefined);
  store.update = vi.fn(); store.create = vi.fn();
  await expect(store.saveProject({ ...project })).rejects.toThrow('access setup is incomplete');
  await expect(store.saveProject({ ...project })).resolves.toEqual(project);
  expect(store.update).not.toHaveBeenCalled(); expect(store.create).not.toHaveBeenCalled();
  expect(store.permissions.syncProject).toHaveBeenCalledTimes(2);
  expect(store.permissions.currentRole).toHaveBeenCalledTimes(2);
  await store.saveProject({ ...project, title: 'Changed title' });
  expect(store.update).toHaveBeenCalledExactlyOnceWith('projects', 7, expect.objectContaining({ Title: 'Changed title' }));
  expect(store.permissions.syncProject).toHaveBeenCalledTimes(3);
});
