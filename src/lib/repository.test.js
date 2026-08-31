import { beforeEach, describe, expect, it, vi } from 'vitest';
import { createRepository, createStarterTasks, isOwnedByUser, userIdentityKey } from './repository';

describe('repository selection and SharePoint identity', () => {
  beforeEach(() => {
    localStorage.clear();
    window.MOD_TRACKER_CONFIG = {};
  });

  it('selects SharePoint from the resolved web URL even inside a srcdoc frame', () => {
    window.MOD_TRACKER_CONFIG = { webUrl: 'https://tenant.sharepoint.com/sites/Modernization' };
    const repo = createRepository();
    expect(repo.mode).toBe('sharepoint');
    expect(repo.store.webUrl).toBe('https://tenant.sharepoint.com/sites/Modernization');
    window.MOD_TRACKER_CONFIG = { webUrl: 'https://tenant.sharepoint-mil.us/sites/Modernization' };
    expect(createRepository().mode).toBe('sharepoint');
  });

  it('starts local development as a clean slate', async () => {
    window.MOD_TRACKER_CONFIG = { forceLocal: true };
    const repo = createRepository();
    expect(repo.mode).toBe('local');
    await expect(repo.store.load()).resolves.toEqual({ projects: [], tasks: [], updates: [], risks: [] });
  });

  it('keys My Work to the signed-in SharePoint login with legacy email fallback', () => {
    const user = { id: 17, loginName: 'i:0#.f|membership|engineer@navy.mil', email: 'engineer@navy.mil' };
    expect(userIdentityKey(user)).toBe('i:0#.f|membership|engineer@navy.mil');
    expect(isOwnedByUser({ ownerKey: 'I:0#.F|MEMBERSHIP|ENGINEER@NAVY.MIL' }, user)).toBe(true);
    expect(isOwnedByUser({ ownerEmail: 'ENGINEER@NAVY.MIL' }, user)).toBe(true);
    expect(isOwnedByUser({ ownerKey: 'someone-else', ownerEmail: 'engineer@navy.mil' }, user)).toBe(false);
  });

  it('creates a clean seven-stage WBS when a live project is added', () => {
    const tasks = createStarterTasks('new-project');
    expect(tasks).toHaveLength(34);
    expect(new Set(tasks.map((task) => task.phaseKey))).toHaveLength(7);
    expect(tasks.every((task) => task.projectKey === 'new-project' && task.status === 'Not Started')).toBe(true);
    expect(tasks.every((task) => !task.dueDate && !task.ownerKey && !task.dataIssue)).toBe(true);
  });
});
