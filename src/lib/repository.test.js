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
    window.MOD_TRACKER_CONFIG = { webUrl: 'https://tenant.example.invalid/sites/Modernization', forceSharePoint: true };
    expect(createRepository().mode).toBe('sharepoint');
  });

  it('starts local development as a clean slate', async () => {
    window.MOD_TRACKER_CONFIG = { forceLocal: true };
    const repo = createRepository();
    expect(repo.mode).toBe('local');
    await expect(repo.store.load()).resolves.toEqual({ projects: [], tasks: [], updates: [], risks: [], acronyms: [] });
  });

  it('keys My Work to the signed-in SharePoint login with legacy email fallback', () => {
    const user = { id: 17, loginName: 'i:0#.f|membership|engineer@example.invalid', email: 'engineer@example.invalid' };
    expect(userIdentityKey(user)).toBe('i:0#.f|membership|engineer@example.invalid');
    expect(isOwnedByUser({ ownerKey: 'I:0#.F|MEMBERSHIP|ENGINEER@EXAMPLE.INVALID' }, user)).toBe(true);
    expect(isOwnedByUser({ ownerEmail: 'ENGINEER@EXAMPLE.INVALID' }, user)).toBe(true);
    expect(isOwnedByUser({ ownerKey: 'someone-else', ownerEmail: 'engineer@example.invalid' }, user)).toBe(true);
    expect(isOwnedByUser({ ownerKey: 'i:0#.f|membership|engineer@example.invalid' }, { email: 'engineer@example.invalid' })).toBe(true);
  });

  it('creates a clean five-phase WBS when a live project is added', () => {
    const owner = { ownerName: 'Doe, Jordan T CIV (USA)', ownerEmail: 'engineer@example.invalid', ownerKey: 'i:0#.f|membership|engineer@example.invalid' };
    const tasks = createStarterTasks('new-project', owner);
    expect(tasks).toHaveLength(5);
    expect(new Set(tasks.map((task) => task.phaseKey))).toHaveLength(5);
    expect(tasks.every((task) => task.projectKey === 'new-project' && task.status === 'Not Started')).toBe(true);
    expect(tasks.every((task) => !task.dueDate && task.ownerKey === owner.ownerKey && task.ownerEmail === owner.ownerEmail && task.ownerName === owner.ownerName && !task.dataIssue)).toBe(true);
  });
});
