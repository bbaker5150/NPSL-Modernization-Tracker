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
    const data = await repo.store.load();
    expect(data.projects).toEqual([]);
    expect(data.tasks).toEqual([]);
    expect(data.acronyms).toHaveLength(37);
    expect(data.acronyms.find((entry) => entry.acronym === 'CSS')?.term).toBe('Calibration Standard Specification');
  });

  it('enables scoped access on metsoft and disables testing promotion without changing ISEA export behavior', () => {
    window.MOD_TRACKER_CONFIG = { webUrl: 'https://flankspeed.sharepoint-mil.us/sites/metsoft', testingManagerPassword: 'admin123' };
    const repo = createRepository();
    expect(repo.store.scopedAccess).toBe(true);
    expect(repo.config.testingManagerPassword).toBe(false);
    window.MOD_TRACKER_CONFIG = { webUrl: 'https://flankspeed.sharepoint-mil.us/sites/ISEAMETENG' };
    expect(createRepository().store.scopedAccess).toBe(false);
  });

  it('uses the confirmed metsoft assets folder and respects deployment overrides', () => {
    window.MOD_TRACKER_CONFIG = { webUrl: 'https://flankspeed.sharepoint-mil.us/sites/metsoft' };
    expect(createRepository().store.invitationAssetFolders).toEqual(['/sites/metsoft/SiteAssets/Modernization Tracker']);
    window.MOD_TRACKER_CONFIG.invitationAssetFolders = ['/sites/metsoft/SiteAssets/Other'];
    window.MOD_TRACKER_CONFIG.invitationAssetUrls = ['/sites/metsoft/SiteAssets/Other/app.html'];
    expect(createRepository().store.invitationAssetFolders).toEqual(window.MOD_TRACKER_CONFIG.invitationAssetFolders);
    expect(createRepository().store.invitationAssetUrls).toEqual(window.MOD_TRACKER_CONFIG.invitationAssetUrls);
    window.MOD_TRACKER_CONFIG = { webUrl: 'https://tenant.sharepoint.com/sites/metsoft' };
    expect(createRepository().store.invitationAssetFolders).toEqual([]);
  });

  it('keys My Work to the signed-in SharePoint login with legacy email fallback', () => {
    const user = { id: 17, loginName: 'i:0#.f|membership|engineer@example.invalid', email: 'engineer@example.invalid' };
    expect(userIdentityKey(user)).toBe('i:0#.f|membership|engineer@example.invalid');
    expect(isOwnedByUser({ ownerKey: 'I:0#.F|MEMBERSHIP|ENGINEER@EXAMPLE.INVALID' }, user)).toBe(true);
    expect(isOwnedByUser({ ownerEmail: 'ENGINEER@EXAMPLE.INVALID' }, user)).toBe(true);
    expect(isOwnedByUser({ ownerKey: 'someone-else', ownerEmail: 'engineer@example.invalid' }, user)).toBe(true);
    expect(isOwnedByUser({ ownerKey: 'i:0#.f|membership|engineer@example.invalid' }, { email: 'engineer@example.invalid' })).toBe(true);
  });

  it('creates a clean four-phase WBS when a live project is added', () => {
    const owner = { ownerName: 'Doe, Jordan T CIV (USA)', ownerEmail: 'engineer@example.invalid', ownerKey: 'i:0#.f|membership|engineer@example.invalid' };
    const tasks = createStarterTasks('new-project', owner);
    expect(tasks).toHaveLength(4);
    expect(new Set(tasks.map((task) => task.phaseKey))).toHaveLength(4);
    expect(tasks.every((task) => task.projectKey === 'new-project' && task.status === 'Not Started')).toBe(true);
    expect(tasks.every((task) => !task.dueDate && task.ownerKey === owner.ownerKey && task.ownerEmail === owner.ownerEmail && task.ownerName === owner.ownerName && !task.dataIssue)).toBe(true);
  });
});
