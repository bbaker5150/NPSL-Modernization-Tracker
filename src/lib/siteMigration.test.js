// @vitest-environment node
import { describe, expect, it, vi } from 'vitest';
import { CONTAINERS } from './spStore';
import { migrationSite, MIGRATION_SOURCE, MIGRATION_TARGET, previewMigration, runMigration } from './siteMigration';

function memorySite(webUrl, seed = {}) {
  const rows = Object.fromEntries(CONTAINERS.map(c => [c.key, structuredClone(seed[c.key] || [])]));
  const files = new Map(); let nextId = 100;
  const fileKey = (key, id, name) => `${key}/${id}/${name}`;
  const api = {
    webUrl, rows, files,
    canMigrate: vi.fn(async () => true),
    schema: vi.fn(async container => new Set(['Title', ...container.fields.map(f => f.name)])),
    readRows: vi.fn(async key => structuredClone(rows[key])),
    readItem: vi.fn(async (key, id) => ({ ...structuredClone(rows[key].find(row => row.Id === id)), etag: '"1"' })),
    create: vi.fn(async (key, fields) => { const Id = nextId++; rows[key].push({ ...fields, Id, Modified: 'new' }); return Id; }),
    update: vi.fn(async (key, id, fields) => Object.assign(rows[key].find(row => row.Id === id), fields, { Modified: 'updated' })),
    attachments: vi.fn(async (key, id) => [...files.keys()].filter(path => path.startsWith(`${key}/${id}/`)).map(path => path.slice(`${key}/${id}/`.length))),
    bytes: vi.fn(async (key, id, name) => { const bytes = files.get(fileKey(key, id, name)); if (!bytes) throw new Error('Missing file'); return bytes.slice(0); }),
    attach: vi.fn(async (key, id, name, bytes) => files.set(fileKey(key, id, name), bytes.slice(0))),
  };
  return api;
}
const record = (RecordId, fields = {}) => ({ Id: 1, RecordId, Title: RecordId, Archived: false, Modified: 'original', ...fields });
function fixture() {
  const source = memorySite(MIGRATION_SOURCE, {
    projects: [record('p', { ProjectKey: 'project-key' })],
    tasks: [record('t', { ProjectKey: 'project-key', Archived: true, ArchivedDocuments: '["hidden.pdf"]', EstimatedHours: 0 })],
    updates: [record('u', { ProjectKey: 'project-key', Summary: 'A decision' })],
    risks: [record('r', { ProjectKey: 'project-key' })],
    users: [record('person', { LoginKey: 'i:0#.f|membership|owner@example.invalid', Email: 'owner@example.invalid', AppRole: 'Manager' })],
    acronyms: [record('a', { Acronym: 'NPSL', Definition: 'Custom definition' })],
    references: [record('folder', { EntryKind: 'folder' }), record('file', { Id: 2, ReferenceParentId: 'folder', EntryKind: 'file', FileName: 'template.xlsx', FileSize: 4 })],
  });
  source.files.set('tasks/1/hidden.pdf', new Uint8Array([0, 1, 254, 255]).buffer);
  source.files.set('references/2/template.xlsx', new Uint8Array([4, 3, 2, 1]).buffer);
  return { source, target: memorySite(MIGRATION_TARGET) };
}

describe('owner-run site migration', () => {
  it('preserves archived user history while matching the recreated active account, including retries', async () => {
    const { source, target } = fixture();
    source.rows.users.unshift(record('old-person', { Id: 8, LoginKey: 'owner@example.invalid', AppRole: 'Viewer', Archived: true }));
    target.rows.users.push(record('bootstrap-person', { Id: 9, LoginKey: 'owner@example.invalid', AppRole: 'Viewer' }));
    const report = await runMigration(source, target, await previewMigration(source, target), { replaceConflicts: true });
    expect(report.verified).toBe(true);
    expect(target.rows.users).toHaveLength(2);
    expect(target.rows.users.find(row => row.RecordId === 'old-person')).toMatchObject({ Archived: true, AppRole: 'Viewer' });
    expect(target.rows.users.find(row => row.RecordId === 'person')).toMatchObject({ Id: 9, Archived: false, AppRole: 'Manager' });
    const creates = target.create.mock.calls.length;
    await runMigration(source, target, await previewMigration(source, target));
    expect(target.create).toHaveBeenCalledTimes(creates);
    expect(source.update).not.toHaveBeenCalled();
  });
  it('does not reuse or reactivate an archived destination account by email', async () => {
    const { source, target } = fixture();
    target.rows.users.push(record('destination-history', { Id: 9, Email: 'owner@example.invalid', AppRole: 'Manager', Archived: true }));
    await runMigration(source, target, await previewMigration(source, target));
    expect(target.rows.users).toHaveLength(2);
    expect(target.rows.users.find(row => row.Id === 9)).toMatchObject({ RecordId: 'destination-history', Archived: true });
  });
  it('identifies the site, account, item IDs and roles for genuinely ambiguous active users', async () => {
    const { source, target } = fixture();
    source.rows.users.push(record('other-active', { Id: 7, LoginKey: 'OWNER@EXAMPLE.INVALID', AppRole: 'Viewer' }));
    await expect(previewMigration(source, target)).rejects.toThrow(/ISEA METENG source: multiple active user records.*item 1, role Manager.*item 7, role Viewer/);
    expect(target.create).not.toHaveBeenCalled();
    source.rows.users.pop();
    target.rows.users = structuredClone(source.rows.users);
    target.rows.users.push(record('destination-duplicate', { Id: 8, Email: 'owner@example.invalid', AppRole: 'Viewer' }));
    await expect(previewMigration(source, target)).rejects.toThrow('metsoft destination: multiple active user records');
  });
  it('preserves all seven lists, archived data, relationships and binary documents without source writes', async () => {
    const { source, target } = fixture();
    const before = structuredClone(source.rows);
    const plan = await previewMigration(source, target);
    expect(target.create).not.toHaveBeenCalled();
    const report = await runMigration(source, target, plan);
    expect(report.verified).toBe(true);
    expect(report.lists).toHaveLength(7);
    expect(report.lists.reduce((sum, list) => sum + list.records, 0)).toBe(8);
    expect(report.files).toHaveLength(2);
    expect(report.files.every(file => /^[a-f0-9]{64}$/.test(file.sha256))).toBe(true);
    expect(target.rows.tasks[0]).toMatchObject({ ProjectKey: 'project-key', Archived: true, ArchivedDocuments: '["hidden.pdf"]', EstimatedHours: 0 });
    expect(target.rows.references[1].ReferenceParentId).toBe('folder');
    expect(target.rows.references[1].Id).not.toBe(2);
    expect(target.rows.users[0].AppRole).toBe('Manager');
    expect(source.rows).toEqual(before);
    expect(source.create).not.toHaveBeenCalled(); expect(source.update).not.toHaveBeenCalled(); expect(source.attach).not.toHaveBeenCalled();
    await runMigration(source, target, await previewMigration(source, target));
    expect(target.create).toHaveBeenCalledTimes(8);
    expect(target.attach).toHaveBeenCalledTimes(2);
  });
  it('requires explicit conflict approval and matches an existing user by identity', async () => {
    const { source, target } = fixture();
    target.rows.users.push(record('new-site-person', { Id: 5, LoginKey: 'owner@example.invalid', AppRole: 'Viewer' }));
    const plan = await previewMigration(source, target);
    await expect(runMigration(source, target, plan)).rejects.toThrow('explicitly approve');
    expect(target.create).not.toHaveBeenCalled();
    await runMigration(source, target, plan, { replaceConflicts: true });
    expect(target.rows.users).toHaveLength(1);
    expect(target.rows.users[0]).toMatchObject({ Id: 5, RecordId: 'person', AppRole: 'Manager' });
    expect(target.update.mock.calls[0][3]).toBe('"1"');
  });
  it('resumes after an interrupted upload without duplicate records or files', async () => {
    const { source, target } = fixture();
    target.attach.mockRejectedValueOnce(new Error('Network interrupted'));
    await expect(runMigration(source, target, await previewMigration(source, target))).rejects.toThrow('Network interrupted');
    const result = await runMigration(source, target, await previewMigration(source, target));
    expect(result.verified).toBe(true);
    expect(target.create).toHaveBeenCalledTimes(8);
  });
  it('never overwrites a same-name destination document with different contents', async () => {
    const { source, target } = fixture();
    target.rows.tasks = structuredClone(source.rows.tasks);
    target.files.set('tasks/1/hidden.pdf', new Uint8Array([9]).buffer);
    await expect(runMigration(source, target, await previewMigration(source, target))).rejects.toThrow('No file was overwritten');
    expect(new Uint8Array(target.files.get('tasks/1/hidden.pdf'))).toEqual(new Uint8Array([9]));
    expect(target.attach).not.toHaveBeenCalled();
  });
  it('rejects a changed source snapshot before any writes', async () => {
    const { source, target } = fixture();
    const plan = await previewMigration(source, target);
    source.rows.tasks[0].Notes = 'Changed';
    await expect(runMigration(source, target, plan)).rejects.toThrow('Source changed');
    expect(target.create).not.toHaveBeenCalled();
  });
  it('detects source document changes before reporting success', async () => {
    const { source, target } = fixture();
    const plan = await previewMigration(source, target);
    const original = source.bytes.getMockImplementation();
    source.bytes.mockImplementation(async (...args) => {
      const value = await original(...args);
      source.files.set('tasks/1/hidden.pdf', new Uint8Array([9]).buffer);
      return value;
    });
    await expect(runMigration(source, target, plan)).rejects.toThrow('Document changed');
  });
  it('blocks changed destination records before writes and rejects ambiguous identities', async () => {
    const { source, target } = fixture();
    target.rows.projects = structuredClone(source.rows.projects);
    const plan = await previewMigration(source, target);
    target.rows.projects[0].Title = 'Concurrent change';
    await expect(runMigration(source, target, plan)).rejects.toThrow('Destination changed');
    expect(target.create).not.toHaveBeenCalled();
    source.rows.projects.push(record('other-project', { Id: 3, ProjectKey: 'project-key' }));
    await expect(previewMigration(source, target)).rejects.toThrow('duplicate logical identity');
  });
  it('rejects orphan tasks, cyclic folders, and the wrong destination', async () => {
    const { source, target } = fixture();
    source.rows.tasks[0].ProjectKey = 'missing';
    await expect(previewMigration(source, target)).rejects.toThrow('project missing');
    source.rows.tasks[0].ProjectKey = 'project-key';
    source.rows.references[0].ReferenceParentId = 'folder';
    await expect(previewMigration(source, target)).rejects.toThrow('folder relationships');
    target.webUrl = MIGRATION_SOURCE;
    await expect(previewMigration(source, target)).rejects.toThrow('only supports');
  });
  it('requires actual SharePoint owner permissions at preview and again at execution', async () => {
    const { source, target } = fixture();
    target.canMigrate.mockResolvedValue(false);
    await expect(previewMigration(source, target)).rejects.toThrow('Manage Permissions');
    target.canMigrate.mockResolvedValue(true);
    const plan = await previewMigration(source, target);
    target.canMigrate.mockResolvedValue(false);
    await expect(runMigration(source, target, plan)).rejects.toThrow('Manage Permissions');
    expect(target.create).not.toHaveBeenCalled();
  });
});

describe('SharePoint migration adapter', () => {
  const response = body => new Response(JSON.stringify(body), { headers: { 'Content-Type': 'application/json' } });
  it('distinguishes site members from owners', async () => {
    const fetcher = vi.fn().mockResolvedValueOnce(response({ Low: '2048' })).mockResolvedValueOnce(response({ Low: String(2048 + 33554432) }));
    const site = migrationSite(MIGRATION_TARGET, 'Modernization', fetcher);
    expect(await site.canMigrate()).toBe(false);
    expect(await site.canMigrate()).toBe(true);
  });
  it('reads every pagination page without filtering archived records and rejects foreign pagination', async () => {
    const fetcher = vi.fn().mockResolvedValueOnce(response({ value: [{ Id: 1 }], 'odata.nextLink': `${MIGRATION_SOURCE}/_api/next` })).mockResolvedValueOnce(response({ value: [{ Id: 2, Archived: true }] }));
    const site = migrationSite(MIGRATION_SOURCE, 'Modernization', fetcher);
    expect(await site.readRows('tasks', ['RecordId', 'Archived'])).toEqual([{ Id: 1 }, { Id: 2, Archived: true }]);
    expect(fetcher.mock.calls[0][0]).not.toContain('$filter');
    fetcher.mockResolvedValueOnce(response({ value: [], 'odata.nextLink': 'https://example.invalid/_api/next' }));
    await expect(site.readRows('tasks', ['RecordId'])).rejects.toThrow('Unexpected pagination');
  });
  it('refuses an existing-record update without an ETag', async () => {
    const fetcher = vi.fn();
    await expect(migrationSite(MIGRATION_TARGET, 'Modernization', fetcher).update('tasks', 1, {}, '')).rejects.toThrow('concurrency token');
    expect(fetcher).not.toHaveBeenCalled();
  });
});
