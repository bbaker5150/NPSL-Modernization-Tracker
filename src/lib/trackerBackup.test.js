// @vitest-environment node
import { expect, it, vi } from 'vitest';
import { fixture, record } from './backupTestFixtures';
import { exportTrackerBackup, validateTrackerBackup, previewTrackerImport, importTrackerBackup } from './trackerBackup';
import { migrationSite, sha256 } from './siteMigration';
import { makeMigrationBatch, checkMigrationBatch } from './migrationBatch';

function batchTarget(target) {
  target.writeBatch = vi.fn(async (key, operations) => {
    for (const op of operations) if (op.id) await target.update(key, op.id, op.fields, op.etag); else await target.create(key, op.fields);
  });
  return target;
}
it('exports all seven lists and binary attachments, then imports from the package with no source connection', async () => {
  const { source, target } = fixture(); batchTarget(target);
  source.rows.users.unshift(record('archived-user', { Id: 9, LoginKey: 'owner@example.invalid', Archived: true, AppRole: 'Viewer' }));
  const text = await exportTrackerBackup(source, undefined, { includeReferences: true });
  expect(source.readRows).toHaveBeenCalledTimes(14);
  expect(source.attachments).not.toHaveBeenCalled();
  const backup = JSON.parse(text);
  expect(backup.payload.files).toHaveLength(2);
  expect(backup.payload.lists).toHaveLength(7);
  for (const key of ['readRows', 'schema', 'bytes', 'attachments']) source[key].mockRejectedValue(new Error('Source is offline'));
  const report = await importTrackerBackup(backup, target, await previewTrackerImport(backup, target));
  expect(report.verified).toBe(true);
  expect(report.files).toHaveLength(2);
  expect(target.readItem).not.toHaveBeenCalled();
  expect(target.rows.tasks[0]).toMatchObject({ Archived: true, ArchivedDocuments: '["hidden.pdf"]', ProjectKey: 'project-key' });
  expect(target.rows.references[1].ReferenceParentId).toBe('folder');
  expect(target.rows.users).toHaveLength(2);
  expect(target.writeBatch).toHaveBeenCalledTimes(7);
  const writes = target.writeBatch.mock.calls.length;
  await importTrackerBackup(backup, target, await previewTrackerImport(backup, target));
  expect(target.writeBatch).toHaveBeenCalledTimes(writes);
  expect(target.attach).toHaveBeenCalledTimes(2);
  expect(source.create).not.toHaveBeenCalled();
});
it('groups 61 tasks into batches of 25, 25 and 11', async () => {
  const { source, target } = fixture(); batchTarget(target);
  for (let id = 2; id <= 61; id++) source.rows.tasks.push(record(`task-${id}`, { Id: id, ProjectKey: 'project-key' }));
  const backup = JSON.parse(await exportTrackerBackup(source, undefined, { includeReferences: true }));
  await importTrackerBackup(backup, target, await previewTrackerImport(backup, target));
  expect(target.writeBatch.mock.calls.filter(([key]) => key === 'tasks').map(([, ops]) => ops.length)).toEqual([25, 25, 11]);
});
it('resumes a partially committed batch without duplicates', async () => {
  const { source, target } = fixture(); batchTarget(target);
  source.rows.projects.push(record('p2', { Id: 2, ProjectKey: 'second' }));
  const backup = JSON.parse(await exportTrackerBackup(source, undefined, { includeReferences: true }));
  target.writeBatch.mockImplementationOnce(async (key, operations) => {
    await target.create(key, operations[0].fields);
    throw new Error('SharePoint request timed out');
  });
  await expect(importTrackerBackup(backup, target, await previewTrackerImport(backup, target), { wait: vi.fn() })).rejects.toThrow('not sent again');
  expect(target.rows.projects).toHaveLength(1);
  expect((await importTrackerBackup(backup, target, await previewTrackerImport(backup, target))).verified).toBe(true);
  expect(target.rows.projects).toHaveLength(2);
});
it('requires reviewed conflict approval and sends the destination ETag', async () => {
  const { source, target } = fixture(); batchTarget(target);
  target.rows.users.push(record('bootstrap', { Id: 6, LoginKey: 'owner@example.invalid', AppRole: 'Viewer' }));
  const backup = JSON.parse(await exportTrackerBackup(source, undefined, { includeReferences: true })), plan = await previewTrackerImport(backup, target);
  await expect(importTrackerBackup(backup, target, plan)).rejects.toThrow('approve');
  expect(target.writeBatch).not.toHaveBeenCalled();
  await importTrackerBackup(backup, target, plan, { replaceConflicts: true });
  expect(target.writeBatch.mock.calls.find(([key]) => key === 'users')[1][0]).toMatchObject({ id: 6, etag: '"1"' });
});
it('rejects corrupted packages before writes and preserves different existing documents', async () => {
  const { source, target } = fixture(); batchTarget(target);
  const backup = JSON.parse(await exportTrackerBackup(source, undefined, { includeReferences: true }));
  const changed = structuredClone(backup); changed.payload.lists[0].records[0].Title = 'Changed';
  await expect(validateTrackerBackup(changed)).rejects.toThrow('checksum');
  target.rows.tasks = structuredClone(source.rows.tasks);
  target.files.set('tasks/1/hidden.pdf', new Uint8Array([99]).buffer);
  await expect(importTrackerBackup(backup, target, await previewTrackerImport(backup, target))).rejects.toThrow('Document differs');
  expect(target.attach).not.toHaveBeenCalled();
  expect(new Uint8Array(target.files.get('tasks/1/hidden.pdf'))).toEqual(new Uint8Array([99]));
});
it('detects edits after preview and preserves destination-only records', async () => {
  const { source, target } = fixture(); batchTarget(target);
  target.rows.projects.push(record('unrelated', { Id: 3, ProjectKey: 'unrelated' }));
  target.rows.users.push(record('bootstrap', { Id: 4, LoginKey: 'owner@example.invalid', AppRole: 'Viewer' }));
  const backup = JSON.parse(await exportTrackerBackup(source, undefined, { includeReferences: true })), plan = await previewTrackerImport(backup, target);
  target.rows.users[0].Title = 'Changed';
  await expect(importTrackerBackup(backup, target, plan, { replaceConflicts: true })).rejects.toThrow('Destination changed');
  expect(target.writeBatch).not.toHaveBeenCalled();
  await importTrackerBackup(backup, target, await previewTrackerImport(backup, target), { replaceConflicts: true });
  expect(target.rows.projects.some(row => row.RecordId === 'unrelated')).toBe(true);
});
it('builds same-site multipart writes with ETags and checks every response status', () => {
  const batch = makeMigrationBatch('https://example.invalid/sites/app', "/_api/web/lists/getbytitle('Tasks')", [{ fields: { Title: 'A' } }, { id: 2, etag: '"3"', fields: { Title: 'B' } }]);
  expect(batch.body).toContain('POST https://example.invalid/sites/app/_api/web/lists');
  expect(batch.body).toContain('X-HTTP-Method: MERGE\r\nIf-Match: "3"');
  expect(batch.contentType).toMatch(/^multipart\/mixed; boundary=batch_/);
  expect(() => makeMigrationBatch('https://example.invalid', '/_api/list', [{ id: 2, fields: {} }])).toThrow('ETag');
  expect(() => checkMigrationBatch('--batch\r\nHTTP/1.1 201 Created\r\n\r\n{}\r\n--batch\r\nHTTP/1.1 204 No Content\r\n', 2)).not.toThrow();
  expect(() => checkMigrationBatch('HTTP/1.1 201 Created\r\nHTTP/1.1 403 Denied\r\n', 2)).toThrow('every batch');
  expect(() => checkMigrationBatch('HTTP/1.1 201 Created\r\n', 2)).toThrow('every batch');
});

it('rejects a corrupt document even if the outer package checksum was recalculated', async () => {
  const { source } = fixture();
  const backup = JSON.parse(await exportTrackerBackup(source, undefined, { includeReferences: true }));
  backup.payload.files[0].data = btoa('corrupted');
  backup.checksum = await sha256(new TextEncoder().encode(JSON.stringify(backup.payload)));
  await expect(validateTrackerBackup(backup)).rejects.toThrow('Document integrity');
});
it('requires SharePoint owner permissions for export and import', async () => {
  const { source, target } = fixture(); batchTarget(target);
  const backup = JSON.parse(await exportTrackerBackup(source, undefined, { includeReferences: true }));
  const plan = await previewTrackerImport(backup, target);
  source.canMigrate = async () => false;
  target.canMigrate = async () => false;
  await expect(exportTrackerBackup(source, undefined, { includeReferences: true })).rejects.toThrow('owner permissions');
  await expect(importTrackerBackup(backup, target, plan)).rejects.toThrow('Manage Permissions');
  expect(target.writeBatch).not.toHaveBeenCalled();
});
it('sends authenticated multipart batches and rejects an inner failure without retrying writes', async () => {
  const fetcher = vi.fn()
    .mockResolvedValueOnce(new Response(JSON.stringify({ FormDigestValue: 'test-digest', FormDigestTimeoutSeconds: 1800 })))
    .mockResolvedValueOnce(new Response('HTTP/1.1 201 Created\r\n\r\n{}\r\nHTTP/1.1 403 Denied\r\n', { headers: { 'Content-Type': 'multipart/mixed; boundary=test' } }));
  const site = migrationSite('https://example.invalid/sites/batch-adapter', 'Modernization', fetcher);
  await expect(site.writeBatch('tasks', [{ fields: { Title: 'A' } }, { fields: { Title: 'B' } }])).rejects.toThrow('every batch');
  expect(fetcher).toHaveBeenCalledTimes(2);
  expect(fetcher.mock.calls[1][0]).toBe('https://example.invalid/sites/batch-adapter/_api/$batch');
  expect(fetcher.mock.calls[1][1]).toMatchObject({ method: 'POST', credentials: 'include', headers: { 'X-RequestDigest': 'test-digest' } });
  expect(fetcher.mock.calls[1][1].body).toContain("/getbytitle('ModernizationTasks')/items HTTP/1.1");
});

it('excludes Reference Documents completely by default and leaves destination references untouched', async () => {
  const { source, target } = fixture(); batchTarget(target);
  const schema = source.schema.getMockImplementation(), read = source.readRows.getMockImplementation(), bytes = source.bytes.getMockImplementation();
  source.schema.mockImplementation(container => { if (container.key === 'references') throw new Error('References must not be accessed'); return schema(container); });
  source.readRows.mockImplementation((key, ...args) => { if (key === 'references') throw new Error('References must not be read'); return read(key, ...args); });
  source.bytes.mockImplementation((key, ...args) => { if (key === 'references') throw new Error('Reference download hangs'); return bytes(key, ...args); });
  target.rows.references.push(record('existing-reference', { EntryKind: 'folder' }));
  const existing = structuredClone(target.rows.references);
  const backup = JSON.parse(await exportTrackerBackup(source));
  expect(backup.version).toBe(2);
  expect(backup.payload.scope).toEqual({ includeReferences: false, includeTaskAttachments: true });
  expect(backup.payload.lists).toHaveLength(6);
  expect(backup.payload.files.map(file => file.list)).toEqual(['tasks']);
  const report = await importTrackerBackup(backup, target, await previewTrackerImport(backup, target));
  expect(report.verified).toBe(true);
  expect(report.manualTransfer).toEqual(['Reference Documents and folders']);
  expect(target.rows.references).toEqual(existing);
  expect(target.schema.mock.calls.some(([container]) => container.key === 'references')).toBe(false);
  expect(target.readRows.mock.calls.some(([key]) => key === 'references')).toBe(false);
  expect(target.writeBatch.mock.calls.some(([key]) => key === 'references')).toBe(false);
  expect(target.attach.mock.calls.map(([key]) => key)).toEqual(['tasks']);
});
it('can export and import all six record lists without any attachment requests', async () => {
  const { source, target } = fixture(); batchTarget(target);
  for (const site of [source, target]) for (const key of ['bytes', 'attachments', 'attach']) site[key].mockImplementation(() => { throw new Error('No attachment operation is allowed'); });
  const backup = JSON.parse(await exportTrackerBackup(source, undefined, { includeTaskAttachments: false }));
  expect(backup.payload.files).toEqual([]);
  expect(source.readRows.mock.calls.every(([, , attached]) => attached === false)).toBe(true);
  const report = await importTrackerBackup(backup, target, await previewTrackerImport(backup, target));
  expect(report.verified).toBe(true);
  expect(report.manualTransfer).toEqual(['Reference Documents and folders', 'task attachments']);
  expect(target.rows.tasks[0].ArchivedDocuments).toBe('["hidden.pdf"]');
  for (const site of [source, target]) for (const key of ['bytes', 'attachments', 'attach']) expect(site[key]).not.toHaveBeenCalled();
  await importTrackerBackup(backup, target, await previewTrackerImport(backup, target));
  expect(target.writeBatch).toHaveBeenCalledTimes(6);
});
it('still imports complete version 1 backups and rejects undeclared omissions', async () => {
  const { source, target } = fixture(); batchTarget(target);
  const backup = JSON.parse(await exportTrackerBackup(source, undefined, { includeReferences: true }));
  backup.version = 1; delete backup.payload.scope;
  backup.checksum = await sha256(new TextEncoder().encode(JSON.stringify(backup.payload)));
  expect((await importTrackerBackup(backup, target, await previewTrackerImport(backup, target))).manualTransfer).toEqual([]);
  expect(target.rows.references).toHaveLength(2);
  backup.payload.lists = backup.payload.lists.filter(list => list.key !== 'references');
  backup.checksum = await sha256(new TextEncoder().encode(JSON.stringify(backup.payload)));
  await expect(validateTrackerBackup(backup)).rejects.toThrow('missing tracker lists');
});
it('requires valid explicit scope and rejects task files declared excluded', async () => {
  const { source } = fixture();
  const backup = JSON.parse(await exportTrackerBackup(source));
  backup.payload.scope.includeTaskAttachments = false;
  backup.checksum = await sha256(new TextEncoder().encode(JSON.stringify(backup.payload)));
  await expect(validateTrackerBackup(backup)).rejects.toThrow('excluded by its scope');
  delete backup.payload.scope;
  await expect(validateTrackerBackup(backup)).rejects.toThrow('scope is missing');
});
