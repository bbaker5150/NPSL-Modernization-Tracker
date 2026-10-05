import { CONTAINERS } from './spStore';
import { MIGRATION_SOURCE, MIGRATION_TARGET, fieldsFor, writable, same, sha256, previewMigration } from './siteMigration';
import { reconcileMigrationWrite } from './migrationRecovery';
import { validateAttachmentName } from './taskAttachments';

export const MAX_BACKUP_BYTES = 256 * 1024 * 1024;
const format = 'npsl-tracker-backup';
const encoder = new TextEncoder();
const urlEquals = (a, b) => String(a).replace(/\/+$/, '').toLowerCase() === b.toLowerCase();
const names = row => {
  if (row.AttachmentFiles?.__next || row['AttachmentFiles@odata.nextLink'] || row['AttachmentFiles@odata.nextLinkUrl']) throw new Error('Attachment listing was incomplete. Export stopped instead of omitting documents.');
  return (row.AttachmentFiles?.results || row.AttachmentFiles || []).map(file => file.FileName);
};
const fingerprint = value => sha256(encoder.encode(JSON.stringify(value)));
function base64(bytes) {
  const view = new Uint8Array(bytes); let value = '';
  for (let i = 0; i < view.length; i += 8192) value += String.fromCharCode(...view.subarray(i, i + 8192));
  return btoa(value);
}
function decode(value) { return Uint8Array.from(atob(value), char => char.charCodeAt(0)).buffer; }
function fileKey(key, recordId, name) { return JSON.stringify([key, recordId, name.toLowerCase()]); }

function backupScope(backup) {
  if (backup.version === 1) return { includeReferences: true, includeTaskAttachments: true };
  const scope = backup.payload?.scope;
  if (!scope || typeof scope.includeReferences !== 'boolean' || typeof scope.includeTaskAttachments !== 'boolean') throw new Error('Backup scope is missing or invalid. Export a fresh package.');
  return scope;
}
export function manualTransferItems(backup) {
  const scope = backupScope(backup);
  return [...(scope.includeReferences ? [] : ['Reference Documents and folders']), ...(scope.includeTaskAttachments ? [] : ['task attachments'])];
}

export async function exportTrackerBackup(site, onProgress = () => {}, { includeReferences = false, includeTaskAttachments = true } = {}) {
  if (!urlEquals(site.webUrl, MIGRATION_SOURCE) || !await site.canMigrate()) throw new Error('Export requires SharePoint owner permissions on ISEA METENG.');
  const payload = { source: MIGRATION_SOURCE, createdAt: new Date().toISOString(), scope: { includeReferences, includeTaskAttachments }, lists: [], files: [] };
  let bytesUsed = 0;
  for (const container of CONTAINERS.filter(container => includeReferences || container.key !== 'references')) {
    onProgress(`Exporting ${container.suffix} records…`);
    const schema = await site.schema(container);
    const fields = fieldsFor(container).filter(field => schema.has(field));
    const attached = container.key === 'references' || (container.key === 'tasks' && includeTaskAttachments);
    const rows = await site.readRows(container.key, fields, attached);
    const records = rows.map(row => ({ Id: row.Id, Modified: row.Modified, ...writable(row, fields) }));
    payload.lists.push({ key: container.key, fields, records });
    for (const row of rows) {
      for (const name of attached ? names(row) : []) {
        onProgress(`Exporting document ${payload.files.length + 1}: ${name}…`);
        const bytes = await site.bytes(container.key, row.Id, name);
        bytesUsed += bytes.byteLength;
        if (bytesUsed > 150 * 1024 * 1024) throw new Error('Documents exceed the 150 MB browser backup limit. No incomplete package was downloaded.');
        payload.files.push({ list: container.key, recordId: row.RecordId, name, size: bytes.byteLength, sha256: await sha256(bytes), data: base64(bytes) });
      }
    }
    // One list read catches edits during export; no per-record verification calls.
    const after = await site.readRows(container.key, fields, attached);
    if (after.length !== rows.length || rows.some(row => !after.some(now => now.RecordId === row.RecordId && now.Modified === row.Modified && same(now, row, fields) && JSON.stringify(names(now).sort()) === JSON.stringify(names(row).sort())))) throw new Error(`${container.suffix} changed during export. Pause source edits and export again.`);
  }
  const backup = { format, version: 2, payload, checksum: await fingerprint(payload) };
  await validateTrackerBackup(backup);
  const text = JSON.stringify(backup);
  if (encoder.encode(text).byteLength > MAX_BACKUP_BYTES) throw new Error('Backup exceeds the 256 MB browser package limit.');
  return text;
}

export async function validateTrackerBackup(backup) {
  if (backup?.format !== format || ![1, 2].includes(backup.version) || !urlEquals(backup.payload?.source, MIGRATION_SOURCE)) throw new Error('Choose a supported NPSL tracker backup exported from ISEA METENG.');
  const { payload } = backup;
  const scope = backupScope(backup);
  const expected = CONTAINERS.filter(container => scope.includeReferences || container.key !== 'references');
  if (!Array.isArray(payload.lists) || payload.lists.length !== expected.length || !Array.isArray(payload.files)) throw new Error('Backup is missing tracker lists or documents.');
  if (backup.checksum !== await fingerprint(payload)) throw new Error('Backup checksum failed. Export a fresh package; no data was imported.');
  const recordMaps = new Map(), seen = new Set();
  for (const list of payload.lists) {
    const container = expected.find(c => c.key === list.key);
    if (!container || seen.has(list.key) || !Array.isArray(list.fields) || !Array.isArray(list.records)) throw new Error('Invalid or duplicated backup list.');
    seen.add(list.key);
    const allowed = fieldsFor(container);
    if (new Set(list.fields).size !== list.fields.length || !['Title', 'RecordId', 'Archived'].every(field => list.fields.includes(field)) || list.fields.some(field => !allowed.includes(field))) throw new Error(`Invalid field schema: ${list.key}.`);
    const byId = new Map(), ids = new Set();
    for (const row of list.records) {
      if (!row || !Number.isSafeInteger(row.Id) || row.Id <= 0 || ids.has(row.Id) || typeof row.RecordId !== 'string' || !row.RecordId || byId.has(row.RecordId)) throw new Error(`Invalid or duplicate record in ${list.key}.`);
      ids.add(row.Id); byId.set(row.RecordId, row);
      for (const field of list.fields) {
        const value = row[field], type = container.fields.find(f => f.name === field)?.type || 'Text';
        if (value == null) continue;
        if (type === 'Boolean' ? typeof value !== 'boolean' : type === 'Number' ? typeof value !== 'number' || !Number.isFinite(value) : typeof value !== 'string' || (type === 'DateTime' && value && !Number.isFinite(Date.parse(value)))) throw new Error(`Invalid ${list.key}.${field} value.`);
      }
    }
    recordMaps.set(list.key, byId);
  }
  const files = new Set(); let totalBytes = 0;
  for (const file of payload.files) {
    if (file.list === 'tasks' && !scope.includeTaskAttachments) throw new Error('Backup contains documents excluded by its scope.');
    if (!['tasks', 'references'].includes(file.list) || !recordMaps.get(file.list)?.has(file.recordId)) throw new Error('Document refers to a missing record.');
    validateAttachmentName(file.name);
    const key = fileKey(file.list, file.recordId, file.name);
    if (files.has(key)) throw new Error('Duplicate document in backup.');
    files.add(key);
    if (typeof file.data !== 'string' || file.data.length % 4 !== 0 || /[^A-Za-z0-9+/=]/.test(file.data)) throw new Error('Invalid document encoding.');
    const bytes = decode(file.data); totalBytes += bytes.byteLength;
    if (totalBytes > 150 * 1024 * 1024 || file.size !== bytes.byteLength || file.sha256 !== await sha256(bytes)) throw new Error(`Document integrity check failed: ${file.name}.`);
  }
  for (const row of recordMaps.get('references')?.values() || []) if (row.EntryKind === 'file' && !files.has(fileKey('references', row.RecordId, row.FileName || ''))) throw new Error(`Reference document is missing: ${row.Title}.`);
  // Run the existing relationship and identity checks entirely in memory.
  const source = backupSource(backup);
  await previewMigration(source, { ...source, webUrl: MIGRATION_TARGET, canMigrate: async () => true, readRows: async () => [] }, undefined, scope);
  return backup;
}

function backupSource(backup) {
  const { lists, files } = backup.payload;
  return {
    webUrl: MIGRATION_SOURCE,
    schema: async container => new Set(lists.find(list => list.key === container.key).fields),
    readRows: async key => lists.find(list => list.key === key).records,
    attachments: async (key, id) => {
      const row = lists.find(list => list.key === key).records.find(row => row.Id === id);
      return files.filter(file => file.list === key && file.recordId === row.RecordId).map(file => file.name);
    },
  };
}

export async function previewTrackerImport(backup, target, onProgress = () => {}) {
  await validateTrackerBackup(backup);
  const plan = await previewMigration(backupSource(backup), target, onProgress, backupScope(backup));
  return { ...plan, backupChecksum: backup.checksum };
}

export async function importTrackerBackup(backup, target, plan, { replaceConflicts = false, onProgress = () => {}, wait } = {}) {
  const fresh = await previewTrackerImport(backup, target, onProgress);
  if (plan.backupChecksum !== backup.checksum) throw new Error('Backup changed after preview. Preview again.');
  for (const list of fresh.lists) {
    const reviewed = plan.lists.find(value => value.key === list.key);
    if (!reviewed || list.entries.some((entry, index) => {
      const before = reviewed.entries[index];
      return !before || entry.destination?.Id !== before.destination?.Id || entry.destination?.Modified !== before.destination?.Modified || !!entry.destination && !same(entry.destination, before.destination, list.fields);
    })) throw new Error('Destination changed after preview. Preview again.');
  }
  if (!replaceConflicts && fresh.lists.some(list => list.entries.some(entry => entry.conflict))) throw new Error('Review and approve conflicting destination records before importing.');
  const scope = backupScope(backup);
  const report = { source: backup.payload.source, target: target.webUrl, backupChecksum: backup.checksum, scope, manualTransfer: manualTransferItems(backup), startedAt: new Date().toISOString(), verified: false, lists: [], files: [] };
  const targetIds = new Map();
  for (const list of fresh.lists) {
    const pending = list.entries.filter(entry => !entry.destination || entry.conflict);
    for (let start = 0; start < pending.length; start += 25) {
      const batch = pending.slice(start, start + 25);
      onProgress(`Importing ${list.title}: ${Math.min(start + 25, pending.length)}/${pending.length} records…`);
      const operations = batch.map(entry => ({ id: entry.destination?.Id, etag: entry.destination?.['odata.etag'] || entry.destination?.['@odata.etag'] || entry.destination?.etag, fields: writable(entry.source, list.fields) }));
      await reconcileMigrationWrite(() => target.writeBatch(list.key, operations), async () => {
        const saved = await target.readRows(list.key, list.fields);
        return batch.every(entry => saved.filter(row => row.RecordId === entry.source.RecordId && same(row, entry.source, list.fields)).length === 1) ? { value: undefined } : null;
      }, { wait, onProgress, label: `${list.title} batch ${Math.floor(start / 25) + 1}` });
    }
    // Verify list values once and obtain the new SharePoint IDs for documents.
    onProgress(`Verifying ${list.title} records…`);
    const attached = list.key === 'references' || (list.key === 'tasks' && scope.includeTaskAttachments);
    const saved = await target.readRows(list.key, list.fields, attached);
    for (const entry of list.entries) {
      const matches = saved.filter(row => row.RecordId === entry.source.RecordId);
      if (matches.length !== 1 || !same(matches[0], entry.source, list.fields)) throw new Error(`${list.title}/${entry.source.RecordId} did not verify. Preview again to resume.`);
      targetIds.set(fileKey(list.key, entry.source.RecordId, ''), { id: matches[0].Id, names: names(matches[0]) });
    }
    report.lists.push({ list: list.title, records: list.entries.length, written: pending.length, matching: list.entries.length - pending.length, destinationExtrasRetained: list.extra });
  }
  for (const [index, file] of backup.payload.files.entries()) {
    onProgress(`Importing document ${index + 1}/${backup.payload.files.length}: ${file.name}…`);
    const destination = targetIds.get(fileKey(file.list, file.recordId, ''));
    const matching = destination.names.find(name => name.toLowerCase() === file.name.toLowerCase());
    const bytes = decode(file.data);
    if (!matching) await reconcileMigrationWrite(() => target.attach(file.list, destination.id, file.name, bytes), async () => {
      const names = await target.attachments(file.list, destination.id);
      const found = names.find(name => name.toLowerCase() === file.name.toLowerCase());
      if (!found) return null;
      if (await sha256(await target.bytes(file.list, destination.id, found)) !== file.sha256) throw new Error(`Different destination document: ${file.name}. No file was overwritten.`);
      return { value: undefined };
    }, { wait, onProgress, label: file.name });
    if (await sha256(await target.bytes(file.list, destination.id, matching || file.name)) !== file.sha256) throw new Error(`Document differs from backup: ${file.name}. No existing file was overwritten.`);
    report.files.push({ list: file.list, recordId: file.recordId, name: file.name, size: file.size, sha256: file.sha256 });
  }
  report.verified = true; report.completedAt = new Date().toISOString();
  return report;
}
