import { vi } from 'vitest';
import { CONTAINERS } from './spStore';
import { MIGRATION_SOURCE, MIGRATION_TARGET } from './siteMigration';
export function memorySite(webUrl, seed = {}) {
  const rows = Object.fromEntries(CONTAINERS.map(c => [c.key, structuredClone(seed[c.key] || [])]));
  const files = new Map(); let nextId = 100;
  const fileKey = (key, id, name) => `${key}/${id}/${name}`;
  const api = {
    webUrl, rows, files,
    canMigrate: vi.fn(async () => true),
    schema: vi.fn(async container => new Set(['Title', ...container.fields.map(f => f.name)])),
    readRows: vi.fn(async (key, fields, attached) => structuredClone(rows[key]).map(row => ({ ...row, 'odata.etag': '\"1\"', ...(attached ? { AttachmentFiles: [...files.keys()].filter(path => path.startsWith(`${key}/${row.Id}/`)).map(path => ({ FileName: path.slice(`${key}/${row.Id}/`.length) })) } : {}) }))),
    readItem: vi.fn(async (key, id) => ({ ...structuredClone(rows[key].find(row => row.Id === id)), etag: '"1"' })),
    create: vi.fn(async (key, fields) => { const Id = nextId++; rows[key].push({ ...fields, Id, Modified: 'new' }); return Id; }),
    update: vi.fn(async (key, id, fields) => Object.assign(rows[key].find(row => row.Id === id), fields, { Modified: 'updated' })),
    attachments: vi.fn(async (key, id) => [...files.keys()].filter(path => path.startsWith(`${key}/${id}/`)).map(path => path.slice(`${key}/${id}/`.length))),
    bytes: vi.fn(async (key, id, name) => { const bytes = files.get(fileKey(key, id, name)); if (!bytes) throw new Error('Missing file'); return bytes.slice(0); }),
    attach: vi.fn(async (key, id, name, bytes) => files.set(fileKey(key, id, name), bytes.slice(0))),
  };
  return api;
}
export const record = (RecordId, fields = {}) => ({ Id: 1, RecordId, Title: RecordId, Archived: false, Modified: 'original', ...fields });
export function fixture() {
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

