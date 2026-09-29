import { describe, it, expect, vi } from 'vitest';
import { validateReference, referencePath } from './referenceDocuments';
import { authorizedStore } from './access';
import { SharePointStore } from './spStore';

const entries = [
  { id: 'a', name: 'Templates', kind: 'folder', parentId: '' },
  { id: 'b', name: 'Acquisition', kind: 'folder', parentId: 'a' },
  { id: 'doc', name: 'Template.docx', fileName: 'Template.docx', kind: 'file', parentId: 'b', spId: 3 },
];
describe('reference document organization', () => {
  it('rejects duplicate names, missing parents, cycles, and missing records', () => {
    expect(() => validateReference(entries, { name: 'templates', parentId: '' })).toThrow('already exists');
    expect(() => validateReference(entries, { name: 'New', parentId: 'missing' })).toThrow('existing folder');
    expect(() => validateReference(entries, { id: 'a', name: 'Templates', parentId: 'b' })).toThrow('inside itself');
    expect(() => validateReference(entries, { id: 'missing', name: 'New' })).toThrow('no longer exists');
    expect(referencePath(entries, 'b').map((row) => row.name)).toEqual(['Templates', 'Acquisition']);
    expect(validateReference(entries, { id: 'doc', name: 'Updated.docx', parentId: '' })).toMatchObject({ spId: 3, fileName: 'Template.docx', name: 'Updated.docx', parentId: '' });
  });
  it('blocks SME mutations but permits download; user mutations are validated', async () => {
    const raw = { currentUser: async () => ({ loginName: 'reader' }), load: async () => ({ users: [{ loginName: 'reader', role: 'SME' }] }), listReferenceEntries: async () => entries, saveReferenceEntry: vi.fn(), downloadReferenceEntry: vi.fn() };
    const store = authorizedStore(raw);
    await expect(store.saveReferenceEntry({ name: 'Folder' })).rejects.toThrow('read-only');
    await store.downloadReferenceEntry('doc');
    expect(raw.downloadReferenceEntry).toHaveBeenCalledWith(entries[2]);
    await expect(store.downloadReferenceEntry('a')).rejects.toThrow('no longer exists');
    raw.load = async () => ({ users: [] });
    await store.saveReferenceEntry({ name: 'New', parentId: '' });
    expect(raw.saveReferenceEntry).toHaveBeenCalledWith({ name: 'New', parentId: '', kind: 'folder' }, undefined);
  });
  it('keeps a failed SharePoint upload archived and only exposes successful uploads', async () => {
    const store = new SharePointStore({ webUrl: 'https://example.sharepoint.com/sites/mod' });
    store.create = vi.fn(async () => 8);
    store.update = vi.fn();
    store.post = vi.fn().mockRejectedValueOnce(new Error('Upload failed')).mockResolvedValue({});
    store.listReferenceEntries = vi.fn(async () => [{ id: 'saved', name: 'Template.docx', parentId: '' }]);
    const file = { name: 'Template.docx', arrayBuffer: async () => new ArrayBuffer(2) };
    const row = { id: 'saved', name: file.name, fileName: file.name, parentId: '', kind: 'file', size: 2 };
    await expect(store.saveReferenceEntry(row, file)).rejects.toThrow('Upload failed');
    expect(store.create).toHaveBeenCalledWith('references', expect.objectContaining({ Archived: true }));
    expect(store.update).not.toHaveBeenCalled();
    await store.saveReferenceEntry(row, file);
    expect(store.update).toHaveBeenCalledWith('references', 8, { Archived: false });
  });
  it('downloads reference bytes without navigating to SharePoint', async () => {
    const blob = new Blob(['template']);
    const fetchImpl = vi.fn(async () => ({ ok: true, blob: async () => blob }));
    const store = new SharePointStore({ webUrl: 'https://example.sharepoint.com/sites/mod', fetchImpl });
    store.get = vi.fn(async () => ({ value: [{ FileName: 'Template.docx', ServerRelativeUrl: '/sites/mod/Lists/References/Attachments/3/Template.docx' }] }));
    expect(await store.downloadReferenceEntry(entries[2])).toBe(blob);
    expect(fetchImpl).toHaveBeenCalledWith(expect.stringContaining('/Attachments/3/Template.docx'), expect.objectContaining({ credentials: 'include' }));
  });
});
