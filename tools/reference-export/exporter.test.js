// @vitest-environment node
import { describe, it, expect, vi } from 'vitest';
import JSZip from 'jszip';
import { SOURCE, sourceUrl, readSource, scanReferences, planReferences, acceptFile, makeZip } from './exporter.js';
const list = { Id: '11111111-1111-1111-1111-111111111111', Title: 'Modernization-Tracker - Reference Documents', BaseTemplate: 100 };
const root = '/sites/ISEAMETENG/Lists/ModernizationReferenceDocuments';
const folder = { Id: 1, RecordId: 'folder', Title: 'Templates', EntryKind: 'folder', Modified: '2026-01-01', AttachmentFiles: [] };
const file = { Id: 2, RecordId: 'file', Title: 'Sample.docx', EntryKind: 'file', ReferenceParentId: 'folder', FileName: 'Original #50%.docx', FileSize: 3, Modified: '2026-01-01', AttachmentFiles: [{ FileName: 'Original #50%.docx' }] };
const makePlan = rows => planReferences(rows || [folder, file], list, root);

describe('one-time reference export', () => {
  it('keeps source requests in the old site and rejects foreign pagination/redirects', async () => {
    expect(() => sourceUrl('https://example.com/file')).toThrow('outside');
    expect(() => sourceUrl('/sites/metsoft/_api/web')).toThrow('outside');
    expect(() => sourceUrl(`${SOURCE}/../metsoft/_api/web`)).toThrow('outside');
    await expect(readSource(`${SOURCE}/file`, { fetchImpl: async () => ({ ok: true, url: 'https://example.com/login', json: async () => ({}) }) })).rejects.toThrow('outside');
  });
  it('scans all pages with GET only and resolves the list ID before reading files', async () => {
    const fetchImpl = vi.fn(async url => new Response(JSON.stringify(url.includes('/lists?') ? { value: [list] } : url.includes('/RootFolder?') ? { ServerRelativeUrl: root } : url.includes('page=2') ? { value: [file] } : { value: [folder], '@odata.nextLink': `${SOURCE}/_api/page=2` })));
    const plan = await scanReferences({ fetchImpl });
    expect(plan.files).toHaveLength(1);
    expect(plan.files[0].zipPath).toBe('Active/Templates/Sample.docx');
    expect(plan.files[0].url).toBe(`${SOURCE}/Lists/ModernizationReferenceDocuments/Attachments/2/Original%20%2350%25.docx`);
    expect(fetchImpl.mock.calls.every(([, options]) => options.method === 'GET')).toBe(true);
    expect(fetchImpl.mock.calls[0][0]).toContain('Modernization-Tracker');
  });
  it('settles timeouts and cancellation even when the host ignores AbortSignal', async () => {
    await expect(readSource(`${SOURCE}/file`, { timeoutMs: 10, fetchImpl: () => new Promise(() => {}) })).rejects.toThrow('timed out');
    await expect(readSource(`${SOURCE}/file`, { binary: true, timeoutMs: 10, fetchImpl: async () => ({ ok: true, arrayBuffer: () => new Promise(() => {}) }) })).rejects.toThrow('timed out');
    const controller = new AbortController();
    const pending = readSource(`${SOURCE}/file`, { signal: controller.signal, fetchImpl: () => new Promise(() => {}) });
    controller.abort(); await expect(pending).rejects.toThrow('Stopped');
  });
  it('does not package sign-in pages or incomplete inventories as a successful export', async () => {
    await expect(readSource(`${SOURCE}/file.docx`, { binary: true, fetchImpl: async () => new Response('<!doctype html><html>Sign in</html>') })).rejects.toThrow('sign-in');
    await expect(makeZip(makePlan(), new Map())).rejects.toThrow('Every');
    expect(() => makePlan([folder, { ...file, AttachmentFiles: [] }])).toThrow('no attached file');
    expect(() => makePlan([file])).toThrow('Broken folder');
    expect(() => makePlan([folder, { ...folder, Id: 9 }])).toThrow('duplicate');
    expect(() => makePlan([{ ...folder, ReferenceParentId: 'folder' }])).toThrow('Broken folder');
  });
  it('preserves bytes, renamed display names, archived copies, extra attachments and empty folders in ZIP', async () => {
    const archived = { ...file, Id: 3, RecordId: 'archive', Archived: true };
    const extra = { ...file, Id: 4, RecordId: 'extra', Title: 'More.docx', AttachmentFiles: [...file.AttachmentFiles, { FileName: 'Notes.txt' }] };
    const empty = { ...folder, Id: 5, RecordId: 'empty', Title: 'Empty' };
    const plan = makePlan([folder, file, archived, extra, empty]);
    const downloads = new Map();
    for (const item of plan.files) downloads.set(item.key, await acceptFile(item, new Uint8Array([1, 2, 3]).buffer));
    const { blob, manifest } = await makeZip(plan, downloads);
    const zip = await JSZip.loadAsync(await blob.arrayBuffer(), { checkCRC32: true });
    expect([...await zip.file('Active/Templates/Sample.docx').async('uint8array')]).toEqual([1, 2, 3]);
    expect(zip.files['Active/Empty/'].dir).toBe(true);
    expect(zip.file('Archived/Templates/Sample.docx')).toBeTruthy();
    expect(manifest.files).toHaveLength(4);
    expect(manifest.files.every(item => item.sha256.length === 64)).toBe(true);
    const guide = await zip.file('Reference-Document-Upload-Guide.html').async('text');
    expect(guide).toContain('Active/Templates/Sample.docx');
    expect(guide).toContain('metsoft/SitePages/Modernization-Tracker.aspx');
    expect(guide).toContain('Archived — recovery only');
  });
  it('checks manually supplied names and sizes and disambiguates ZIP name collisions', async () => {
    const plan = makePlan([folder, file, { ...file, Id: 3, RecordId: 'other' }]);
    expect(new Set(plan.files.map(item => item.zipPath)).size).toBe(2);
    await expect(acceptFile(plan.files[0], new Uint8Array(3).buffer, { name: 'wrong.docx' })).rejects.toThrow('original file');
    await expect(acceptFile(plan.files[0], new Uint8Array(2).buffer)).rejects.toThrow('size does not match');
  });
});
