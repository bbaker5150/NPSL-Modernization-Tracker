import { afterEach, expect, it, vi } from 'vitest';
import { SharePointStore } from './spStore';
import { downloadDocument } from './downloadDocument';
import { sharePointDownload } from './firepit';

const webUrl = 'https://tenant.sharepoint-mil.us/sites/metsoft';
const name = 'PFA - PFR Lookup Table.xlsm';
const path = '/sites/metsoft/Lists/ModernizationReferenceDocuments/Attachments/7/PFA%20-%20PFR%20Lookup%20Table.xlsm';
afterEach(() => { delete window.firepit; vi.restoreAllMocks(); vi.unstubAllGlobals(); });

it.each(['reference', 'task', 'update'])('uses SharePoint native navigation for a %s download without reading macro-enabled workbook bytes', async kind => {
  const bridge = { siteUrl: webUrl, fetch: vi.fn() };
  window.firepit = { sharepoint: bridge };
  const store = new SharePointStore({ webUrl });
  store.get = vi.fn(async () => ({ value: [{ FileName: name, ServerRelativeUrl: path }] }));
  store.listTaskAttachments = store.listUpdateAttachments = vi.fn(async () => [{ name, url: new URL(path, webUrl).href }]);
  const result = kind === 'reference' ? await store.downloadReferenceEntry({ spId: 7, fileName: name })
    : kind === 'task' ? await store.downloadTaskAttachment({ spId: 7 }, name)
      : await store.downloadUpdateAttachment({ spId: 7 }, name);
  expect(result).toEqual({ type: 'sharepoint-download', url: new URL(path, webUrl).href, fileName: name });
  expect(bridge.fetch).not.toHaveBeenCalled();
  const clicks = [];
  vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function () { clicks.push({ href: this.href, name: this.download, target: this.target }); });
  const createObjectURL = vi.fn(); vi.stubGlobal('URL', class extends URL { static createObjectURL = createObjectURL; });
  downloadDocument(result, 'A displayed name with no extension');
  expect(clicks).toEqual([{ href: new URL(path, webUrl).href, name, target: '_blank' }]);
  expect(createObjectURL).not.toHaveBeenCalled();
});

it('does not create a native link for another tenant or outside the configured web', () => {
  window.firepit = { sharepoint: { siteUrl: webUrl, fetch: vi.fn() } };
  expect(() => sharePointDownload(webUrl, 'https://outside.test/file.xlsm', name)).toThrow('outside the configured');
  expect(() => sharePointDownload(webUrl, '/sites/other/file.xlsm', name)).toThrow('outside the configured');
});
