import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import { chromium } from 'playwright';
import JSZip from 'jszip';
const artifact = await fs.readFile('build-reference-export/reference-document-export.html', 'utf8');
const source = 'https://flankspeed.sharepoint-mil.us/sites/ISEAMETENG';
const list = { Id: '11111111-1111-1111-1111-111111111111', Title: 'Modernization-Tracker - Reference Documents', BaseTemplate: 100 };
const rows = [
  { Id: 1, RecordId: 'folder', Title: 'Templates', EntryKind: 'folder', AttachmentFiles: [], Modified: '2026-01-01' },
  ...[2, 3].map(id => ({ Id: id, RecordId: `file-${id}`, Title: `File-${id}.docx`, EntryKind: 'file', ReferenceParentId: 'folder', FileName: `File-${id}.docx`, FileSize: 3, AttachmentFiles: [{ FileName: `File-${id}.docx` }], Modified: '2026-01-01' })),
];
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1280, height: 980 }, acceptDownloads: true });
const requests = [], errors = [], counts = new Map();
let forceManual = false;
page.on('pageerror', error => errors.push(error.message));
await page.route('**/*', async route => {
  const request = route.request(), url = request.url();
  requests.push({ url, method: request.method() });
  if (url === `${source}/SitePages/Export-test.aspx`) return route.fulfill({ contentType: 'text/html', body: '<!doctype html><html><body><iframe id="app" style="width:100%;height:96vh;border:0"></iframe></body></html>' });
  assert(url.startsWith(`${source}/`), `Unexpected host request: ${url}`);
  assert.equal(request.method(), 'GET', 'Exporter must never write SharePoint');
  let body;
  if (url.includes('/lists?')) body = { value: [list] };
  else if (url.includes('/RootFolder?')) body = { ServerRelativeUrl: '/sites/ISEAMETENG/Lists/ModernizationReferenceDocuments' };
  else if (url.includes('/items?')) body = { value: rows };
  else if (url.includes('/Attachments/')) {
    counts.set(url, (counts.get(url) || 0) + 1);
    if (url.endsWith('File-3.docx') && (counts.get(url) === 1 || forceManual)) return route.fulfill({ status: 503, body: 'Temporary read failure' });
    return route.fulfill({ contentType: 'application/octet-stream', body: Buffer.from([1, 2, 3]) });
  } else throw new Error(`Unexpected source request: ${url}`);
  return route.fulfill({ contentType: 'application/json', body: JSON.stringify(body) });
});
try {
  await page.goto(`${source}/SitePages/Export-test.aspx`);
  await page.locator('#app').evaluate((frame, html) => { frame.srcdoc = html; }, artifact);
  const frame = page.frameLocator('#app');
  await frame.getByRole('heading', { name: 'Reference document export', exact: true }).waitFor();
  await frame.getByRole('button', { name: 'Scan reference documents' }).click();
  await frame.getByRole('heading', { name: '2 attachments · 1 folders' }).waitFor();
  await frame.getByRole('button', { name: 'Download ZIP', exact: true }).click();
  await frame.locator('#error').filter({ hasText: 'HTTP 503' }).waitFor();
  assert.match(await frame.locator('#status').innerText(), /1 of 2 attachments retained/);
  const downloadPromise = page.waitForEvent('download');
  await frame.getByRole('button', { name: 'Retry / Download ZIP', exact: true }).click();
  const download = await downloadPromise;
  const zip = await JSZip.loadAsync(await fs.readFile(await download.path()), { checkCRC32: true });
  assert.deepEqual([...await zip.file('Active/Templates/File-2.docx').async('uint8array')], [1, 2, 3]);
  assert.equal(counts.get(`${source}/Lists/ModernizationReferenceDocuments/Attachments/2/File-2.docx`), 1, 'Retry redownloaded a completed file');
  assert.equal(JSON.parse(await zip.file('reference-manifest.json').async('text')).files.length, 2);
  assert.match(await zip.file('Reference-Document-Upload-Guide.html').async('text'), /Reference Documents \/ Templates/);
  // Exercise the manual fallback using a real browser File and another full ZIP.
  forceManual = true;
  await frame.getByRole('button', { name: 'Scan reference documents' }).click();
  await frame.getByRole('button', { name: 'Download ZIP', exact: true }).click();
  await frame.locator('#error').filter({ hasText: 'HTTP 503' }).waitFor();
  await frame.getByLabel('Choose downloaded file for File-3.docx', { exact: true }).setInputFiles({ name: 'File-3.docx', mimeType: 'application/octet-stream', buffer: Buffer.from([1, 2, 3]) });
  await frame.locator('.file-status').filter({ hasText: 'manually supplied' }).waitFor();
  const manualPromise = page.waitForEvent('download');
  await frame.getByRole('button', { name: 'Retry / Download ZIP', exact: true }).click();
  const manualZip = await JSZip.loadAsync(await fs.readFile(await (await manualPromise).path()));
  assert.match(JSON.parse(await manualZip.file('reference-manifest.json').async('text')).files[1].retrievedBy, /Manually/);
  for (const colorScheme of ['light', 'dark']) {
    await page.emulateMedia({ colorScheme });
    assert(await frame.locator('body').evaluate(body => body.scrollWidth <= window.innerWidth + 1), 'Horizontal overflow');
  }
  await page.setViewportSize({ width: 390, height: 844 });
  assert(await frame.locator('body').evaluate(body => body.scrollWidth <= window.innerWidth + 1), 'Mobile horizontal overflow');
  assert.deepEqual(errors, []);
  console.log(`Exporter browser smoke passed: srcdoc boot, GET-only transport, retry reuse, manual fallback, ZIP CRC/bytes/manifest, guide, light/dark/mobile (${requests.length} requests).`);
} finally { await browser.close(); }
