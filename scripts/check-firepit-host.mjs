import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { createServer } from 'vite';

// Optional integration check against an independently supplied Firepit checkout.
// No host implementation is included in this repository. Install that checkout's
// pinned dependencies first (or expose them through NODE_PATH).
const source = process.env.FIREPIT_SOURCE;
if (!source) throw new Error('Set FIREPIT_SOURCE to the Firepit 9.2 source directory before running this optional check.');
const hostRoot = path.resolve(source), require = createRequire(import.meta.url);
const tool = name => path.join(hostRoot, 'tools', name);
const developmentBundle = path.join(hostRoot, 'bundle.js');
const deployedBundle = path.join(hostRoot, 'src/package/ClientSideAssets/bundle.js');
let hostHarness;
const readFileSync = fs.readFileSync;
try {
  // Some source archives omit the unused development bundle. The upstream
  // harness reads it eagerly even when testing the deployed bundle. Redirect
  // only that absent read, then restore fs immediately; do not edit host code.
  fs.readFileSync = function (file, ...args) {
    if (typeof file === 'string' && path.resolve(file) === developmentBundle && !fs.existsSync(developmentBundle)) {
      return readFileSync.call(fs, deployedBundle, ...args);
    }
    return readFileSync.call(fs, file, ...args);
  };
  hostHarness = require(tool('permission-stress-harness.cjs'));
} finally { fs.readFileSync = readFileSync; }
const { listTargetsRuntime } = require(tool('list-targets.js'));
const { siteScopeRuntime } = require(tool('site-scope.js'));
const { versionProtectionRuntime } = require(tool('version-protection.js'));
const { downloadPolicyRuntime } = require(tool('download-policy.js'));
const { harness, web, flush } = hostHarness;
const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const vite = await createServer({ root: repo, logLevel: 'error', appType: 'custom',
  optimizeDeps: { noDiscovery: true, include: [] }, server: { middlewareMode: true } });
const originalWindow = Object.getOwnPropertyDescriptor(globalThis, 'window');
const rows = [
  ['tasks', 'Tasks', '11111111-1111-4111-8111-111111111111'],
  ['references', 'Reference Documents', '22222222-2222-4222-8222-222222222222'],
  ['projects', 'Projects', '33333333-3333-4333-8333-333333333333'],
].map(([key, title, Id]) => ({ key, Id, Title: 'Modernization-Tracker - ' + title,
  BaseType: 0, BaseTemplate: 100, Hidden: false, ListItemEntityTypeFullName: 'SP.Data.TrackerListItem',
  RootFolder: { ServerRelativeUrl: '/sites/team/Lists/Modernization' + title.replaceAll(' ', '') },
  EnableVersioning: true, MajorVersionLimit: 50000, EnableFolderCreation: true,
  HasUniqueRoleAssignments: false, ReadSecurity: 1,
  EffectiveBasePermissions: { Low: '2147483647', High: '2147483647' } }));
const tasks = rows.find(row => row.key === 'tasks');
const referenceList = rows.find(row => row.key === 'references');
const folderPath = tasks.RootFolder.ServerRelativeUrl + '/tracker-project-7';
const pdfPath = tasks.RootFolder.ServerRelativeUrl + '/Attachments/7/guide.pdf';
const macroName = 'PFA - PFR Lookup Table.xlsm';
const macroPath = referenceList.RootFolder.ServerRelativeUrl + '/Attachments/9/' + macroName;
const expectedForm = { formValues: [{ FieldName: 'TaskStatus', FieldValue: 'Complete' }], bNewDocumentUpdate: false };
const json = value => new Response(JSON.stringify(value), { headers: { 'Content-Type': 'application/json;odata=nometadata' } });
const pathOf = value => decodeURIComponent(new URL(value).pathname);

try {
  const { SharePointStore } = await vite.ssrLoadModule('/src/lib/spStore.js');
  const { fileDownloadUrl } = await vite.ssrLoadModule('/src/lib/firepit.js');
  function fixture() {
    const h = harness({ versionProtection: true }), calls = [];
    h.context.FirepitSiteScope = siteScopeRuntime();
    h.context.FirepitListTargets = listTargetsRuntime(h.context.FirepitRequestCoordinator);
    // Rebind the real versioning implementation to the same restored scope.
    vm.runInContext('FirepitVersionProtection=(' + versionProtectionRuntime.toString() + ')(FirepitSiteScope,FirepitConfirmations)', h.context);
    h.host.context.pageContext.user = { loginName: 'fixture@example.test' };
    let folderKey = '';
    h.useSharePointClient(async (address, _configuration, init) => {
      const url = new URL(address), pathname = pathOf(address);
      const headers = new Headers(init.headers), method = init.method || 'GET';
      const body = init.body == null ? undefined : typeof init.body === 'string' ? init.body : new TextDecoder().decode(init.body);
      calls.push({ url: url.href, method, headers, body });
      if (pathname.toLowerCase().endsWith('/contextinfo')) return json({ d: { GetContextWebInformation: { FormDigestValue: 'fixture', FormDigestTimeoutSeconds: 1800 } } });
      if (pathname.toLowerCase().endsWith('/_api/web/lists')) {
        const filter = url.searchParams.get('$filter');
        return json({ value: filter ? rows.filter(row => filter.includes("'" + row.Title + "'")) : rows });
      }
      const id = /lists\(guid'([^']+)'\)/i.exec(pathname)?.[1];
      const list = rows.find(row => row.Id === id);
      if (list) {
        if (/\/fields$/i.test(pathname)) return json({ value: [{ InternalName: 'TaskStatus', TypeAsString: 'Choice', ReadOnlyField: false, Hidden: false }] });
        if (/\/ValidateUpdateListItem(?:\(\))?$/i.test(pathname)) {
          assert.equal(method, 'POST'); assert.deepEqual(JSON.parse(body), expectedForm);
          return json({ value: [{ FieldName: 'TaskStatus', FieldValue: 'Complete', HasException: false, ErrorMessage: null }] });
        }
        if (/\/AttachmentFiles$/i.test(pathname)) return json({ value: [{ FileName: macroName, ServerRelativeUrl: macroPath }] });
        if (/\/items$/i.test(pathname)) return json({ value: [{ Id: 41 }] });
        if (/\/items\(41\)$/i.test(pathname)) {
          if (headers.get('X-HTTP-Method') === 'MERGE') {
            assert.equal(headers.get('IF-MATCH'), '"1"');
            assert.deepEqual(JSON.parse(body), { ProjectKey: 'project-7' });
            folderKey = 'project-7'; return json({});
          }
          return json({ Id: 41, FileSystemObjectType: 1, FileRef: folderPath, ProjectKey: folderKey, 'odata.etag': '"1"' });
        }
        if (/\/items\(7\)$/i.test(pathname) && list.key === 'projects') return json({ Id: 7, ProjectKey: 'project-7' });
        return json(list);
      }
      if (/GetFolderByServerRelativePath/i.test(pathname)) return json({ ServerRelativeUrl: folderPath, ItemCount: 0 });
      if (/GetFileByServerRelativePath/i.test(pathname) && /\/\$value$/i.test(pathname)) return new Response('%PDF-1.7\nfixture', { headers: { 'Content-Type': 'application/pdf' } });
      throw new Error('Unexpected SharePoint fixture request: ' + method + ' ' + url.href);
    });
    globalThis.window = { firepit: { sharepoint: h.api } };
    const store = new SharePointStore({ webUrl: web, scopedAccess: true, fetchImpl() { throw new Error('Direct network fallback must not run'); } });
    return { h, calls, store };
  }
  let passed = 0;
  async function check(name, action) {
    const test = fixture();
    try { await action(test); await flush(); test.h.clean(); console.log('PASS ' + name); passed++; }
    finally { delete globalThis.window; }
  }
  await check('actual Tracker update traverses schema and version checks, then sends one unchanged write', async ({ store, calls }) => {
    await store.update('tasks', 7, { TaskStatus: 'Complete' });
    assert.equal(calls.filter(call => call.method === 'POST').length, 1);
    assert(calls.some(call => new URL(call.url).searchParams.get('$select')?.includes('EnableVersioning')));
  });
  for (const [kind, target, suffix] of [['Folder', folderPath, ''], ['File', pdfPath, '/$value']]) {
    await check('old literal ' + kind.toLowerCase() + ' path reproduces the exact reported host rejection', async ({ store, calls }) => {
      await assert.rejects(store.fetchImpl(web + '/_api/web/Get' + kind + "ByServerRelativePath(decodedurl='" + target + "')" + suffix), /List validation: this list target cannot be identified safely\./);
      assert(calls.every(call => /\/_api\/web\/lists$/i.test(pathOf(call.url))), 'No file/folder business request should be dispatched');
    });
  }
  await check('actual aliased file helper returns unchanged bytes through the full host', async ({ store }) => {
    const response = await store.fetchImpl(fileDownloadUrl(web, new URL(pdfPath, web).href));
    assert.equal(await response.text(), '%PDF-1.7\nfixture');
  });
  await check('actual folder recovery verifies identity, emptiness, parent key and ETag before its write', async ({ store, calls }) => {
    const result = await store.permissions.folder('tasks', { spId: 7, projectKey: 'project-7' }, true);
    assert.equal(result.path, folderPath);
    assert(calls.some(call => /decodedurl=@folder/i.test(call.url)));
    assert.equal(calls.filter(call => call.method === 'POST').length, 1);
  });
  await check('actual XLSM reference download returns native navigation without reading/reconstructing file bytes', async ({ store, calls, h }) => {
    const result = await store.downloadReferenceEntry({ spId: 9, fileName: macroName });
    assert.deepEqual(result, { type: 'sharepoint-download', url: new URL(macroPath, web).href, fileName: macroName });
    assert(calls.some(call => /\/AttachmentFiles$/i.test(pathOf(call.url))));
    assert(!calls.some(call => /\/\$value$/i.test(pathOf(call.url))));
    assert.equal(h.context.FirepitSiteScope.hasDataAccess(h.host), true, 'The metadata response requires parent navigation confirmation before opening the stored file');
    assert.equal(downloadPolicyRuntime(true).reason(macroName), 'Firepit blocks .xlsm downloads.', 'A renamed Blob export must not replace native navigation');
  });
  console.log(`${passed} checks passed with real generated child/parent handlers, list validation, site scope, version protection and transport. DOM, clock and tenant responses are fixtures; no live SharePoint requests were made.`);
} finally {
  if (originalWindow) Object.defineProperty(globalThis, 'window', originalWindow);
  else delete globalThis.window;
  await vite.close();
}
