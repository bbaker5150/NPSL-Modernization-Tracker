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
const { listPermissionWarningRuntime } = require(tool('list-permission-warning.js'));
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
  ['updates', 'Updates', '44444444-4444-4444-8444-444444444444'],
  ['risks', 'Risks', '55555555-5555-4555-8555-555555555555'],
  ['users', 'Users', '66666666-6666-4666-8666-666666666666'],
  ['acronyms', 'Acronyms', '77777777-7777-4777-8777-777777777777'],
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
  const { SharePointStore, CONTAINERS } = await vite.ssrLoadModule('/src/lib/spStore.js');
  const { authorizedStore } = await vite.ssrLoadModule('/src/lib/access.js');
  const { fileDownloadUrl } = await vite.ssrLoadModule('/src/lib/firepit.js');
  function fixture({ authorization = false } = {}) {
    const h = harness({ versionProtection: true }), calls = [];
    // Confirmation UI also uses intervals; extend the harness's simulated clock
    // without changing when a review approves, cancels, or checks its lifetime.
    h.context.setInterval = (callback, delay) => {
      const timer = { active: true };
      const tick = () => { if (timer.active) { callback(); if (timer.active) timer.id = h.clock.setTimeout(tick, delay); } };
      timer.id = h.clock.setTimeout(tick, delay);
      return timer;
    };
    h.context.clearInterval = timer => { if (timer) { timer.active = false; h.clock.clearTimeout(timer.id); } };
    h.context.FirepitSiteScope = siteScopeRuntime();
    h.context.FirepitListTargets = listTargetsRuntime(h.context.FirepitRequestCoordinator);
    // Rebind the real versioning implementation to the same restored scope.
    vm.runInContext('FirepitVersionProtection=(' + versionProtectionRuntime.toString() + ')(FirepitSiteScope,FirepitConfirmations)', h.context);
    // The upstream stress harness replaces list validation after constructing
    // this runtime. Bind it to the real validator and real confirmation UI too.
    vm.runInContext('FirepitListPermissionWarning=(' + listPermissionWarningRuntime.toString() + ')(FirepitListTargets,FirepitConfirmations,FirepitRequestCoordinator)', h.context);
    h.host.context.pageContext.user = { loginName: 'fixture@example.test' };
    const metadata = row => ({ ...row, HasUniqueRoleAssignments: authorization && ['users', 'acronyms'].includes(row.key) });
    const state = { role: 'Manager', progressMode: 'phases', ownerKey: 'fixture@example.test', glossaryTerm: 'First glossary value' };
    const items = key => ({
      projects: [{ Id: 7, RecordId: 'project-id-7', ProjectKey: 'project-7', Title: 'Fixture project', OwnerKey: state.ownerKey, OwnerEmail: state.ownerKey, ProgressMode: state.progressMode }],
      tasks: [{ Id: 8, RecordId: 'task-id-8', ProjectKey: 'project-7', TaskTitle: 'Fixture task', TaskStatus: 'Not Started' }],
      users: [{ Id: 1, RecordId: 'user-id-1', LoginKey: 'fixture@example.test', Email: 'fixture@example.test', Title: 'Fixture user', AppRole: 'Manager' }],
      acronyms: [{ Id: 2, RecordId: 'acronym-id-2', Acronym: 'QA', FullTerm: state.glossaryTerm }],
      updates: [], risks: [],
    }[key] || []);
    let folderKey = '';
    h.useSharePointClient(async (address, _configuration, init) => {
      const url = new URL(address), pathname = pathOf(address);
      const headers = new Headers(init.headers), method = init.method || 'GET';
      const body = init.body == null ? undefined : typeof init.body === 'string' ? init.body : new TextDecoder().decode(init.body);
      calls.push({ url: url.href, method, headers, body });
      if (pathname.toLowerCase().endsWith('/contextinfo')) return json({ d: { GetContextWebInformation: { FormDigestValue: 'fixture', FormDigestTimeoutSeconds: 1800 } } });
      if (/\/_api\/web\/currentuser$/i.test(pathname)) return json({ Id: 1, LoginName: 'fixture@example.test', Email: 'fixture@example.test', Title: 'Fixture user', IsSiteAdmin: false });
      if (/\/_api\/web\/currentuser\/groups$/i.test(pathname)) return json({ value: state.role === 'Viewer' ? [] : [{ Title: state.role === 'Manager' ? 'Tracker Managers' : 'Tracker Project Engineers' }] });
      if (/\/_api\/web\/effectivebasepermissions$/i.test(pathname)) return json({ Low: '1', High: '0' });
      if (pathname.toLowerCase().endsWith('/_api/web/lists')) {
        const filter = url.searchParams.get('$filter');
        return json({ value: (filter ? rows.filter(row => filter.includes("'" + row.Title + "'")) : rows).map(metadata) });
      }
      const id = /lists\(guid'([^']+)'\)/i.exec(pathname)?.[1];
      const list = rows.find(row => row.Id === id);
      if (list) {
        if (/\/roleassignments$/i.test(pathname)) {
          const principal = authorization && list.key === 'users' ? 11 : authorization && list.key === 'acronyms' ? 12 : 10;
          return json({ value: [7, principal].map(PrincipalId => ({ PrincipalId, RoleDefinitionBindings: [{ BasePermissions: { Low: '1', High: '0' } }] })) });
        }
        if (/\/fields$/i.test(pathname)) return json({ value: CONTAINERS.find(container => container.key === list.key).fields.map(field => ({ InternalName: field.name, TypeAsString: field.type, ReadOnlyField: false, Hidden: false })) });
        if (/\/ValidateUpdateListItem(?:\(\))?$/i.test(pathname)) {
          assert.equal(method, 'POST');
          const form = JSON.parse(body);
          if (authorization) {
            assert.equal(list.key, 'projects');
            assert.equal(form.bNewDocumentUpdate, false);
            assert.deepEqual(form.formValues.map(value => value.FieldName), ['ProgressMode']);
            state.progressMode = form.formValues[0].FieldValue;
          } else assert.deepEqual(form, expectedForm);
          return json({ value: form.formValues.map(value => ({ ...value, HasException: false, ErrorMessage: null })) });
        }
        if (/\/AttachmentFiles$/i.test(pathname)) return json({ value: [{ FileName: macroName, ServerRelativeUrl: macroPath }] });
        if (/\/items$/i.test(pathname)) return json({ value: authorization ? items(list.key) : [{ Id: 41 }] });
        if (/\/items\(41\)$/i.test(pathname)) {
          if (headers.get('X-HTTP-Method') === 'MERGE') {
            assert.equal(headers.get('IF-MATCH'), '"1"');
            assert.deepEqual(JSON.parse(body), { ProjectKey: 'project-7' });
            folderKey = 'project-7'; return json({});
          }
          return json({ Id: 41, FileSystemObjectType: 1, FileRef: folderPath, ProjectKey: folderKey, 'odata.etag': '"1"' });
        }
        if (/\/items\(7\)$/i.test(pathname) && list.key === 'projects') return json({ Id: 7, ProjectKey: 'project-7' });
        return json(metadata(list));
      }
      if (/GetFolderByServerRelativePath/i.test(pathname)) return json({ ServerRelativeUrl: folderPath, ItemCount: 0 });
      if (/GetFileByServerRelativePath/i.test(pathname) && /\/\$value$/i.test(pathname)) return new Response('%PDF-1.7\nfixture', { headers: { 'Content-Type': 'application/pdf' } });
      throw new Error('Unexpected SharePoint fixture request: ' + method + ' ' + url.href);
    });
    globalThis.window = { firepit: { sharepoint: h.api } };
    const store = new SharePointStore({ webUrl: web, scopedAccess: true, fetchImpl() { throw new Error('Direct network fallback must not run'); } });
    return { h, calls, store, state, authorized: authorizedStore(store) };
  }
  let passed = 0;
  async function check(name, action, options) {
    const test = fixture(options);
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
  const listReads = calls => calls.filter(call => call.method === 'GET' && /\/items$/i.test(pathOf(call.url)))
    .map(call => rows.find(row => call.url.includes(row.Id))?.key);
  const writes = calls => calls.filter(call => call.method === 'POST');
  const walk = node => [node, ...node.children.flatMap(walk)];
  await check('fresh authorization omits unused glossary, skips unchanged writes and preserves real cross-list consent', async ({ authorized, store, calls, h, state }) => {
    const assertAuthorizationReads = start => {
      assert.deepEqual([...new Set(listReads(calls.slice(start)))].sort(), ['projects', 'risks', 'tasks', 'updates', 'users']);
      assert(calls.slice(start).some(call => /\/currentuser\/groups$/i.test(pathOf(call.url))), 'Each authorization must recheck live role membership');
      assert(calls.slice(start).some(call => /\/effectivebasepermissions$/i.test(pathOf(call.url))), 'Each authorization must recheck current site rights');
    };
    async function decideFlow(pending, source, approve) {
      // Observe the real host-created dialog, then click its actual button.
      // Do not replace the approval API or seed its accepted-flow cache.
      let failure;
      pending.catch(error => { failure = error; });
      await flush();
      if (failure) throw failure;
      const dialogs = h.dialogs();
      assert.equal(dialogs.length, 1, 'A verified permission difference must require a visible review');
      assert.match(dialogs[0].textContent, /This write could share data with more people/);
      assert(dialogs[0].textContent.includes('Modernization-Tracker - ' + source));
      assert(dialogs[0].textContent.includes('Modernization-Tracker - Projects'));
      assert(!dialogs[0].textContent.includes('Modernization-Tracker - ' + (source === 'Users' ? 'Acronyms' : 'Users')),
        'Only the newly unapproved source should be named');
      const button = walk(dialogs[0]).find(node => node.tag === 'button' && node.textContent === (approve ? 'Approve flow for this session' : 'Cancel write'));
      assert(button, 'Use the real flow decision button');
      button.onclick();
      return pending;
    }
    let start = calls.length;
    assert.equal((await authorized.saveProgressMode('project-id-7', 'phases')).progressMode, 'phases');
    assertAuthorizationReads(start);
    assert.equal(writes(calls).length, 0, 'Unchanged progress must not issue a write');
    assert.equal(h.shown(), 0);

    start = calls.length;
    await assert.rejects(decideFlow(authorized.saveProgressMode('project-id-7', 'tasks'), 'Users', false), /list permission warning was declined/i);
    assertAuthorizationReads(start);
    assert.equal(writes(calls).length, 0, 'Declining the real host review must prevent dispatch');

    start = calls.length;
    await decideFlow(authorized.saveProgressMode('project-id-7', 'tasks'), 'Users', true);
    assertAuthorizationReads(start);
    assert.equal(writes(calls).length, 1);
    assert.equal(h.shown(), 2, 'Declined flow must be reviewed again before its first successful write');
    start = calls.length;
    await authorized.saveProgressMode('project-id-7', 'tasks');
    assertAuthorizationReads(start);
    assert.equal(writes(calls).length, 1, 'Saving the new current value must not write again');
    start = calls.length;
    await authorized.saveProgressMode('project-id-7', 'phases');
    assertAuthorizationReads(start);
    assert.equal(writes(calls).length, 2);
    assert.equal(h.shown(), 2, 'Stable approved permission snapshots reuse the host session approval');

    const task = (await authorized.load({ includeAcronyms: false })).tasks[0];
    start = calls.length;
    assert.deepEqual(await authorized.saveTask(task), task);
    assertAuthorizationReads(start);
    assert.equal(writes(calls).length, 2, 'An unchanged task must not write after its fresh authorization');
    assert(!listReads(calls).includes('acronyms'), 'Ordinary authorization and saves must never read glossary contents');

    start = calls.length;
    assert.equal((await store.listAcronyms())[0].term, 'First glossary value');
    state.glossaryTerm = 'Updated glossary value';
    assert.equal((await store.listAcronyms())[0].term, 'Updated glossary value');
    assert.deepEqual(listReads(calls.slice(start)), ['acronyms', 'acronyms'], 'Explicit glossary reads must remain fresh');
    await decideFlow(authorized.saveProgressMode('project-id-7', 'tasks'), 'Acronyms', true);
    assert.equal(writes(calls).length, 3);
    assert.equal(h.shown(), 3, 'Reading a newly differing source must retain the legitimate new-flow review');

    for (const role of ['Viewer', 'Project Engineer']) {
      state.role = role;
      state.ownerKey = 'someone-else@example.test';
      start = calls.length;
      await assert.rejects(authorized.saveProgressMode('project-id-7', 'tasks'), /Only managers or assigned project engineers/);
      assertAuthorizationReads(start);
      assert.equal(writes(calls).length, 3, 'Identical values must not bypass live role or assignment checks');
    }
    assert.equal(h.nativePrompts.length, 0, 'Consent must use the real parent dialog, not a native confirm fallback');
  }, { authorization: true });
  await check('default raw and authorized exports still read the complete fresh glossary', async ({ authorized, store, calls, state }) => {
    assert.equal((await store.load()).acronyms[0].term, 'First glossary value');
    state.glossaryTerm = 'Latest export value';
    assert.equal((await authorized.load()).acronyms[0].term, 'Latest export value');
    assert.equal(listReads(calls).filter(key => key === 'acronyms').length, 2);
    assert.equal(writes(calls).length, 0);
  }, { authorization: true });
  console.log(`${passed} checks passed with real generated child/parent handlers, list validation, site scope, version protection, list-flow consent and transport. DOM, clock and tenant responses are fixtures; no live SharePoint requests were made.`);
} finally {
  if (originalWindow) Object.defineProperty(globalThis, 'window', originalWindow);
  else delete globalThis.window;
  await vite.close();
}
