import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import { chromium } from 'playwright';

// Snapshot of the deployed schema before PR #80. Keep this fixture independent
// of current CONTAINERS so adding another mandatory column cannot hide a regression.
const priorSchema = JSON.parse(await fs.readFile('scripts/fixtures/pre-update-editing-schema.json', 'utf8'));
const artifact = await fs.readFile('build-singlefile/modernization-project-tracker.html', 'utf8');
await fs.mkdir('test-artifacts', { recursive: true });
const webUrl = 'https://tenant.sharepoint-mil.us/sites/metsoft';
const firepit = process.argv.includes('--firepit');
// Contract fixture, not a copy of the host implementation. The opaque child can
// only ask the parent bridge for REST responses; CSP blocks direct networking/eval.
const bridgeScript = `
  let requestId = 0; const pending = new Map();
  addEventListener('message', event => {
    if (event.source !== parent || event.data?.type !== 'mock-bridge-result') return;
    const item = pending.get(event.data.id); if (!item) return;
    pending.delete(event.data.id);
    item.resolve(new Response(event.data.body, {status: event.data.status, headers: {'Content-Type':'application/json'}}));
  });
  window.firepit = {sharepoint: {siteUrl: ${JSON.stringify(webUrl)}, fetch: (input, init = {}) => new Promise((resolve, reject) => {
    const id = ++requestId; pending.set(id, {resolve, reject});
    parent.postMessage({type:'mock-bridge-request', id, url: typeof input === 'string' ? input : input.url,
      method:init.method || input.method || 'GET', headers:[...new Headers(init.headers).entries()], body:init.body}, '*');
  })}};
`;
const parentRelay = `<script>addEventListener('message', async event => {
  if (event.source !== document.querySelector('#app').contentWindow || event.data?.type !== 'mock-bridge-request') return;
  const {id, url, method, headers, body} = event.data;
  const result = await fetch(url, {method, headers, body});
  event.source.postMessage({type:'mock-bridge-result', id, status:result.status, body:await result.text()}, '*');
});</script>`;
const html = artifact.replace('<head>', `<head>${firepit ? `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data: blob:; font-src data:; connect-src 'none'"><script>${bridgeScript}</script>` : ''}<script>window.MOD_TRACKER_CONFIG=${JSON.stringify({ ...(firepit ? {} : { webUrl }), forceSharePoint: true, scopedAccess: true })}</script>`);
const suffixes = ['Projects', 'Tasks', 'Updates', 'Risks', 'Acronyms', 'Users', 'Reference Documents'];
const lists = Object.keys(priorSchema).map((key, index) => ({ key, Id: `${String(index + 1).padStart(8, '0')}-1111-1111-1111-111111111111`, Title: `Modernization-Tracker - ${suffixes[index]}`, BaseTemplate: 100, Hidden: false }));
const user = { Id: 10, Title: 'Viewer Test Account', LoginName: 'i:0#.f|membership|viewer@example.test', Email: 'viewer@example.test', IsSiteAdmin: false };
const rows = {
  projects: [{ Id: 1, RecordId: 'p', ProjectKey: 'p', Title: 'Viewer access project', MeasurementArea: 'Thermometry', OwnerName: 'Project Engineer', OwnerEmail: 'engineer@example.test', Organization: 'NAVAIR METCAL 260B', Health: 'On Track', Priority: 'Medium', ProjectStatus: 'Active', CurrentStageKey: 'requirement' }],
  tasks: [{ Id: 2, RecordId: 't', ProjectKey: 'p', TaskTitle: 'Review requirement', PhaseKey: 'requirement', TaskStatus: 'In Progress', DueDate: '2027-01-01', Organization: 'NAVAIR METCAL 260B' }],
  updates: [{ Id: 3, RecordId: 'u', ProjectKey: 'p', Summary: 'Existing project history', AuthorName: 'Original author', PostedAt: '2026-11-01T18:00:00Z', UpdateType: 'Status', 'odata.etag': '"1"' }],
  risks: [], acronyms: [], users: [{ Id: 10, RecordId: 'viewer', Title: user.Title, LoginKey: user.LoginName, Email: user.Email, AppRole: 'Viewer' }], references: [],
};
const browser = await chromium.launch();
try {
  for (const scenario of ['prior-schema', 'missing-core-field', 'missing-list', 'denied-items']) {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
    page.setDefaultTimeout(15000);
    const requests = [], unexpected = [], errors = [];
    page.on('pageerror', error => errors.push(error.message));
    const schemas = structuredClone(priorSchema);
    if (scenario === 'missing-core-field') schemas.updates = schemas.updates.filter(name => name !== 'Archived');
    await page.route('**/*', async route => {
      const request = route.request(), url = new URL(request.url()), path = url.pathname.slice(new URL(webUrl).pathname.length);
      const send = (body, status = 200) => route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(body) });
      if (url.href === `${webUrl}/preview`) return route.fulfill({ contentType: 'text/html', body: `<iframe id="app" ${firepit ? 'sandbox="allow-scripts allow-downloads"' : ''} style="position:fixed;inset:0;width:100%;height:100%;border:0"></iframe>${firepit ? parentRelay : ''}` });
      if (firepit) assert.equal(request.headers()['x-requestdigest'], undefined, 'The child must not supply a form digest');
      requests.push({ method: request.method(), path, select: url.searchParams.get('$select') });
      if (request.method() !== 'GET') { unexpected.push(`Unexpected write: ${request.method()} ${path}`); return send({ error: { message: 'Viewer cannot write setup or data' } }, 403); }
      if (path.toLowerCase().endsWith('/effectivebasepermissions')) return send({ Low: '1', High: '0' });
      if (path.endsWith('/currentuser/groups')) return send({ value: [{ Id: 20, Title: 'Tracker Users' }] });
      if (path.endsWith('/currentuser')) return send(user);
      if (path.endsWith('/associatedownergroup')) return send({ Id: 99 });
      if (path.endsWith('/lists')) {
        const filter = url.searchParams.get('$filter') || '';
        return send({ value: lists.filter(list => filter.includes(`'${list.Title}'`) && !(scenario === 'missing-list' && list.key === 'tasks')) });
      }
      const id = path.match(/lists\(guid'([^']+)'\)/)?.[1];
      const list = lists.find(item => item.Id === id);
      if (list) {
        if (path.endsWith('/fields')) return send({ value: schemas[list.key].map(InternalName => ({ InternalName })) });
        if (path.endsWith('/AttachmentFiles')) return send({ value: [] });
        const itemId = path.match(/\/items\((\d+)\)$/)?.[1];
        if (itemId) return send({ ...rows[list.key].find(row => row.Id === Number(itemId)), ArchivedDocuments: '[]' });
        if (path.endsWith('/items')) {
          if (scenario === 'denied-items' && list.key === 'projects') return send({ error: { message: 'Project read denied by SharePoint' } }, 403);
          const select = (url.searchParams.get('$select') || '').split(',');
          const unknown = select.filter(name => !['Id', 'Title', 'FileSystemObjectType'].includes(name) && !schemas[list.key].includes(name));
          if (unknown.length) { unexpected.push(`Missing columns selected: ${unknown}`); return send({ error: { message: `Unknown columns: ${unknown}` } }, 400); }
          return send({ value: rows[list.key] });
        }
        if (path.endsWith(`lists(guid'${id}')`)) return send({ Id: list.Id, EnableAttachments: true });
      }
      unexpected.push(`Unexpected read: ${url.href}`);
      return send({ error: { message: 'Unexpected mock request' } }, 400);
    });
    try {
      await page.goto(`${webUrl}/preview`);
      await page.locator('#app').evaluate((iframe, content) => { iframe.srcdoc = content; }, html);
      const app = page.frameLocator('#app');
      if (scenario === 'prior-schema') {
        await app.getByRole('heading', { name: 'Modernization at a glance', exact: true }).waitFor();
        const glossaryId = lists.find(list => list.key === 'acronyms').Id;
        assert.ok(!requests.some(request => request.path.includes(glossaryId) && request.path.endsWith('/items')), 'Ordinary startup must not read glossary rows before the glossary or export is requested');
        if (firepit) {
          assert.equal(await app.locator('body').evaluate(() => window.origin), 'null', 'Firepit smoke must use an opaque frame');
          assert.equal(await app.locator('body').evaluate(() => { try { localStorage.getItem('theme'); return true; } catch { return false; } }), false);
          const downloadReady = page.waitForEvent('download');
          await app.getByRole('button', { name: 'Export Excel', exact: true }).click();
          const download = await downloadReady;
          assert.match(download.suggestedFilename(), /\.xlsx$/);
          const workbook = await fs.readFile(await download.path());
          assert.equal(workbook.subarray(0, 2).toString(), 'PK', 'Workbook export must work without eval permission');
        }
        assert.equal(await app.getByRole('button', { name: 'New project', exact: true }).count(), 0);
        assert.equal(await app.getByRole('button', { name: 'Users and managers', exact: true }).count(), 0);
        assert.equal(await app.getByRole('heading', { name: 'Users and managers', exact: true }).count(), 0);
        await app.getByRole('button', { name: 'Portfolio', exact: true }).press('Control+m');
        assert.equal(await app.getByLabel('Testing password').count(), 0, 'SharePoint viewers cannot enable testing manager access');
        await app.locator('.project-table-row:not(.table-header)').click();
        await app.locator('.drawer-tabs').getByRole('button', { name: /^Updates/ }).click();
        await app.getByText('Existing project history', { exact: true }).waitFor();
        await app.getByText('Original author', { exact: true }).waitFor();
        assert.equal(await app.getByRole('button', { name: 'Edit update', exact: true }).count(), 0);
        assert.equal(await app.getByRole('button', { name: 'Delete update', exact: true }).count(), 0);
        assert.equal(await app.getByRole('textbox', { name: 'Your update', exact: true }).count(), 0);
        await page.screenshot({ path: 'test-artifacts/viewer-sharepoint-startup.png' });
        assert.ok(requests.some(request => request.path.endsWith('/items') && request.select.includes('Summary')));
        assert.ok(requests.filter(request => request.path.endsWith('/items')).every(request => !/EditedAt|EditorName|EditorKey/.test(request.select)));
      } else {
        await app.getByRole('button', { name: 'Try again', exact: true }).waitFor();
        const error = await app.locator('.gate-card').innerText();
        if (scenario === 'missing-core-field') { assert.match(error, /Modernization-Tracker - Updates: missing fields/); assert.match(error, /Archived/); }
        if (scenario === 'missing-list') assert.match(error, /Modernization-Tracker - Tasks: list is missing or not visible/);
        if (scenario === 'denied-items') assert.match(error, /Project read denied by SharePoint/);
        assert.equal(await app.getByRole('heading', { name: 'Modernization at a glance', exact: true }).count(), 0);
      }
      assert.ok(requests.every(request => request.method === 'GET'), 'Viewer startup must not write schema, roles, ACLs or records');
      assert.ok(requests.every(request => !request.path.includes('Feedback')), 'Startup must not require private feedback access');
      assert.deepEqual(unexpected, []); assert.deepEqual(errors, []);
    } catch (error) { await page.screenshot({ path: `test-artifacts/viewer-startup-${scenario}-failure.png` }).catch(() => {}); throw error; }
    finally { await page.close(); }
  }
  console.log(`${firepit ? 'Firepit opaque-frame/CSP bridge' : 'Viewer SharePoint'} startup smoke passed: prior schema loads read-only without writes; missing core fields, missing lists and denied reads remain explicit failures.`);
} finally { await browser.close(); }
