import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import { chromium } from 'playwright';

const schema = JSON.parse(await fs.readFile('scripts/fixtures/pre-update-editing-schema.json', 'utf8'));
schema.updates.push('EditedAt', 'EditorName', 'EditorKey');
const artifact = await fs.readFile('build-singlefile/modernization-project-tracker.html', 'utf8');
// All requests are intercepted below. Use the deployment URL to exercise its default folder.
const webUrl = 'https://flankspeed.sharepoint-mil.us/sites/metsoft';
const assetsFolder = '/sites/metsoft/Shared Documents/Modernization Tracker';
const trackerPage = `${webUrl}/SitePages/Modernization-Tracker.aspx`;
const html = artifact.replace('<head>', `<head><script>window.MOD_TRACKER_CONFIG=${JSON.stringify({ webUrl, appUrl: trackerPage, forceSharePoint: true, scopedAccess: true })}</script>`);
const suffixes = ['Projects', 'Tasks', 'Updates', 'Risks', 'Acronyms', 'Users', 'Reference Documents'];
const lists = Object.keys(schema).map((key, index) => ({ key, Id: `${String(index + 1).padStart(8, '0')}-1111-1111-1111-111111111111`, Title: `Modernization-Tracker - ${suffixes[index]}`, BaseTemplate: 100, Hidden: false }));
const manager = { Id: 10, Title: 'Tracker Manager', LoginName: 'i:0#.f|membership|manager@example.test', Email: 'manager@example.test', IsSiteAdmin: false };
const invitee = { Id: 11, Title: 'New Tracker Viewer', LoginName: 'i:0#.f|membership|new.viewer@example.test', Email: 'new.viewer@example.test' };
await fs.mkdir('test-artifacts', { recursive: true });
const browser = await chromium.launch();
try {
  for (const scenario of ['invite', 'missing-tracker-group']) {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
    page.setDefaultTimeout(15000);
    const writes = [], requests = [], unexpected = [], errors = [];
    let signedIn = manager;
    const groups = [
      { Id: 20, Title: 'Tracker Users', users: [manager] },
      { Id: 21, Title: 'Tracker Project Engineers', users: [] },
      { Id: 22, Title: 'Tracker Managers', users: [manager] },
      { Id: 30, Title: 'Metrology App User', users: [] },
    ];
    const records = Object.fromEntries(Object.keys(schema).map(key => [key, []]));
    records.users.push({ Id: 10, RecordId: 'manager', Title: manager.Title, LoginKey: manager.LoginName, Email: manager.Email, AppRole: 'Manager' });
    page.on('pageerror', error => errors.push(error.message));
    await page.route('**/*', async route => {
      const request = route.request(), url = new URL(request.url()), path = url.pathname.slice(new URL(webUrl).pathname.length);
      const decoded = decodeURIComponent(path);
      const send = (body, status = 200) => route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(body) });
      if (url.href === `${webUrl}/preview`) return route.fulfill({ contentType: 'text/html', body: '<iframe id="app" style="position:fixed;inset:0;width:100%;height:100%;border:0"></iframe>' });
      requests.push({ method: request.method(), path: decoded });
      const rawBody = request.postData();
      const body = rawBody && !path.endsWith('/ProcessQuery') ? JSON.parse(rawBody) : rawBody;
      if (request.method() === 'POST') writes.push({ path: decoded, body });
      if (path.endsWith('/contextinfo')) return send({ FormDigestValue: 'test-digest', FormDigestTimeoutSeconds: 1800 });
      if (path.toLowerCase().endsWith('/effectivebasepermissions')) return send({ Low: '1', High: '0' });
      if (path.endsWith('/currentuser/groups')) return send({ value: groups.filter(group => group.users.some(user => user.LoginName === signedIn.LoginName)).map(({ Id, Title }) => ({ Id, Title })) });
      if (path.endsWith('/currentuser')) return send(signedIn);
      if (path.endsWith('/_vti_bin/client.svc/ProcessQuery')) return send([{ ErrorInfo: null }]);
      if (path.endsWith('/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser')) return send({ ClientPeoplePickerSearchUser: JSON.stringify([{ Key: invitee.LoginName, DisplayText: invitee.Title, EntityData: { Email: invitee.Email }, IsResolved: true, EntityType: 'User' }]) });
      if (path.endsWith('/ensureuser')) return send(invitee);
      if (path.endsWith('/associatedownergroup')) return send({ Id: 99 });
      const groupName = decoded.match(/sitegroups\/getbyname\('([^']+)'\)/)?.[1];
      if (groupName) {
        const group = groups.find(group => group.Title === groupName && !(scenario === 'missing-tracker-group' && group.Id === 20));
        return group ? send({ Id: group.Id, Title: group.Title }) : send({ error: { message: 'Group not found' } }, 404);
      }
      const groupId = path.match(/sitegroups\((\d+)\)\/users/)?.[1];
      if (groupId) {
        const group = groups.find(group => group.Id === Number(groupId));
        if (request.method() === 'POST') {
          if (group.Id !== 20 || body.LoginName !== invitee.LoginName) unexpected.push(`Unexpected membership mutation: ${decoded}`);
          group.users.push(invitee); return send(invitee);
        }
        const filter = url.searchParams.get('$filter') || '';
        return send({ value: group.users.filter(user => !filter || filter.includes(user.LoginName)) });
      }
      if (decoded.includes('GetFolderByServerRelativePath') && !decoded.includes(`decodedurl='${assetsFolder}'`)) {
        return send({ error: { message: 'Folder not found' } }, 404);
      }
      if (path.includes('/getusereffectivepermissions')) return send({ GetUserEffectivePermissions: { Low: groups[0].users.some(user => user.LoginName === invitee.LoginName) ? '33' : '0' } });
      if (path.includes('GetFileByServerRelativePath')) return send({ Exists: true, Level: 1 });
      if (path.endsWith('/roledefinitions')) return send({ value: [{ Id: 124, RoleTypeKind: 2 }] });
      if (path.endsWith('/SP.Web.ShareObject')) return send({ StatusCode: 0 });
      if (path.endsWith('/lists')) {
        const filter = url.searchParams.get('$filter') || '';
        return send({ value: lists.filter(list => filter.includes(`'${list.Title}'`)) });
      }
      const id = path.match(/lists\(guid'([^']+)'\)/)?.[1];
      const list = lists.find(list => list.Id === id);
      if (list) {
        if (path.endsWith('/fields')) return send({ value: schema[list.key].map(InternalName => ({ InternalName })) });
        if (path.endsWith('/items')) {
          if (request.method() === 'POST') {
            if (list.key !== 'users') unexpected.push(`Unexpected record write: ${list.key}`);
            const row = { ...body, Id: 11 }; records[list.key].push(row); return send(row);
          }
          const selected = (url.searchParams.get('$select') || '').split(',');
          const unknown = selected.filter(name => !['Id', 'Title', 'FileSystemObjectType'].includes(name) && !schema[list.key].includes(name));
          if (unknown.length) unexpected.push(`Unexpected fields: ${unknown}`);
          return send({ value: records[list.key] });
        }
        const itemId = path.match(/\/items\((\d+)\)$/)?.[1];
        if (itemId) return send(records[list.key].find(row => row.Id === Number(itemId)));
        if (path.endsWith(`lists(guid'${id}')`)) return send({ Id: list.Id });
      }
      unexpected.push(`Unexpected ${request.method()}: ${decoded}`);
      return send({ error: { message: 'Unexpected mock request' } }, 400);
    });
    try {
      await page.goto(`${webUrl}/preview`);
      await page.locator('#app').evaluate((iframe, content) => { iframe.srcdoc = content; }, html);
      const app = page.frameLocator('#app');
      await app.getByRole('heading', { name: 'Modernization at a glance', exact: true }).waitFor();
      await app.locator('.sidebar nav').getByRole('button', { name: 'Users and managers', exact: true }).click();
      const form = app.getByRole('region', { name: 'Invite people', exact: true });
      await form.getByText(/adds them to Tracker Users/).waitFor();
      await form.getByRole('searchbox', { name: 'Find a person', exact: true }).fill('New Tracker');
      await form.getByRole('button', { name: /New Tracker Viewer/ }).click();
      await form.getByRole('combobox', { name: 'Invitation role', exact: true }).selectOption('Viewer');
      await form.getByRole('button', { name: 'Invite and grant access', exact: true }).click();
      if (scenario === 'invite') {
        await form.getByRole('status').filter({ hasText: 'SharePoint accepted the invitation email request' }).waitFor();
        assert.equal(groups[0].users.filter(user => user.LoginName === invitee.LoginName).length, 1);
        assert.ok(requests.some(({ path }) => path.includes(`GetFolderByServerRelativePath(decodedurl='${assetsFolder}')/ListItemAllFields/getusereffectivepermissions`)));
        const emails = writes.filter(write => write.path === '/_api/SP.Web.ShareObject');
        assert.equal(emails.length, 1);
        assert.equal(emails[0].body.url, trackerPage);
        assert.equal(emails[0].body.sendEmail, true); assert.equal(emails[0].body.propagateAcl, false);
        assert.equal(await form.getByRole('button', { name: 'Invitation requested', exact: true }).isDisabled(), true);
        await page.screenshot({ path: 'test-artifacts/tracker-users-invitation.png' });

        // Reopen as the newly enrolled Viewer. Basic access must not expose
        // management controls or require reading the private feedback list.
        signedIn = invitee;
        const viewerWriteStart = writes.length, viewerRequestStart = requests.length;
        await page.reload();
        await page.locator('#app').evaluate((iframe, content) => { iframe.srcdoc = content; }, html);
        await app.getByRole('heading', { name: 'Modernization at a glance', exact: true }).waitFor();
        assert.equal(await app.getByRole('button', { name: 'New project', exact: true }).count(), 0);
        assert.equal(await app.getByRole('button', { name: 'Users and managers', exact: true }).count(), 0);
        assert.equal(await app.getByRole('heading', { name: 'Users and managers', exact: true }).count(), 0);
        assert.equal(await app.getByRole('region', { name: 'Invite people', exact: true }).count(), 0);
        assert.equal(await app.getByRole('button', { name: 'Edit', exact: true }).count(), 0);
        assert.equal(writes.length, viewerWriteStart, 'Viewer startup must not provision or enroll users');
        await app.getByRole('button', { name: 'Report a bug or request a feature' }).click();
        await app.getByLabel('Summary', { exact: true }).fill('Tracker-only Viewer feedback');
        await app.getByLabel('What happened?', { exact: true }).fill('A newly invited Tracker Viewer can submit private feedback.');
        await app.getByRole('button', { name: 'Send feedback', exact: true }).click();
        await app.getByRole('heading', { name: 'Thank you for the feedback' }).waitFor();
        assert.equal(await app.getByRole('button', { name: 'Feedback History', exact: true }).count(), 0);
        const viewerWrites = writes.slice(viewerWriteStart).filter(write => !write.path.endsWith('/contextinfo'));
        assert.deepEqual(viewerWrites.map(write => write.path), ['/_vti_bin/client.svc/ProcessQuery']);
        assert.match(viewerWrites[0].body, /Tracker-only Viewer feedback/);
        assert.ok(requests.slice(viewerRequestStart).every(request => !request.path.includes('Feedback')), 'Viewer must never read private submissions');
        await page.screenshot({ path: 'test-artifacts/tracker-users-viewer-feedback.png' });
      } else {
        await form.getByRole('alert').filter({ hasText: 'Tracker Users' }).waitFor();
        assert.equal(writes.filter(write => write.path.includes('/sitegroups') || write.path.includes('ShareObject')).length, 0);
        assert.equal(await form.getByRole('button', { name: 'Retry invitation', exact: true }).isEnabled(), true);
      }
      assert.equal(records.users.filter(row => row.LoginKey === invitee.LoginName && row.AppRole === 'Viewer').length, 1);
      assert.deepEqual(groups[3].users, []);
      assert.ok(requests.every(request => !/Metrology App User|sitegroups\(30\)|associatedmembergroup/.test(request.path)));
      assert.deepEqual(unexpected, []); assert.deepEqual(errors, []);
    } catch (error) {
      await page.screenshot({ path: `test-artifacts/tracker-users-${scenario}-failure.png` }).catch(() => {});
      console.error({ scenario, unexpected, errors, alerts: await page.frameLocator('#app').getByRole('alert').allTextContents().catch(() => []) });
      throw error;
    }
    finally { await page.close(); }
  }
  console.log('Tracker Users invitation smoke passed: non-owner Manager enrolls only Tracker Users and requests one Tracker-page email; invited Viewer opens read-only and submits private feedback; missing group never falls back to suite membership.');
} finally { await browser.close(); }
