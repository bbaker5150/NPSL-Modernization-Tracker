import fs from 'node:fs/promises';
import http from 'node:http';
import assert from 'node:assert/strict';
import { chromium } from 'playwright';
const html = await fs.readFile('build-singlefile/modernization-project-tracker.html', 'utf8');
await fs.mkdir('test-artifacts', { recursive: true });
const server = http.createServer((req, res) => { res.setHeader('Content-Type', 'text/html'); res.end('<iframe id="app" style="position:fixed;inset:0;width:100%;height:100%;border:0"></iframe>'); });
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 1050 } });
page.setDefaultTimeout(20000);
const app = page.frameLocator('#app'), errors = [];
page.on('pageerror', err => errors.push(err.message));
async function boot() {
  await page.reload();
  await page.locator('#app').evaluate((el, html) => { el.srcdoc = html; }, html);
  await app.getByRole('heading', { name: 'Modernization at a glance', exact: true }).waitFor();
  await app.locator('body').evaluate(el => el.ownerDocument.addEventListener('click', event => { if (event.target.closest('.project-updates button')) event.preventDefault(); }, true));
}
async function openUpdates() {
  await app.locator('.project-table-row:not(.table-header)').click();
  assert.equal(await app.locator('.drawer-tabs').getByRole('button', { name: /^Risks/ }).count(), 0);
  await app.locator('.drawer-tabs').getByRole('button', { name: /^Updates/ }).click();
}
try {
  await page.goto(`http://127.0.0.1:${server.address().port}`);
  await page.evaluate(() => localStorage.setItem('modernization-project-tracker:v2', JSON.stringify({
    projects: [{ id: 'p', projectKey: 'p', title: 'METCAL modernization', organization: 'PMA 260B', ownerName: 'Local Engineer', ownerKey: 'local', currentStageKey: 'requirement', health: 'On Track', priority: 'Medium' }],
    tasks: ['Chief Engineer Office', 'Team 2 Lead', 'Team 3 Lead', ''].map((subAssignment, index) => ({ id: `t${index}`, projectKey: 'p', title: `Technical review ${index + 1}`, organization: 'PMA 260B', subAssignment, phaseKey: 'requirement', status: 'In Progress', dueDate: '2027-06-01', assignedDate: '2026-10-01', ownerKey: 'local', ownerName: 'Local Engineer' })),
    updates: [], risks: [{ id: 'r', projectKey: 'p', title: 'Legacy supply risk', severity: 'High', status: 'Open', ownerName: 'Jordan Engineer', mitigation: 'Review vendors' }],
    users: [{ id: 'local', title: 'Local Engineer', loginName: 'local', email: 'local.engineer@example.invalid', role: 'Manager' }], acronyms: [], references: []
  })));
  await boot();
  assert.equal(await app.getByText('PMA 260B', { exact: true }).count(), 0);
  await app.locator('.kpi-card').filter({ hasText: 'NAVAIR METCAL 260B' }).click();
  assert.equal(await app.locator('.attention-task-card').count(), 4);
  await app.getByRole('button', { name: 'Team 2 Lead', exact: true }).click();
  assert.equal(await app.locator('.attention-task-card').count(), 1);
  await app.locator('.attention-responsibility').getByText('Team 2 Lead', { exact: true }).waitFor();
  await page.screenshot({ path: 'test-artifacts/metcal-responsibility.png' });
  await app.getByRole('button', { name: 'Technical review 2', exact: true }).click();
  await app.getByRole('combobox', { name: 'METCAL responsibility' }).selectOption('Chief Engineer Office');
  await app.getByRole('button', { name: 'Save task', exact: true }).click();
  await app.getByRole('heading', { name: 'Nothing needs attention' }).waitFor();
  await app.locator('.sidebar nav').getByRole('button', { name: 'Portfolio', exact: true }).click();
  await openUpdates();
  await app.getByText(/Legacy supply risk/).waitFor();
  await app.getByRole('combobox', { name: 'Task reference', exact: true }).selectOption('t1');
  await app.getByRole('textbox', { name: 'Your update', exact: true }).fill('Reviewed the engineering package. Team 2 will coordinate the next handoff.');
  await app.getByLabel('Update documents', { exact: true }).setInputFiles({ name: 'update-only.txt', mimeType: 'text/plain', buffer: Buffer.from('Engineering decision document') });
  await app.getByRole('button', { name: 'Post update', exact: true }).click();
  await app.getByText('Update posted with all selected documents.', { exact: true }).waitFor();
  const card = app.locator('.update-card').filter({ hasText: 'Reviewed the engineering package.' });
  await card.getByRole('button', { name: 'update-only.txt', exact: true }).waitFor();
  assert.ok((await card.locator('time').getAttribute('datetime')).includes('T'));
  assert.equal(await card.locator('header strong').innerText(), 'Local Engineer');
  const download = page.waitForEvent('download');
  await card.getByRole('button', { name: 'update-only.txt', exact: true }).click();
  const downloaded = await download;
  assert.equal(await fs.readFile(await downloaded.path(), 'utf8'), 'Engineering decision document');
  await app.getByRole('textbox', { name: 'Your update', exact: true }).fill('Project-wide decision with no task required.');
  await app.getByRole('button', { name: 'Post update', exact: true }).click();
  await app.getByText('Update posted.', { exact: true }).waitFor();
  assert.equal(await app.locator('.update-card').filter({ hasText: 'Project-wide decision' }).locator('.update-task-reference').count(), 0);
  for (const width of [1440, 390]) {
    await page.setViewportSize({ width, height: 1050 });
    for (const theme of ['light', 'dark']) {
      await app.locator('html').evaluate((el, value) => { el.dataset.theme = value; }, theme);
      await page.waitForTimeout(250); // Let theme colors settle for visual review.
      assert.ok(await app.locator('.project-updates').evaluate(el => el.scrollWidth <= el.clientWidth + 1));
      await page.screenshot({ path: `test-artifacts/project-updates-${width}-${theme}.png` });
    }
  }
  assert.equal(await app.getByRole('textbox', { name: 'Your update', exact: true }).isEnabled(), true);
  await app.locator('.update-history').scrollIntoViewIfNeeded();
  await page.screenshot({ path: 'test-artifacts/project-update-history-mobile.png' });
  await page.setViewportSize({ width: 1440, height: 1050 });
  await app.locator('.drawer-tabs').getByRole('button', { name: 'Overview', exact: true }).click();
  assert.equal(await app.locator('.project-documents').getByText('update-only.txt').count(), 0);
  await page.evaluate(() => { const key = 'modernization-project-tracker:v2', data = JSON.parse(localStorage.getItem(key)); data.users[0].role = 'Project Engineer'; localStorage.setItem(key, JSON.stringify(data)); });
  await boot();
  await openUpdates();
  await app.getByRole('button', { name: 'update-only.txt', exact: true }).waitFor();
  await app.locator('.update-card').filter({ hasText: 'Reviewed the engineering package.' }).getByRole('button', { name: 'Technical review 2', exact: true }).click();
  assert.equal(await app.getByLabel('Original due date', { exact: true }).isDisabled(), true);
  assert.equal(await app.getByLabel('Assigned / creation date', { exact: true }).isDisabled(), true);
  assert.equal(await app.getByLabel('Deferred date', { exact: true }).isEnabled(), true);
  await app.getByLabel('Deferred date', { exact: true }).fill('2027-07-01');
  await app.getByRole('textbox', { name: 'Deferral justification', exact: true }).fill('Awaiting delivery');
  await app.getByRole('button', { name: 'Save task', exact: true }).click();
  const saved = await page.evaluate(() => JSON.parse(localStorage.getItem('modernization-project-tracker:v2')).tasks.find(row => row.id === 't1'));
  assert.equal(saved.assignedDate, '2026-10-01'); assert.equal(saved.dueDate, '2027-06-01'); assert.equal(saved.deferredDate, '2027-07-01');
  assert.equal(saved.subAssignment, 'Chief Engineer Office');
  const originalCard = app.locator('.update-card').filter({ hasText: 'Reviewed the engineering package.' });
  const originalTime = await originalCard.locator('header time').getAttribute('datetime');
  await originalCard.getByRole('button', { name: 'Edit update', exact: true }).click();
  await app.getByRole('textbox', { name: 'Edit update', exact: true }).fill('Discarded edit');
  await app.getByRole('button', { name: 'Cancel edit', exact: true }).click();
  await originalCard.getByRole('button', { name: 'Edit update', exact: true }).click();
  await app.getByRole('textbox', { name: 'Edit update', exact: true }).fill('Revised engineering decision. Package accepted by the project owner.');
  await app.getByRole('combobox', { name: 'Edit task reference', exact: true }).selectOption('');
  for (const width of [1440, 390]) {
    await page.setViewportSize({ width, height: 1050 });
    await app.locator('.update-editor').scrollIntoViewIfNeeded();
    await page.screenshot({ path: `test-artifacts/update-editor-${width}.png` });
    assert.ok(await app.locator('.project-updates').evaluate(el => el.scrollWidth <= el.clientWidth + 1));
  }
  await app.getByRole('button', { name: 'Save changes', exact: true }).click();
  const revised = app.locator('.update-card').filter({ hasText: 'Revised engineering decision.' });
  await revised.getByText(/Edited by Local Engineer/).waitFor();
  assert.equal(await revised.locator('header strong').innerText(), 'Local Engineer');
  assert.equal(await revised.locator('header time').getAttribute('datetime'), originalTime);
  assert.equal(await revised.locator('.update-task-reference').count(), 0);
  await revised.getByRole('button', { name: 'update-only.txt', exact: true }).waitFor();
  await revised.getByRole('button', { name: 'Delete update', exact: true }).click();
  await revised.getByRole('button', { name: 'Keep update', exact: true }).click();
  await boot(); await openUpdates();
  await revised.getByText(/Edited by Local Engineer/).waitFor();
  await revised.getByRole('button', { name: 'Delete update', exact: true }).click();
  await app.locator('.update-delete-confirm').scrollIntoViewIfNeeded();
  await page.screenshot({ path: 'test-artifacts/update-delete-confirm-mobile.png' });
  await revised.getByRole('button', { name: 'Confirm delete', exact: true }).click();
  await revised.waitFor({ state: 'detached' });
  await boot(); await openUpdates();
  assert.equal(await app.locator('.update-card').count(), 2); // Other update + historical risk.
  const legacyRisk = app.locator('.update-card').filter({ hasText: 'Risk owner' });
  await legacyRisk.getByRole('button', { name: 'Edit risk', exact: true }).click();
  await legacyRisk.getByRole('textbox', { name: 'Mitigation', exact: true }).fill('Expedite vendor review.');
  await legacyRisk.getByRole('button', { name: 'Save changes' }).click();
  await legacyRisk.getByText(/Expedite vendor review/).waitFor();
  assert.equal(await app.getByRole('button', { name: 'update-only.txt', exact: true }).count(), 0);
  await app.getByText('Project-wide decision with no task required.', { exact: true }).waitFor();
  await page.evaluate(() => { const key = 'modernization-project-tracker:v2', data = JSON.parse(localStorage.getItem(key)); data.users[0].role = 'Viewer'; localStorage.setItem(key, JSON.stringify(data)); });
  await boot(); await openUpdates();
  assert.equal(await app.getByRole('button', { name: 'Edit update', exact: true }).count(), 0);
  assert.equal(await app.getByRole('button', { name: 'Delete update', exact: true }).count(), 0);
  assert.equal(await app.getByRole('button', { name: 'Edit risk', exact: true }).count(), 0);
  assert.equal(await app.getByRole('button', { name: 'Delete risk', exact: true }).count(), 0);

  assert.deepEqual(errors, []);
  console.log('Project updates smoke passed: METCAL rename/filter/subassignment, engineer date restrictions with editable deferrals, no Risks tab, optional task links, update-only attachments/downloads, author/timestamps, per-update edits/deletes with author preservation, cancel/save/reload, Viewer restrictions, persistence, desktop/mobile themes.');
} catch (err) {
  await page.screenshot({ path: 'test-artifacts/project-updates-failure.png' }).catch(() => {});
  throw err;
} finally { await browser.close(); await new Promise(resolve => server.close(resolve)); }
