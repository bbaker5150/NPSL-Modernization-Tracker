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
const errors = [];
page.on('pageerror', error => errors.push(error.message));
const frame = page.frameLocator('#app');
async function boot() {
  await page.reload();
  await page.locator('#app').evaluate((el, value) => { el.srcdoc = value; }, html);
  await frame.getByRole('heading', { name: 'Modernization at a glance', exact: true }).waitFor();
}
async function role(value) {
  await page.evaluate(value => { const key = 'modernization-project-tracker:v2'; const data = JSON.parse(localStorage.getItem(key)); data.users[0].role = value; localStorage.setItem(key, JSON.stringify(data)); }, value);
  await boot();
}
try {
  await page.goto(`http://127.0.0.1:${server.address().port}`);
  await page.evaluate(() => {
    const projects = ['Electrical standards', 'Pressure modernization'].map((title, i) => ({ id: `p${i}`, projectKey: `p${i}`, title, measurementArea: i ? 'Pressure' : 'AC Voltage', ownerName: i ? 'Other Engineer' : 'Local Engineer', ownerKey: i ? 'other' : 'local', ownerEmail: i ? 'other@example.test' : 'local.engineer@example.invalid', currentStageKey: 'requirement', health: 'On Track', priority: 'Medium', targetFinish: '2027-09-30' }));
    const tasks = Array.from({ length: 6 }, (_, i) => ({ id: `t${i}`, projectKey: `p${i % 2}`, phaseKey: 'requirement', title: ['Review requirements and confirm technical approach', 'Resolve procurement funding dependency', 'Prepare calibration standard specification', 'Coordinate stakeholder review', 'Completed work', 'Excluded work'][i], status: ['In Progress', 'Blocked', 'Not Started', 'In Progress', 'Complete', 'Not Required'][i], ownerName: i % 2 ? 'Jordan Engineer' : 'Local Engineer', ownerKey: 'local', organization: 'NPSL', dueDate: '2026-09-15', deferredDate: i === 1 ? '2027-01-15' : '', deferredJustification: i === 1 ? 'Awaiting funding release at the next program review.' : '', estimatedHours: i === 2 ? '' : i === 3 ? 0 : 8 + i * 4, notes: 'Confirm the technical requirements with the project team before proceeding.', assignedDate: '2026-08-01' }));
    localStorage.setItem('modernization-project-tracker:v2', JSON.stringify({ projects, tasks, updates: [], risks: [], users: [{ id: 'local', title: 'Local Engineer', loginName: 'local', email: 'local.engineer@example.invalid', role: 'Manager' }, { id: 'other', title: 'Other Engineer', loginName: 'other', email: 'other@example.test', role: 'Project Engineer' }], glossary: [], references: [], recycleBin: [] }));
  });
  await boot();
  await frame.locator('.kpi-card').filter({ hasText: 'NPSL Needs Attention' }).click();
  assert.equal(await frame.locator('.attention-task-card').count(), 4);
  assert.equal(await frame.locator('.attention-summary strong').first().innerText(), '4');
  assert.equal(await frame.locator('.attention-task-details[open]').count(), 0);
  const emptyEstimate = frame.locator('.attention-task-card').filter({ hasText: 'Prepare calibration standard specification' });
  assert.equal(await emptyEstimate.getByText('Est. hours', { exact: true }).count(), 0);
  assert.equal(await emptyEstimate.getByText('Deferred to', { exact: true }).count(), 0);
  const zeroEstimate = frame.locator('.attention-task-card').filter({ hasText: 'Coordinate stakeholder review' });
  assert.equal(await zeroEstimate.locator('.attention-task-facts > div').filter({ hasText: 'Est. hours' }).locator('dd').innerText(), '0');
  for (const width of [1440, 768, 390]) {
    await page.setViewportSize({ width, height: 1050 });
    for (const theme of ['light', 'dark']) {
      await frame.locator('html').evaluate((el, theme) => { el.dataset.theme = theme; }, theme);
      await page.waitForTimeout(250); // Let theme transitions settle before visual review.
      assert.ok(await frame.locator('.attention-page').evaluate(el => el.scrollWidth <= el.clientWidth + 1), `Attention overflow at ${width}px/${theme}`);
      const columns = await frame.locator('.attention-groups-status > .attention-section').evaluateAll(elements => elements.map(el => { const r = el.getBoundingClientRect(); return { top: r.top, left: r.left, right: r.right, width: r.width }; }));
      assert.equal(columns.length, 3);
      assert.ok(columns.every(column => Math.abs(column.top - columns[0].top) < 2), 'Statuses must share a horizontal row');
      assert.ok(columns.every((column, i) => column.width >= 290 && (!i || column.left > columns[i - 1].right)), 'Columns must remain readable without overlapping');
      assert.ok(await frame.locator('.attention-task-card').evaluateAll(cards => cards.every(card => card.scrollWidth <= card.clientWidth + 1)), 'Task card content must fit');
      assert.notEqual(await frame.getByLabel('Search attention tasks').evaluate(el => getComputedStyle(el, '::placeholder').color), 'rgba(0, 0, 0, 0)');
      await page.screenshot({ path: `test-artifacts/attention-${width}-${theme}.png`, fullPage: true });
    }
  }
  await page.setViewportSize({ width: 1440, height: 1050 });
  await frame.locator('.attention-groups-status > .attention-section > summary').first().click();
  assert.equal(await frame.locator('.attention-groups-status > .attention-section[open]').count(), 2);
  await frame.locator('.attention-groups-status > .attention-section > summary').first().click();
  assert.equal(await frame.locator('.attention-groups-status > .attention-section[open]').count(), 3);
  await frame.getByRole('button', { name: 'By project', exact: true }).click();
  assert.equal(await frame.locator('.attention-section').count(), 2);
  await frame.locator('.attention-section > summary').first().click();
  assert.equal(await frame.locator('.attention-section[open]').count(), 1);
  await frame.getByLabel('Search attention tasks').fill('funding');
  assert.equal(await frame.locator('.attention-task-card').count(), 1);
  await role('Project Engineer');
  assert.equal(await frame.locator('.project-table-row:not(.table-header)').count(), 1);
  await frame.locator('.project-table-row:not(.table-header)').click();
  await frame.getByRole('button', { name: 'Project settings', exact: true }).click();
  await frame.getByRole('button', { name: 'Edit project', exact: true }).click();
  assert.ok(await frame.getByLabel('Target completion', { exact: true }).isDisabled());
  assert.equal(await frame.locator('.modal input[type="checkbox"]').count(), 0);
  await page.screenshot({ path: 'test-artifacts/project-engineer-editor.png' });
  await frame.getByRole('button', { name: 'Cancel', exact: true }).click();
  await frame.getByRole('button', { name: /Work breakdown/ }).click();
  await frame.getByRole('button', { name: /Add task to/ }).first().click();
  await frame.getByLabel('Task name', { exact: true }).fill('New task with documents');
  await frame.getByLabel('Notes', { exact: true }).fill('Review the attached requirements before proceeding.');
  assert.ok(await frame.getByLabel('Original due date').isDisabled());
  assert.equal(await frame.getByLabel('Pipeline stage', { exact: true }).count(), 0);
  await frame.getByRole('combobox', { name: 'Task owner', exact: true }).fill('Other');
  await frame.getByRole('option', { name: /Other Engineer/ }).click();
  await frame.getByLabel('Attach documents', { exact: true }).setInputFiles({ name: 'requirements.txt', mimeType: 'text/plain', buffer: Buffer.from('Task requirements') });
  await page.screenshot({ path: 'test-artifacts/task-create-with-document.png' });
  await frame.getByRole('button', { name: 'Save task', exact: true }).click();
  await frame.locator('.modal').waitFor({ state: 'detached' });
  const saved = await page.evaluate(() => JSON.parse(localStorage.getItem('modernization-project-tracker:v2')).tasks.find(row => row.title === 'New task with documents'));
  assert.equal(saved.ownerKey, 'other');
  assert.equal(saved.phaseKey, 'requirement');
  assert.equal(saved.dueDate, '');
  await frame.getByText('New task with documents', { exact: true }).click();
  await frame.getByText('requirements.txt', { exact: true }).waitFor();
  await frame.getByRole('button', { name: 'Cancel', exact: true }).click();
  await frame.locator('.project-drawer').getByRole('button', { name: 'Close', exact: true }).click();
  await frame.locator('.kpi-card').filter({ hasText: 'NPSL Needs Attention' }).click();
  const documentCard = frame.locator('.attention-task-card').filter({ hasText: 'New task with documents' });
  assert.equal(await documentCard.getByText('Original due', { exact: true }).count(), 0);
  assert.equal(await documentCard.getByText('Est. hours', { exact: true }).count(), 0);
  await documentCard.locator('.attention-task-details > summary').click();
  await documentCard.getByRole('button', { name: 'requirements.txt', exact: true }).waitFor();
  const layout = await documentCard.evaluate(card => {
    const summary = card.querySelector('.attention-task-details > summary').getBoundingClientRect();
    const content = card.querySelector('.attention-details-content').getBoundingClientRect();
    const notes = [...card.querySelectorAll('p')].find(el => el.textContent.includes('Review the attached requirements')).getBoundingClientRect();
    const documents = card.querySelector('.task-documents').getBoundingClientRect();
    return { gap: content.top + parseFloat(getComputedStyle(card.querySelector('.attention-details-content')).paddingTop) - summary.bottom, belowNotes: documents.top > notes.bottom };
  });
  assert.ok(layout.gap >= 14 && layout.belowNotes, 'Expanded details need spacing and documents below notes');
  const download = page.waitForEvent('download');
  await documentCard.getByRole('button', { name: 'requirements.txt', exact: true }).click();
  assert.equal(await fs.readFile(await (await download).path(), 'utf8'), 'Task requirements');
  for (const [width, theme] of [[1440, 'light'], [390, 'dark']]) {
    await page.setViewportSize({ width, height: 1050 });
    await frame.locator('html').evaluate((el, value) => { el.dataset.theme = value; }, theme);
    await page.waitForTimeout(250);
    await documentCard.scrollIntoViewIfNeeded();
    assert.ok(await documentCard.evaluate(el => el.scrollWidth <= el.clientWidth + 1));
    await documentCard.screenshot({ path: `test-artifacts/task-documents-${width}-${theme}.png` });
  }
  await page.setViewportSize({ width: 1440, height: 1050 });
  await role('Viewer');
  assert.equal(await frame.locator('.project-table-row:not(.table-header)').count(), 2);
  await frame.locator('.project-table-row:not(.table-header)').first().click();
  assert.equal(await frame.getByRole('button', { name: 'Project settings', exact: true }).count(), 0);
  await frame.locator('.upcoming-list > button').first().click();
  assert.equal(await frame.getByRole('button', { name: 'Save task', exact: true }).count(), 0);
  assert.ok(await frame.getByLabel('Task name', { exact: true }).isDisabled());
  assert.deepEqual(errors, []);
  console.log('Tasking smoke passed: responsive attention, project grouping, role scope, locked deadlines, owner autocomplete and pre-save document persistence.');
} catch (error) {
  await page.screenshot({ path: 'test-artifacts/tasking-failure.png' }).catch(() => {});
  console.error(error);
  process.exitCode = 1;
} finally {
  await browser.close();
  server.close();
}
