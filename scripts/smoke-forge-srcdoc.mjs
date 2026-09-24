import fs from 'node:fs/promises';
import http from 'node:http';
import path from 'node:path';
import { chromium } from 'playwright';

const artifactPath = path.resolve('build-singlefile/modernization-project-tracker.html');
const artifact = await fs.readFile(artifactPath, 'utf8');
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const errors = [];
const externalRequests = [];
const dialogs = [];
const server = http.createServer((_request, response) => {
  response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  response.end('<!doctype html><html><body style="margin:0"><iframe id="app" style="width:100vw;height:100vh;border:0"></iframe></body></html>');
});
await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const { port } = server.address();
const harnessOrigin = `http://127.0.0.1:${port}`;

page.on('console', (message) => { if (message.type() === 'error') errors.push(message.text()); });
page.on('pageerror', (error) => errors.push(error.message));
page.on('dialog', async (dialog) => { dialogs.push(`${dialog.type()}: ${dialog.message()}`); await dialog.dismiss(); });
page.on('request', (request) => {
  const url = request.url();
  if (/^https?:/i.test(url) && !url.startsWith(harnessOrigin)) externalRequests.push(url);
});

try {
  // A real origin matters: an about:blank parent gives its srcdoc child an
  // opaque origin, which forbids localStorage before the app can render. Forge
  // hosts the parent on SharePoint, so serve the harness from localhost to
  // reproduce that same-origin behavior.
  await page.goto(`${harnessOrigin}/`);
  await page.locator('#app').evaluate((frame, html) => { frame.srcdoc = html; }, artifact);
  const frame = page.frameLocator('#app');

  await frame.getByRole('heading', { name: 'My work', exact: true }).waitFor({ timeout: 20_000 });
  if (await frame.getByRole('button', { name: /New project/ }).count()) errors.push('Standard user could create projects');
  await frame.getByRole('button', { name: 'Users and managers', exact: true }).click();
  await frame.getByLabel('Testing password').fill('incorrect-password');
  await frame.getByRole('button', { name: 'Enable manager access' }).click();
  await frame.getByRole('alert').filter({ hasText: 'Testing password is incorrect' }).waitFor();
  if (await frame.getByRole('button', { name: /New project/ }).count()) errors.push('Incorrect password granted manager access');
  await frame.getByLabel('Testing password').fill('admin123');
  await frame.getByRole('button', { name: 'Enable manager access' }).click();
  await frame.getByRole('button', { name: /New project/ }).waitFor();
  // Reload the packaged application to verify persisted role, not just UI state.
  await page.reload();
  await page.locator('#app').evaluate((element, html) => { element.srcdoc = html; }, artifact);
  await frame.getByRole('button', { name: /New project/ }).waitFor();
  await frame.getByRole('button', { name: 'Portfolio', exact: true }).click();
  await frame.getByRole('heading', { name: 'Modernization at a glance' }).waitFor();
  await frame.getByText('0 total measurement areas').waitFor();
  const phaseLabels = await frame.locator('.phase-label').allTextContents();
  if (phaseLabels.join('|') !== 'Requirement (MSA)|Acquisition (TMRR)|Procurement (EMD)|Deployment (P&D)') errors.push(`Incorrect phase sequence: ${phaseLabels}`);
  const tracks = await frame.locator('.pipeline-rail').evaluate((row) => getComputedStyle(row).gridTemplateColumns.split(' ').length);
  if (tracks !== 4) errors.push(`Expected four pipeline columns, got ${tracks}`);
  if (await frame.locator('.portfolio-register').count()) errors.push('Empty portfolio register remained visible');
  if (await frame.getByText('Add first project').count()) errors.push('Duplicate new project prompt remained visible');
  if (await frame.locator('.topbar .search-box').count()) errors.push('Global search remained visible');
  const emptyBarWidths = await frame.locator('.phase-bar span').evaluateAll((bars) => bars.map((bar) => bar.style.width));
  if (emptyBarWidths.some((width) => width !== '0%')) errors.push(`Empty pipeline stages displayed progress: ${emptyBarWidths.join(', ')}`);
  if (await frame.getByText('Open sample portfolio').count()) errors.push('Sample portfolio control remained visible');

  await frame.getByRole('button', { name: /Acronym glossary/ }).click();
  await frame.getByText('Calibration Standard Specification').waitFor();
  await frame.getByLabel('Acronym', { exact: true }).fill('SMK');
  await frame.getByLabel('Full term').fill('Smoke Test Glossary Entry');
  await frame.getByLabel('Definition').fill('Validates shared glossary creation.');
  await frame.getByRole('button', { name: 'Add acronym' }).click();
  await frame.getByText('Smoke Test Glossary Entry').waitFor();

  await frame.getByRole('button', { name: /Portfolio/, exact: true }).click();
  await frame.getByRole('button', { name: /New project/ }).click();
  await frame.locator('.modal .field input').first().fill('Smoke test modernization project');
  await frame.getByRole('button', { name: /Save project/ }).click();
  await frame.locator('.project-drawer').waitFor();
  await frame.getByLabel('Progress calculation').selectOption('tasks');
  await frame.getByText('0 of 4 tasks completed', { exact: true }).waitFor();
  if (await frame.locator('.task-markers > span').count() !== 4) errors.push('Task progress must display one marker per task');
  if (await frame.getByRole('button', { name: 'Claim project', exact: true }).count()) errors.push('Claim project remained visible');
  await frame.getByRole('button', { name: 'Project settings', exact: true }).click();
  await frame.getByRole('button', { name: 'Edit project', exact: true }).waitFor();
  await frame.getByRole('button', { name: 'Project settings', exact: true }).click();
  // Removing the task-code cell must not leave titles in its old 48px track.
  // Exercise realistic long titles in the packaged iframe, at multiple sizes.
  const upcomingTitle = frame.locator('.upcoming-list > button strong').first();
  const originalTitle = await upcomingTitle.textContent();
  await upcomingTitle.evaluate((element) => { element.textContent = 'Develop Capability Development Document (CDD/CSS) and incorporate comments from the technical review team'; });
  for (const width of [1440, 768, 390]) {
    await page.setViewportSize({ width, height: 1000 });
    for (const theme of ['light', 'dark']) {
      await frame.locator('html').evaluate((element, value) => { element.dataset.theme = value; }, theme);
      const caption = await frame.locator('.progress-caption').evaluate((row) => {
        const [percent, title] = [row.querySelector('span'), row.querySelector('strong')];
        return { overlap: percent.getBoundingClientRect().right > title.getBoundingClientRect().left, overflow: row.scrollWidth > row.clientWidth + 1, sameFont: getComputedStyle(percent).fontSize === getComputedStyle(title).fontSize };
      });
      if (caption.overlap || caption.overflow || !caption.sameFont) errors.push(`Progress caption layout failed at ${width}px in ${theme}`);
      const layout = await frame.locator('.upcoming-list > button').first().evaluate((row) => {
        const title = row.querySelector('div').getBoundingClientRect();
        const badge = row.querySelector('.badge').getBoundingClientRect();
        const bounds = row.getBoundingClientRect();
        return { titleWidth: title.width, badgeWidth: badge.width, rowWidth: bounds.width,
          overlap: title.right > badge.left, overflow: row.scrollWidth > row.clientWidth + 1 || badge.right > bounds.right + 1 };
      });
      if (layout.titleWidth < layout.rowWidth * 0.55 || layout.badgeWidth > layout.rowWidth * 0.4 || layout.overlap || layout.overflow) {
        errors.push(`Upcoming work layout failed at ${width}px in ${theme}: ${JSON.stringify(layout)}`);
      }
    }
  }
  await upcomingTitle.evaluate((element, value) => { element.textContent = value; }, originalTitle);
  await page.setViewportSize({ width: 1440, height: 1000 });
  await frame.getByRole('button', { name: /Work breakdown/ }).click();
  await frame.locator('.phase-task').first().waitFor();
  const quickComplete = frame.locator('.task-check:not(.checked)').first();
  const quickCompleteLabel = await quickComplete.getAttribute('aria-label');
  await quickComplete.click();
  const reopenedTask = frame.getByRole('button', { name: quickCompleteLabel.replace('Complete', 'Reopen'), exact: true });
  await reopenedTask.waitFor();
  await reopenedTask.locator('xpath=..').click({ position: { x: 110, y: 12 } });
  await frame.getByRole('heading', { name: 'Update task' }).waitFor();
  await frame.locator('.modal select option', { hasText: 'Not Required' }).waitFor({ state: 'attached' });
  await frame.locator('.modal select').filter({ has: frame.locator('option', { hasText: 'Not Required' }) }).selectOption({ label: 'Not Required' });
  await frame.getByLabel('Not required justification').fill('Not applicable to this modernization project');
  await frame.getByRole('button', { name: 'Save task' }).click();
  await frame.getByRole('button', { name: quickCompleteLabel.replace('Complete', 'Reopen'), exact: true }).waitFor();
  await frame.locator('.project-drawer').getByRole('button', { name: 'Close', exact: true }).click();
  // The register should span the same content width as the pipeline panel.
  for (const width of [1440, 768]) {
    await page.setViewportSize({ width, height: 1000 });
    for (const theme of ['light', 'dark']) {
      await frame.locator('html').evaluate((element, value) => { element.dataset.theme = value; }, theme);
      const layout = await frame.locator('.portfolio-register').evaluate((register) => {
        const bounds = register.getBoundingClientRect();
        const pipeline = document.querySelector('.pipeline-panel').getBoundingClientRect();
        const select = getComputedStyle(register.querySelector('select'));
        const option = getComputedStyle(register.querySelector('option'));
        return { widthDelta: Math.abs(bounds.width - pipeline.width), leftDelta: Math.abs(bounds.left - pipeline.left), selectColor: select.color, selectBackground: select.backgroundColor, optionColor: option.color, optionBackground: option.backgroundColor };
      });
      if (layout.widthDelta > 2 || layout.leftDelta > 2 || layout.selectColor === layout.selectBackground || layout.optionColor === layout.optionBackground || layout.optionBackground === 'rgba(0, 0, 0, 0)') errors.push(`Portfolio layout/theme failure at ${width}px (${theme}): ${JSON.stringify(layout)}`);
    }
  }
  await page.setViewportSize({ width: 1440, height: 1000 });
  await frame.getByRole('button', { name: /Pipeline board/ }).click();
  const beforeDelete = await frame.locator('.project-card').count();
  await frame.locator('.project-card .project-menu').first().click();
  await frame.locator('.project-menu-popover button').click();
  await frame.locator('.project-card').nth(beforeDelete - 1).waitFor({ state: 'detached' });
  const afterDelete = await frame.locator('.project-card').count();
  if (afterDelete !== beforeDelete - 1) errors.push(`Prompt-free project deletion did not update immediately (${beforeDelete} -> ${afterDelete})`);
  if (dialogs.length) errors.push(`Native browser dialog(s) displayed: ${dialogs.join(' | ')}`);
  const overlays = await frame.locator('#pdc-open, #test-recorder-launcher, #test-recorder-panel, .test-recorder-ui').evaluateAll((elements) => elements.filter((element) => getComputedStyle(element).display !== 'none').length);
  if (overlays) errors.push(`${overlays} Forge runtime control(s) were visible`);
} catch (error) {
  errors.push(error.stack || error.message);
}

if (externalRequests.length) errors.push(`External requests: ${externalRequests.join(', ')}`);
if (errors.length) {
  console.error(errors.map((error) => `- ${error}`).join('\n'));
  await browser.close();
  server.close();
  process.exit(1);
}

console.log('Forge srcdoc smoke test passed: clean data booted, glossary writes worked, drawer task layouts passed desktop/tablet/mobile checks in both themes, task edits worked, Not Required justification saved, and deletion completed without browser prompts.');
await browser.close();
server.close();
