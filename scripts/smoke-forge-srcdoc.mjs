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
  const loadingLogoStyle = await frame.locator('body').evaluate((body) => {
    const logo = document.createElement('img');
    logo.className = 'brand-logo loading-logo';
    body.appendChild(logo);
    const style = getComputedStyle(logo);
    const result = { animationName: style.animationName, boxShadow: style.boxShadow, filter: style.filter };
    logo.remove();
    return result;
  });
  if (loadingLogoStyle.animationName !== 'none' || loadingLogoStyle.boxShadow !== 'none' || loadingLogoStyle.filter !== 'none') errors.push(`Loading logo still glows or animates: ${JSON.stringify(loadingLogoStyle)}`);
  if (await frame.getByRole('button', { name: /New project/ }).count()) errors.push('Standard user could create projects');
  await frame.getByRole('button', { name: 'Users and managers', exact: true }).click();
  if (await frame.getByLabel('Testing password').count()) errors.push('Testing manager access was visible before Ctrl+M');
  await frame.locator('body').press('Control+m');
  await frame.getByLabel('Testing password').fill('incorrect-password');
  await frame.getByRole('button', { name: 'Enable manager access' }).click();
  await frame.getByRole('alert').filter({ hasText: 'Testing password is incorrect' }).waitFor();
  if (await frame.getByRole('button', { name: /New project/ }).count()) errors.push('Incorrect password granted manager access');
  await frame.getByLabel('Testing password').fill('admin123');
  await frame.getByRole('button', { name: 'Enable manager access' }).click();
  await frame.getByRole('button', { name: /New project/ }).waitFor();
  // Reload the packaged application to verify persisted role, not just UI state.
  await page.evaluate(() => {
    const key = 'modernization-project-tracker:v2';
    const data = JSON.parse(localStorage.getItem(key));
    data.users.push({ id: 'smoke-user', title: 'SME Test User', email: 'sme@example.test', loginName: 'sme@example.test', role: 'User' });
    localStorage.setItem(key, JSON.stringify(data));
  });
  await page.reload();
  await page.locator('#app').evaluate((element, html) => { element.srcdoc = html; }, artifact);
  await frame.getByRole('button', { name: /New project/ }).waitFor();
  await frame.getByRole('button', { name: 'Users and managers', exact: true }).click();
  // Reproduce a host that cancels native submit-button activation.
  await frame.locator('body').evaluate((body) => body.addEventListener('click', (event) => {
    if (event.target.closest('.directory-form button')) event.preventDefault();
  }, true));
  if (await frame.locator('.directory-form').count()) errors.push('User edit form displayed without Update User');
  const smeRow = frame.locator('.directory-row').filter({ hasText: 'SME Test User' });
  await smeRow.getByRole('button', { name: 'Edit', exact: true }).click();
  if (await frame.locator('.directory-form').count()) errors.push('Edit opened form instead of menu');
  await smeRow.getByRole('button', { name: 'Update User', exact: true }).click();
  await frame.getByLabel('Role', { exact: true }).selectOption('SME');
  await frame.getByRole('button', { name: 'Save user', exact: true }).click();
  await frame.getByRole('status').filter({ hasText: 'SME Test User saved as SME.' }).waitFor();
  await smeRow.getByText('SME', { exact: true }).waitFor();
  await frame.getByRole('button', { name: 'Portfolio', exact: true }).click();
  await frame.getByRole('heading', { name: 'Modernization at a glance' }).waitFor();
  await frame.getByText('0 total measurement areas').waitFor();
  const phaseLabels = await frame.locator('.phase-label').allTextContents();
  if (phaseLabels.join('|') !== 'Requirement (MSA)|Acquisition (TMRR)|Procurement (EMD)|Deployment (P&D)') errors.push(`Incorrect phase sequence: ${phaseLabels}`);
  const tracks = await frame.locator('.pipeline-rail').evaluate((row) => getComputedStyle(row).gridTemplateColumns.split(' ').length);
  if (tracks !== 4) errors.push(`Expected four pipeline columns, got ${tracks}`);
  const stageGeometry = await frame.locator('.phase-node').evaluateAll((nodes) => nodes.map((node) => {
    const label = node.querySelector('.phase-label').getBoundingClientRect();
    const line = node.querySelector('.phase-line').getBoundingClientRect();
    const count = node.querySelector('.phase-count').getBoundingClientRect();
    return label.bottom <= line.top && line.bottom <= count.top && Math.abs((label.left + label.right) / 2 - (count.left + count.right) / 2) < 2;
  }));
  if (stageGeometry.some((valid) => !valid)) errors.push('Pipeline labels, lines, and count circles are not centered in order');
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

  await frame.getByRole('button', { name: 'Portfolio', exact: true }).click();
  await frame.getByRole('button', { name: /New project/ }).click();
  await frame.locator('.modal .field input').first().fill('Smoke test modernization project');
  await frame.getByRole('button', { name: /Save project/ }).click();
  await frame.locator('.project-drawer').waitFor();
  await frame.getByLabel('Progress calculation', { exact: true }).selectOption('tasks');
  await frame.getByText('0 of 4 tasks completed', { exact: true }).waitFor();
  if (await frame.locator('.task-markers > span').count() !== 4) errors.push('Task progress must display one marker per task');
  if (await frame.getByRole('button', { name: 'Claim project', exact: true }).count()) errors.push('Claim project remained visible');
  await frame.getByRole('button', { name: 'Project settings', exact: true }).click();
  await frame.getByRole('button', { name: 'Edit project', exact: true }).waitFor();
  await frame.getByRole('button', { name: 'Project settings', exact: true }).click();
  await frame.getByRole('button', { name: 'Project settings', exact: true }).click();
  await frame.locator('.drawer-progress').click();
  if (await frame.locator('.drawer-header-actions .project-menu-popover').count()) errors.push('Settings menu did not dismiss on outside click');
  const markers = await frame.locator('.task-markers').evaluate((rail) => {
    const balls = [...rail.children].map((el) => el.getBoundingClientRect());
    return { spread: balls.at(-1).left - balls[0].left, width: rail.clientWidth, track: getComputedStyle(rail, '::before').display };
  });
  if (markers.spread < markers.width * .65 || markers.track === 'none') errors.push('Task markers are not distributed across a connected track');
  // Removing the task-code cell must not leave titles in its old 48px track.
  // Exercise realistic long titles in the packaged iframe, at multiple sizes.
  const upcomingTitle = frame.locator('.upcoming-list > button strong').first();
  const originalTitle = await upcomingTitle.textContent();
  await upcomingTitle.evaluate((element) => { element.textContent = 'Develop Capability Development Document (CDD/CSS) and incorporate comments from the technical review team'; });
  for (const width of [1440, 768, 390]) {
    await page.setViewportSize({ width, height: 1000 });
    for (const theme of ['light', 'dark']) {
      await frame.locator('html').evaluate((element, value) => { element.dataset.theme = value; }, theme);
      const tabs = await frame.locator('.drawer-tabs').evaluate((el) => ({ overflowY: getComputedStyle(el).overflowY, scrollbar: getComputedStyle(el).scrollbarWidth }));
      if (tabs.overflowY !== 'hidden' || tabs.scrollbar !== 'none') errors.push('Drawer tab strip has a visible scrollbar');
      const headerButtons = await frame.locator('.top-actions > button').evaluateAll((buttons) => buttons.map((el) => getComputedStyle(el).backgroundColor));
      if (headerButtons.some((color) => color !== 'rgba(0, 0, 0, 0)')) errors.push('Header buttons have a differing background');
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
  await frame.getByLabel('Est. Hours', { exact: true }).fill('3.5');
  await frame.getByLabel('Attach documents', { exact: true }).setInputFiles({ name: 'smoke-report.txt', mimeType: 'text/plain', buffer: Buffer.from('Task document smoke test') });
  await frame.getByRole('button', { name: 'smoke-report.txt', exact: true }).waitFor();
  for (const theme of ['light', 'dark']) {
    await frame.locator('html').evaluate((element, value) => { element.dataset.theme = value; }, theme);
    const uploadStyle = await frame.getByLabel('Attach documents', { exact: true }).evaluate((input) => {
      const button = getComputedStyle(input, '::file-selector-button');
      return { background: button.backgroundColor, color: button.color, radius: button.borderRadius };
    });
    if (uploadStyle.radius !== '7px' || uploadStyle.background === uploadStyle.color) errors.push(`Upload theme styling failed: ${theme}`);
  }


  await frame.locator('.modal select option', { hasText: 'Not Required' }).waitFor({ state: 'attached' });
  await frame.locator('.modal select').filter({ has: frame.locator('option', { hasText: 'Not Required' }) }).selectOption({ label: 'Not Required' });
  await frame.getByLabel('Not required justification').fill('Not applicable to this modernization project');
  await frame.getByRole('button', { name: 'Save task' }).click();
  await frame.getByRole('button', { name: quickCompleteLabel.replace('Complete', 'Reopen'), exact: true }).waitFor();
  await frame.getByRole('button', { name: quickCompleteLabel.replace('Complete', 'Reopen'), exact: true }).locator('xpath=..').click({ position: { x: 110, y: 12 } });
  await frame.getByRole('button', { name: 'smoke-report.txt', exact: true }).waitFor();
  if (await frame.getByLabel('Est. Hours', { exact: true }).inputValue() !== '3.5') errors.push('Estimated hours did not persist');
  await frame.locator('.modal').getByRole('button', { name: 'Cancel', exact: true }).click();
  await frame.getByRole('button', { name: 'Overview', exact: true }).click();
  const documents = frame.locator('.project-documents');
  await documents.getByRole('button', { name: 'smoke-report.txt', exact: true }).waitFor();
  const downloadEvent = page.waitForEvent('download');
  await documents.getByRole('button', { name: 'Download smoke-report.txt', exact: true }).click();
  const download = await downloadEvent;
  if (download.suggestedFilename() !== 'smoke-report.txt' || await fs.readFile(await download.path(), 'utf8') !== 'Task document smoke test') errors.push('Document download failed');
  await documents.getByRole('button', { name: 'Delete smoke-report.txt', exact: true }).click();
  await documents.getByText('No documents attached.', { exact: true }).waitFor();
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

  // Shared references must persist folders and bytes inside the same srcdoc host.
  await frame.getByRole('button', { name: 'Reference Documents', exact: true }).click();
  await frame.getByRole('button', { name: 'New folder', exact: true }).click();
  if (await frame.getByLabel('Folder location', { exact: true }).count()) errors.push('New folder unnecessarily asks for a location');
  await frame.getByLabel('Folder name', { exact: true }).fill('Templates');
  await frame.getByRole('button', { name: 'Create folder', exact: true }).click();
  await frame.getByRole('button', { name: 'Templates', exact: true }).click();
  await frame.getByLabel('Upload reference documents').setInputFiles({ name: 'sample-template.txt', mimeType: 'text/plain', buffer: Buffer.from('Modernization template sample') });
  await frame.getByRole('status').filter({ hasText: '1 of 1 documents uploaded.' }).waitFor();
  const downloadPromise = page.waitForEvent('download');
  await frame.getByRole('button', { name: 'Edit sample-template.txt', exact: true }).click();
  if (await frame.locator('.reference-action-menu svg').count()) errors.push('Reference action menu includes icons');
  await frame.getByRole('button', { name: 'Download', exact: true }).click();
  const downloaded = await downloadPromise;
  if ((await fs.readFile(await downloaded.path(), 'utf8')) !== 'Modernization template sample') errors.push('Reference download content did not match upload');
  await frame.getByRole('button', { name: 'Edit sample-template.txt', exact: true }).click();
  await frame.getByRole('button', { name: 'Move', exact: true }).click();
  await frame.getByLabel('Folder location', { exact: true }).selectOption('');
  await frame.getByRole('button', { name: 'Save reference', exact: true }).click();
  await frame.getByRole('button', { name: 'All references', exact: true }).click();
  await frame.getByRole('button', { name: 'sample-template.txt', exact: true }).waitFor();
  for (const theme of ['light', 'dark']) {
    await frame.locator('html').evaluate((element, value) => { element.dataset.theme = value; }, theme);
    for (const width of [1440, 390]) {
      await page.setViewportSize({ width, height: 1000 });
      const style = await frame.getByLabel('Upload reference documents').evaluate((input) => {
        const selector = getComputedStyle(input, '::file-selector-button');
        return { fg: selector.color, bg: selector.backgroundColor };
      });
      if (style.fg === style.bg || style.bg === 'rgba(0, 0, 0, 0)') errors.push('Reference upload lacks themed file control');
    }
  }
  await page.setViewportSize({ width: 1440, height: 1000 });
  await page.reload();
  await page.locator('#app').evaluate((element, html) => { element.srcdoc = html; }, artifact);
  await frame.getByRole('button', { name: 'Reference Documents', exact: true }).click();
  await frame.getByRole('button', { name: 'sample-template.txt', exact: true }).waitFor();
  await frame.getByRole('button', { name: 'Edit sample-template.txt', exact: true }).click();
  await frame.getByRole('button', { name: 'Delete', exact: true }).click();
  await frame.getByRole('status').filter({ hasText: 'sample-template.txt deleted.' }).waitFor();
  if (await frame.getByRole('button', { name: 'sample-template.txt', exact: true }).count()) errors.push('Deleted reference still visible');
  await page.reload();
  await page.locator('#app').evaluate((element, html) => { element.srcdoc = html; }, artifact);
  await frame.getByRole('button', { name: 'Reference Documents', exact: true }).click();
  await frame.getByRole('button', { name: 'Templates', exact: true }).waitFor();
  if (await frame.getByRole('button', { name: 'sample-template.txt', exact: true }).count()) errors.push('Deleted reference returned after reload');
  await frame.getByRole('button', { name: 'Users and managers', exact: true }).click();
  await frame.getByLabel('Search users', { exact: true }).fill('SME Test');
  if (await frame.locator('.directory-row').count() !== 1) errors.push('Directory search did not filter users');
  await frame.locator('.directory-row').getByRole('button', { name: 'Edit', exact: true }).click();
  if (await frame.getByRole('button', { name: 'Delete User', exact: true }).locator('svg').count()) errors.push('Delete User still has an icon');
  if (await frame.locator('.people-invite input[type="url"], .people-invite a[href^="mailto:"]').count()) errors.push('Invitation link/draft controls were not removed');
  if (dialogs.length) errors.push(`Native browser dialog(s) displayed: ${dialogs.join(' | ')}`);
  const overlays = await frame.locator('#pdc-open, #test-recorder-launcher, #test-recorder-panel, .test-recorder-ui').evaluateAll((elements) => elements.filter((element) => getComputedStyle(element).display !== 'none').length);
  if (overlays) errors.push(`${overlays} Forge runtime control(s) were visible`);
} catch (error) {
  errors.push(error.stack || error.message);
  const referenceState = await page.frameLocator('#app').locator('.reference-library').textContent().catch(() => 'Reference view unavailable');
  errors.push(`Reference view at failure: ${referenceState}`);
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
