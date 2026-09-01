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
const server = http.createServer((_request, response) => {
  response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  response.end('<!doctype html><html><body style="margin:0"><iframe id="app" style="width:100vw;height:100vh;border:0"></iframe></body></html>');
});
await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const { port } = server.address();
const harnessOrigin = `http://127.0.0.1:${port}`;

page.on('console', (message) => { if (message.type() === 'error') errors.push(message.text()); });
page.on('pageerror', (error) => errors.push(error.message));
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

  await frame.getByRole('heading', { name: 'Modernization at a glance' }).waitFor({ timeout: 20_000 });
  await frame.getByText('0 total measurement areas').waitFor();
  await frame.getByText('Your modernization portfolio is ready').waitFor();
  await frame.getByRole('button', { name: /Open sample portfolio/ }).first().click();
  await frame.getByText('14 total measurement areas').waitFor();
  await frame.getByRole('button', { name: /Pipeline board/ }).click();
  await frame.getByRole('heading', { name: 'Pipeline board' }).waitFor();
  await frame.locator('.project-card').first().click();
  await frame.locator('.project-drawer').waitFor();
  await frame.getByRole('button', { name: /Work breakdown/ }).click();
  await frame.locator('.wbs-task').first().waitFor();
  const quickComplete = frame.locator('.task-check:not(.checked)').first();
  const quickCompleteLabel = await quickComplete.getAttribute('aria-label');
  await quickComplete.click();
  const reopenedTask = frame.getByRole('button', { name: quickCompleteLabel.replace('Complete WBS', 'Reopen WBS'), exact: true });
  await reopenedTask.waitFor();
  await reopenedTask.locator('xpath=..').click({ position: { x: 110, y: 12 } });
  await frame.getByRole('heading', { name: 'Update task' }).waitFor();
  await frame.getByRole('option', { name: 'Not Required' }).waitFor();
  await frame.locator('.modal button[aria-label="Close"]').click();
  await frame.locator('.project-drawer .icon-button').first().click();
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

console.log('Forge srcdoc smoke test passed: app booted, sample portfolio rendered, board opened, and task quick/edit interactions loaded.');
await browser.close();
server.close();
