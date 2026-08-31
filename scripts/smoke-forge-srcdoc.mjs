import fs from 'node:fs/promises';
import path from 'node:path';
import { chromium } from 'playwright';

const artifactPath = path.resolve('build-singlefile/modernization-project-tracker.html');
const artifact = await fs.readFile(artifactPath, 'utf8');
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const errors = [];
const externalRequests = [];

page.on('console', (message) => { if (message.type() === 'error') errors.push(message.text()); });
page.on('pageerror', (error) => errors.push(error.message));
page.on('request', (request) => {
  const url = request.url();
  if (/^https?:/i.test(url)) externalRequests.push(url);
});

await page.setContent('<!doctype html><html><body style="margin:0"><iframe id="app" style="width:100vw;height:100vh;border:0"></iframe></body></html>');
await page.locator('#app').evaluate((frame, html) => { frame.srcdoc = html; }, artifact);
const frame = page.frameLocator('#app');

await frame.getByRole('heading', { name: 'Modernization at a glance' }).waitFor({ timeout: 20_000 });
await frame.getByText('14 total measurement areas').waitFor();
await frame.getByRole('button', { name: /Pipeline board/ }).click();
await frame.getByRole('heading', { name: 'Pipeline board' }).waitFor();
await frame.locator('.project-card').first().click();
await frame.locator('.project-drawer').waitFor();
await frame.getByRole('button', { name: /Work breakdown/ }).click();
await frame.locator('.wbs-task').first().waitFor();
await frame.locator('.project-drawer .icon-button').first().click();

if (externalRequests.length) errors.push(`External requests: ${externalRequests.join(', ')}`);
if (errors.length) {
  console.error(errors.map((error) => `- ${error}`).join('\n'));
  await browser.close();
  process.exit(1);
}

console.log('Forge srcdoc smoke test passed: app booted, portfolio rendered, board opened, and project WBS loaded.');
await browser.close();
