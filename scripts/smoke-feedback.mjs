import fs from 'node:fs/promises';
import http from 'node:http';
import assert from 'node:assert/strict';
import { chromium } from 'playwright';
const html = await fs.readFile('build-singlefile/modernization-project-tracker.html', 'utf8');
const server = http.createServer((req, res) => { res.setHeader('Content-Type', 'text/html'); res.end('<iframe id="app" style="position:fixed;inset:0;width:100%;height:100%;border:0"></iframe>'); });
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const errors = []; page.on('pageerror', error => errors.push(error.message));
try {
  await page.goto(`http://127.0.0.1:${server.address().port}`);
  await page.locator('#app').evaluate((frame, html) => { frame.srcdoc = html; }, html);
  const app = page.frameLocator('#app');
  const trigger = app.getByRole('button', { name: 'Report a bug or request a feature' });
  await trigger.click();
  assert.equal(await trigger.evaluate(el => el.nextElementSibling?.tagName === 'DIALOG' ? el.nextElementSibling.nextElementSibling?.getAttribute('aria-label') : el.nextElementSibling?.getAttribute('aria-label')), 'Switch to dark mode');
  assert.equal(await app.getByRole('button', { name: 'Send feedback', exact: true }).count(), 1);
  await app.getByLabel('Summary', { exact: true }).fill('The project filter needs attention');
  await app.getByLabel('What happened?', { exact: true }).fill('Steps to reproduce\nSelect a project and change the filter.');
  await fs.mkdir('test-artifacts', { recursive: true });
  await app.getByRole('dialog').evaluate(el => Promise.all(el.getAnimations().map(a => a.finished)));
  await page.screenshot({ path: 'test-artifacts/feedback-light.png' });
  // SharePoint/Forge hosts can cancel native form-button activation.
  // A direct click handler must still send feedback exactly once.
  await app.locator('body').evaluate(body => body.ownerDocument.addEventListener('click', event => {
    if (event.target.closest('.feedback-actions button')) event.preventDefault();
  }, true));
  await app.getByRole('button', { name: 'Send feedback', exact: true }).click();
  await app.getByRole('heading', { name: 'Thank you for the feedback' }).waitFor();
  await app.getByRole('button', { name: 'Owner inbox' }).click();
  await app.getByRole('heading', { name: 'The project filter needs attention' }).waitFor();
  await page.keyboard.press('Escape');
  assert.equal(await trigger.evaluate(el => document.activeElement === el), true);
  await app.getByRole('button', { name: 'Switch to dark mode' }).click();
  await page.setViewportSize({ width: 390, height: 844 });
  await trigger.click();
  await app.getByRole('radio', { name: /Request a feature/ }).check();
  const box = await app.getByRole('dialog').boundingBox();
  assert.ok(box.x >= 0 && box.x + box.width <= 391);
  await app.getByRole('dialog').evaluate(el => Promise.all(el.getAnimations().map(a => a.finished)));
  assert.equal(await app.getByRole('button', { name: 'Send feedback', exact: true }).last().isVisible(), true);
  await page.screenshot({ path: 'test-artifacts/feedback-dark-mobile.png' });
  const inbox = await app.getByRole('button', { name: 'Owner inbox' }).boundingBox();
  const send = await app.getByRole('button', { name: 'Send feedback', exact: true }).boundingBox();
  assert.ok(inbox.x + inbox.width <= send.x && Math.abs(inbox.y - send.y) < 2, 'Inbox must sit to the left of Send feedback on mobile');
  await app.getByRole('button', { name: 'Owner inbox' }).click();
  await app.getByRole('button', { name: 'Back to feedback' }).click();
  assert.equal(await app.getByRole('button', { name: 'Send feedback', exact: true }).count(), 1);
  assert.deepEqual(errors, []);
  console.log('Feedback smoke passed: placement, submit, owner inbox, focus return, dark mobile layout.');
} finally { await browser.close(); await new Promise(resolve => server.close(resolve)); }
