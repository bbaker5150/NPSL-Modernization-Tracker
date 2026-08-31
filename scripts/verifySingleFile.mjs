import fs from 'node:fs/promises';
import path from 'node:path';
import { applyForgeRuntime, EXPECTED_HASHES, loadForgeRuntime } from './forgeRuntime.mjs';

const filePath = path.resolve('build-singlefile/modernization-project-tracker.html');
const html = await fs.readFile(filePath, 'utf8');
const failures = [];
const expect = (condition, message) => { if (!condition) failures.push(message); };

expect(html.startsWith('<!--WFC-MANIFEST:'), 'Forge manifest is missing.');
expect(html.includes('x-modernization-tracker-build'), 'Build stamp is missing.');
expect(html.includes('window.__MOD_TRACKER_BUILD__'), 'Runtime build marker is missing.');
expect(html.includes('window.MOD_TRACKER_CONFIG'), 'Deployment configuration hook is missing.');
expect(!/<script[^>]+src=/i.test(html), 'External script reference found.');
expect(!/<link[^>]+(?:href|rel=["\']stylesheet)/i.test(html), 'External stylesheet or link reference found.');
expect(!/<img[^>]+src=["\'](?!data:)/i.test(html), 'Non-inlined image reference found.');
expect(loadForgeRuntime().every((file) => file.hash === EXPECTED_HASHES[file.name]), 'Vendored Forge runtime hash mismatch.');

// Exercise the injector independently so the build cannot pass only because a
// stale output happened to be present.
const fixture = '<!doctype html><html><head></head><body><div id="root"></div></body></html>';
const injected = applyForgeRuntime(fixture, { index: 'fixture.html', project: 'Fixture', generated: new Date(0).toISOString(), build: 'test' });
expect(injected.includes('WFC-MANIFEST'), 'Forge injector fixture failed.');

if (failures.length) {
  console.error(failures.map((failure) => `- ${failure}`).join('\n'));
  process.exit(1);
}
console.log(`Verified ${path.basename(filePath)} (${(html.length / 1024).toFixed(1)} KiB): manifest, runtime, build stamp, and zero external subresources.`);
