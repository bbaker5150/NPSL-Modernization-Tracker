import './style.css';
import logo from '../../src/assets/navair-seal-384.webp';
import { scanReferences, readSource, acceptFile, makeZip } from './exporter.js';
import { uploadGuide } from './guide.js';

const $ = id => document.getElementById(id);
$('logo').src = logo;
let plan, busy = false, controller, packing = false, output;
const downloads = new Map(), labels = new Map();
const size = bytes => bytes == null ? 'Size unavailable' : `${(bytes / 1024 / 1024).toFixed(2)} MB`;
function message(text) { $('status').textContent = text; }
function error(caught) { $('error').hidden = !caught; $('error').textContent = caught?.message || ''; }
function setBusy(value) {
  busy = value;
  $('scan').disabled = value; $('export').disabled = value || !plan;
  $('stop').disabled = !value || packing;
  for (const input of document.querySelectorAll('input[type=file]')) input.disabled = value;
  $('export').textContent = output ? 'Download ZIP again' : downloads.size ? 'Retry / Download ZIP' : 'Download ZIP';
}
function save(blob, name) {
  const url = URL.createObjectURL(blob), a = document.createElement('a');
  a.href = url; a.download = name; document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
function text(tag, value, className) { const node = document.createElement(tag); node.textContent = value; if (className) node.className = className; return node; }
function render() {
  $('inventory').hidden = false; $('count').textContent = `${plan.files.length} attachments · ${plan.folders.length} folders`;
  $('warnings').replaceChildren(...plan.warnings.map(warning => text('p', warning, 'warning')));
  $('files').replaceChildren(); labels.clear();
  for (const file of plan.files) {
    const row = document.createElement('article'); row.className = 'file';
    row.append(text('h3', file.displayName || file.sourceName), text('p', `${file.zipPath} · ${size(file.expectedSize)}${file.archived ? ' · Archived' : ''}`, 'muted'));
    const label = text('p', downloads.has(file.key) ? 'Ready for ZIP' : 'Waiting', 'file-status'); labels.set(file.key, label); row.append(label);
    const actions = document.createElement('div'); actions.className = 'actions';
    const link = text('a', 'Download original'); link.href = file.url; link.target = '_blank'; link.rel = 'noopener'; link.download = file.sourceName; actions.append(link);
    const picker = text('label', 'Choose downloaded file'); picker.className = 'file-picker';
    const input = document.createElement('input'); input.type = 'file'; input.setAttribute('aria-label', `Choose downloaded file for ${file.sourceName}`);
    input.addEventListener('change', async () => {
      const selected = input.files[0]; if (!selected || busy) return;
      setBusy(true); error(null);
      try {
        downloads.set(file.key, { ...await acceptFile(file, await selected.arrayBuffer(), { name: selected.name }), retrievedBy: 'Manually supplied original (name and size checked)' });
        output = null; label.textContent = 'Ready for ZIP · manually supplied'; message(`${downloads.size} of ${plan.files.length} files ready.`);
      } catch (caught) { error(caught); }
      finally { input.value = ''; setBusy(false); }
    });
    picker.append(input); actions.append(picker); row.append(actions); $('files').append(row);
  }
}
$('scan').addEventListener('click', async () => {
  if (busy) return;
  setBusy(true); error(null); message('Reading reference folders and attachments…');
  controller = new AbortController();
  try {
    const scanned = await scanReferences({ signal: controller.signal, timeoutMs: 30000 });
    plan = scanned; downloads.clear(); output = null; render();
    message(`${plan.files.length} attachments found. Download ZIP to collect every file.`);
  } catch (caught) { error(caught); message('Scan stopped. No SharePoint data was changed.'); }
  finally { setBusy(false); }
});
const signature = value => JSON.stringify({ files: value.files, folders: value.folders, modified: value.snapshot.map(row => [row.id, row.modified]) });
$('export').addEventListener('click', async () => {
  if (busy || !plan) return;
  if (output) { save(output, 'Modernization-Tracker-Reference-Documents.zip'); return; }
  setBusy(true); error(null); controller = new AbortController(); $('progress').hidden = false;
  try {
    for (const [index, file] of plan.files.entries()) {
      if (controller.signal.aborted) throw new Error('Stopped. Retry ZIP to continue using the files already downloaded.');
      if (downloads.has(file.key)) continue;
      const started = Date.now();
      const update = () => message(`Downloading ${index + 1} of ${plan.files.length}: ${file.sourceName} · ${Math.floor((Date.now() - started) / 1000)}s (120s limit)`);
      update(); const timer = setInterval(update, 1000);
      labels.get(file.key).textContent = 'Downloading…';
      try {
        const bytes = await readSource(file.url, { binary: true, signal: controller.signal });
        downloads.set(file.key, await acceptFile(file, bytes));
        labels.get(file.key).textContent = `Ready for ZIP · ${size(bytes.byteLength)}`;
      } catch (caught) { labels.get(file.key).textContent = 'Needs retry or manual download'; throw caught; }
      finally { clearInterval(timer); }
      $('progress').value = downloads.size / Math.max(1, plan.files.length) * 80;
    }
    message('Checking that the source inventory has not changed…');
    const current = await scanReferences({ signal: controller.signal, timeoutMs: 30000 });
    if (signature(current) !== signature(plan)) throw new Error('Reference documents changed after scanning. Pause edits and scan again to start a consistent export.');
    packing = true; setBusy(true); message('Creating ZIP and upload guide…');
    const result = await makeZip(plan, downloads, percent => { $('progress').value = 80 + percent * 0.2; });
    output = result.blob; plan = { ...plan, ...result.manifest, snapshot: plan.snapshot };
    save(output, 'Modernization-Tracker-Reference-Documents.zip');
    message(`Complete: ${plan.files.length} attachments packaged. Extract the ZIP and open its upload guide.`);
  } catch (caught) { error(caught); message(`Export stopped. ${downloads.size} of ${plan.files.length} attachments retained on this page. No incomplete ZIP was downloaded.`); }
  finally { packing = false; setBusy(false); $('progress').hidden = true; }
});
$('stop').addEventListener('click', () => controller?.abort());
$('guide').addEventListener('click', () => save(new Blob([uploadGuide(plan)], { type: 'text/html;charset=utf-8' }), 'Reference-Document-Upload-Guide.html'));
window.addEventListener('beforeunload', event => { if (busy || (downloads.size && !output)) { event.preventDefault(); event.returnValue = ''; } });
// No parent-window confirmation interception or replacement fetch channel.
// All source access uses the host-provided fetch implementation and its policy.
