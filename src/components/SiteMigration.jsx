import React, { useEffect, useMemo, useRef, useState } from 'react';
import { isMigrationTarget, migrationSite, MIGRATION_SOURCE } from '../lib/siteMigration';
import { MAX_BACKUP_BYTES, exportTrackerBackup, validateTrackerBackup, previewTrackerImport, importTrackerBackup, manualTransferItems } from '../lib/trackerBackup';

function download(text, name) {
  const url = URL.createObjectURL(new Blob([text], { type: 'application/json' }));
  const link = document.createElement('a'); link.href = url; link.download = name; link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
export function SiteMigration({ store }) {
  const source = String(store?.webUrl || '').replace(/\/+$/, '').toLowerCase() === MIGRATION_SOURCE.toLowerCase();
  const target = isMigrationTarget(store?.webUrl);
  const [allowed, setAllowed] = useState(false), [busy, setBusy] = useState(false);
  const [backup, setBackup] = useState(null), [plan, setPlan] = useState(null), [report, setReport] = useState(null);
  const [paused, setPaused] = useState(false), [replace, setReplace] = useState(false);
  const [includeTaskAttachments, setIncludeTaskAttachments] = useState(true);
  const [status, setStatus] = useState(''), [error, setError] = useState('');
  const [activity, setActivity] = useState({ text: '', started: 0 }), [clock, setClock] = useState(Date.now());
  const active = useRef(false), fileInput = useRef(null);
  const site = useMemo(() => source || target ? migrationSite(store.webUrl, store.prefix, undefined, { onActivity: text => { setActivity({ text, started: Date.now() }); setClock(Date.now()); } }) : null, [source, target, store?.webUrl, store?.prefix]);
  useEffect(() => {
    let mounted = true; setAllowed(false);
    site?.canMigrate().then(value => { if (mounted) setAllowed(value); }).catch(() => {});
    return () => { mounted = false; };
  }, [site]);
  useEffect(() => {
    if (!busy) return;
    const warn = event => { event.preventDefault(); event.returnValue = ''; };
    const timer = setInterval(() => setClock(Date.now()), 1000);
    window.addEventListener('beforeunload', warn);
    return () => { clearInterval(timer); window.removeEventListener('beforeunload', warn); };
  }, [busy]);
  if (!site || !allowed) return null;
  const conflicts = plan?.lists.reduce((sum, list) => sum + list.entries.filter(entry => entry.conflict).length, 0) || 0;
  async function perform(operation, writes = false) {
    if (active.current) return;
    active.current = true; setBusy(true); setError(''); setReport(null);
    let last = 'Preparing';
    const progress = text => { last = text; setStatus(text); };
    try { await operation(progress); }
    catch (err) {
      setError(`${err.message}${writes ? ' Some records may already be imported. Keep them and preview the same package again to resume.' : ''}`);
      setStatus(`Stopped at: ${last}`); setPlan(null);
    } finally { active.current = false; setBusy(false); }
  }
  async function choose(file) {
    if (!file) return;
    setPlan(null); setBackup(null); setReplace(false);
    await perform(async progress => {
      progress('Checking backup package…');
      if (file.size > MAX_BACKUP_BYTES) throw new Error('Backup files must be no larger than 256 MB.');
      const value = JSON.parse(await file.text());
      await validateTrackerBackup(value); setBackup(value);
      setStatus(`Backup ready: ${value.payload.lists.reduce((sum, list) => sum + list.records.length, 0)} records and ${value.payload.files.length} documents. Preview to compare with metsoft.`);
    });
  }
  return <details className="panel site-migration">
    <summary>Site owner tools · {source ? 'Export tracker backup' : 'Import tracker backup'}</summary>
    <div className="migration-content">
      <p>{source ? 'Download the tracker records and attached documents in one backup package. Export reads this site only.' : 'Import a backup downloaded from the old tracker. Import uses the package and metsoft only; it does not connect to ISEA METENG.'}</p>
      <p>{source ? 'Exports Projects, Tasks, Updates, Risks, Acronyms, and Users, including archived records, app roles, and assignments. Reference Documents and their folders are excluded; transfer them manually.' : 'Only the lists and files included in the package are imported. Existing Reference Documents stay in place when excluded from the package.'} Site permissions and version history are separate.</p>
      <label className="migration-check"><input type="checkbox" checked={paused} disabled={busy} onChange={event => setPaused(event.target.checked)} /><span>{source ? 'Edits in the old tracker are paused for this export.' : 'Edits in metsoft are paused for this import.'}</span></label>
      {source ? <>
        <label className="migration-check"><input type="checkbox" checked={includeTaskAttachments} disabled={busy} onChange={event => setIncludeTaskAttachments(event.target.checked)} /><span>Include files attached to tasks</span></label>
        <p><small>Task attachments are separate from Reference Documents. If a task file causes a timeout, uncheck this option to export records without downloading any files, then transfer task attachments manually too.</small></p>
        <button type="button" className="button primary" disabled={busy || !paused} onClick={() => perform(async progress => {
          const text = await exportTrackerBackup(site, progress, { includeReferences: false, includeTaskAttachments });
          download(text, `npsl-tracker-without-references-${new Date().toISOString().slice(0, 10)}.npsl-backup.json`);
          setStatus(`Backup downloaded. Transfer Reference Documents and folders${includeTaskAttachments ? '' : ', plus task attachments,'} manually. Open the new tracker and choose Import tracker backup.`);
        })}>Download backup</button>
      </> : <>
        <button type="button" className="button secondary" disabled={busy} onClick={() => fileInput.current?.click()}>Choose backup file</button>
        <input ref={fileInput} aria-label="Choose tracker backup" type="file" accept=".json,application/json" disabled={busy} hidden onChange={event => { const file = event.target.files?.[0]; event.target.value = ''; choose(file); }} />
        {backup && manualTransferItems(backup).length > 0 && <p role="note">Manual transfer required: {manualTransferItems(backup).join('; ')}. These are excluded from this package and its verification.</p>}
        {backup && <button type="button" className="button secondary" disabled={busy} onClick={() => perform(async progress => { setPlan(null); setReplace(false); setPlan(await previewTrackerImport(backup, site, progress)); setStatus('Preview ready. No records have been imported.'); })}>Preview import</button>}
        {plan && <>
          <div className="migration-table"><table><caption>Import preview</caption><thead><tr><th>List</th><th>New</th><th>Already matching</th><th>Conflicts</th><th>Documents</th></tr></thead><tbody>{plan.lists.map(list => <tr key={list.key}><th>{list.title}</th><td>{list.entries.filter(entry => !entry.destination).length}</td><td>{list.entries.filter(entry => entry.destination && !entry.conflict).length}</td><td>{list.entries.filter(entry => entry.conflict).length}</td><td>{list.entries.reduce((sum, entry) => sum + entry.names.length, 0)}</td></tr>)}</tbody></table></div>
          {!!conflicts && <><details><summary>Review {conflicts} conflicting records</summary><ul>{plan.lists.flatMap(list => list.entries.filter(entry => entry.conflict).map(entry => <li key={`${list.key}/${entry.source.RecordId}`}>{list.title}: {entry.source.Title || entry.source.RecordId}{list.key === 'users' ? ` — role: ${entry.destination.AppRole || 'unset'} → ${entry.source.AppRole || 'unset'}` : ''}</li>))}</ul></details><label className="migration-check"><input type="checkbox" checked={replace} disabled={busy} onChange={event => setReplace(event.target.checked)} /><span>Replace these {conflicts} conflicting records with backup values, including the app roles shown above.</span></label></>}
          <p>Already matching records are reused. Destination-only records/files stay in place. Different existing documents are never overwritten.</p>
          <button type="button" className="button primary" disabled={busy || !paused || (!!conflicts && !replace)} onClick={() => perform(async progress => { const result = await importTrackerBackup(backup, site, plan, { replaceConflicts: replace, onProgress: progress }); setReport(result); setPlan(null); setStatus(`Included data verified. ${result.manualTransfer?.length ? `Manual transfer still required: ${result.manualTransfer.join('; ')}. ` : ''}Download the report, then reload the tracker.`); }, true)}>Import and verify</button>
        </>}
      </>}
      {status && <p role="status" aria-live="polite">{status}</p>}
      {busy && activity.text && <p style={{ overflowWrap: 'anywhere' }}><small>{activity.text}<br />Current request: {Math.max(0, Math.floor((clock - activity.started) / 1000))} seconds · 45-second deadline per attempt.</small></p>}
      {error && <p role="alert" className="inline-error">{error}</p>}
      {report && <div className="migration-actions"><button type="button" className="button secondary" onClick={() => download(JSON.stringify(report, null, 2), 'tracker-import-verification.json')}>Download verification report</button><button type="button" className="button primary" onClick={() => window.location.reload()}>Reload tracker</button></div>}
    </div>
  </details>;
}
