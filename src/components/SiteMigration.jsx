import React, { useEffect, useMemo, useState } from 'react';
import { isMigrationTarget, migrationSite, MIGRATION_SOURCE, MIGRATION_TARGET, previewMigration, runMigration } from '../lib/siteMigration';

export function SiteMigration({ store }) {
  const enabled = isMigrationTarget(store?.webUrl);
  const sites = useMemo(() => enabled ? {
    source: migrationSite(MIGRATION_SOURCE, store.prefix),
    target: migrationSite(MIGRATION_TARGET, store.prefix),
  } : null, [enabled, store?.prefix]);
  const [allowed, setAllowed] = useState(false);
  const [busy, setBusy] = useState(false);
  const [plan, setPlan] = useState(null);
  const [paused, setPaused] = useState(false);
  const [replace, setReplace] = useState(false);
  const [status, setStatus] = useState('');
  const [error, setError] = useState('');
  const [report, setReport] = useState(null);
  useEffect(() => {
    let active = true;
    setAllowed(false);
    sites?.target.canMigrate().then(value => { if (active) setAllowed(value); }).catch(() => {});
    return () => { active = false; };
  }, [sites]);
  useEffect(() => {
    if (!busy) return;
    const warn = event => { event.preventDefault(); event.returnValue = ''; };
    window.addEventListener('beforeunload', warn);
    return () => window.removeEventListener('beforeunload', warn);
  }, [busy]);
  if (!enabled || !allowed) return null;
  const conflicts = plan?.lists.reduce((sum, list) => sum + list.entries.filter(entry => entry.conflict).length, 0) || 0;
  async function preview() {
    let lastStep = 'Reading migration preview';
    setBusy(true); setError(''); setPlan(null); setReport(null); setPaused(false); setReplace(false);
    try {
      setPlan(await previewMigration(sites.source, sites.target, message => { lastStep = message; setStatus(message); }));
      setStatus('Preview ready. No records have been copied.');
    } catch (err) { setError(err.message); setStatus(`Preview stopped at: ${lastStep}`); }
    finally { setBusy(false); }
  }
  async function copy() {
    if (!plan || busy || !paused || (conflicts && !replace)) return;
    setBusy(true); setError('');
    let lastStep = 'Checking source and destination';
    try {
      setReport(await runMigration(sites.source, sites.target, plan, { replaceConflicts: replace, onProgress: message => { lastStep = message; setStatus(message); } }));
      setStatus('Copy verified. Download the report, then reload the tracker.');
    } catch (err) {
      setError(`${err.message} The source is unchanged. Some destination records may already be copied; preview again to resume safely.`);
      setStatus(`Migration stopped at: ${lastStep}`);
    } finally { setPlan(null); setBusy(false); }
  }
  function downloadReport() {
    const url = URL.createObjectURL(new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' }));
    const link = document.createElement('a'); link.href = url; link.download = 'modernization-migration-verification.json'; link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  return <details className="panel site-migration">
    <summary>Site owner tools · Migrate from ISEA METENG</summary>
    <div className="migration-content">
      <p>Copy tracker records and attached documents from ISEA METENG to metsoft. Source data is never changed.</p>
      <dl><dt>From</dt><dd>{MIGRATION_SOURCE}</dd><dt>To</dt><dd>{MIGRATION_TARGET}</dd></dl>
      <p>Includes archived records, project relationships, reference folders, and app roles. SharePoint membership, permissions, version history, system timestamps, and files outside the tracker lists are not copied.</p>
      <button type="button" className="button secondary" disabled={busy} onClick={preview}>Preview migration</button>
      {plan && <>
        <div className="migration-table"><table><caption>Records to copy</caption><thead><tr><th scope="col">List</th><th scope="col">New</th><th scope="col">Matching</th><th scope="col">Conflicts</th><th scope="col">Files</th></tr></thead><tbody>{plan.lists.map(list => <tr key={list.key}><th scope="row">{list.title}</th><td>{list.entries.filter(entry => !entry.destination).length}</td><td>{list.entries.filter(entry => entry.destination && !entry.conflict).length}</td><td>{list.entries.filter(entry => entry.conflict).length}</td><td>{list.entries.reduce((sum, entry) => sum + entry.names.length, 0)}</td></tr>)}</tbody></table></div>
        {!!conflicts && <details><summary>Review {conflicts} conflicting records</summary><ul>{plan.lists.flatMap(list => list.entries.filter(entry => entry.conflict).map(entry => <li key={`${list.key}/${entry.source.RecordId}`}>{list.title}: {entry.source.Title || entry.source.RecordId} — differs in {list.fields.filter(field => JSON.stringify(entry.source[field] ?? '') !== JSON.stringify(entry.destination[field] ?? '')).join(', ')}</li>))}</ul></details>}
        <p>Destination-only records and files are retained. Different files with the same name stop the migration without overwriting. Keep this page open until verification finishes.</p>
        <label className="migration-check"><input type="checkbox" checked={paused} disabled={busy} onChange={event => setPaused(event.target.checked)} /><span>I have paused edits in both trackers until verification finishes.</span></label>
        {!!conflicts && <label className="migration-check"><input type="checkbox" checked={replace} disabled={busy} onChange={event => setReplace(event.target.checked)} /><span>Replace the {conflicts} conflicting destination records with source values, including any app roles shown above.</span></label>}
        <button type="button" className="button primary" disabled={busy || !paused || (!!conflicts && !replace)} onClick={copy}>Copy and verify</button>
      </>}
      {status && <p role="status" aria-live="polite">{status}</p>}
      {error && <p role="alert" className="inline-error">{error}</p>}
      {report && <div className="migration-actions"><button type="button" className="button secondary" onClick={downloadReport}>Download verification report</button><button type="button" className="button primary" onClick={() => window.location.reload()}>Reload tracker</button></div>}
    </div>
  </details>;
}
