import React, { useState } from 'react';

export function TrackerAccessSetup({ store }) {
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState('');
  const [error, setError] = useState('');
  if (!store.scopedAccess) return null;
  async function apply() {
    if (busy) return;
    setBusy(true); setError(''); setStatus('Checking tracker access…');
    try {
      const result = await store.prepareTrackerAccess(setStatus);
      setStatus(`Verified tracker groups for ${result.users} users and engineer permissions for ${result.projects} projects. Reload the tracker before testing with another account.`);
    } catch (caught) {
      setError(`Setup stopped. Completed changes remain in place; run this again to finish. ${caught.message}`);
    } finally { setBusy(false); }
  }
  return <details className="panel">
    <summary>Tracker permissions</summary>
    <p>Apply saved roles to Tracker Viewers, Tracker Project Engineers, and Tracker Managers. Apply engineer edit access to assigned projects and their tasks, updates, risks, and task documents.</p>
    <p>Run once as a site owner after deployment or importing records. Subsequent role and engineer changes update access automatically. This keeps existing site groups and their permissions in place.</p>
    <button type="button" className="button primary" disabled={busy} onClick={apply}>{busy ? 'Applying permissions…' : 'Apply tracker permissions'}</button>
    {status && <p role="status" aria-live="polite">{status}</p>}
    {error && <p role="alert" className="inline-error">{error}</p>}
  </details>;
}
