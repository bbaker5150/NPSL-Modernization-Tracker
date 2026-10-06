import React, { useState } from 'react';

export function TrackerListSetup({ store }) {
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState('');
  const [error, setError] = useState('');
  if (!store.organizeTrackerLists) return null;
  async function organize() {
    if (busy) return;
    setBusy(true); setError(''); setStatus('Checking tracker lists…');
    try {
      const count = await store.organizeTrackerLists(setStatus);
      setStatus(`All ${count} tracker lists are visible in Site Contents with matching app names.`);
    } catch (caught) {
      setStatus('');
      setError(`${caught.message} Completed changes remain in place; you can run this again.`);
    } finally { setBusy(false); }
  }
  return <details className="panel">
    <summary>Site Contents organization</summary>
    <p>Run once as a site owner to show the seven existing lists together under the NPSL Tracker prefix. Records, documents, and permissions stay in place.</p>
    <button type="button" className="button primary" disabled={busy} onClick={organize}>{busy ? 'Organizing lists…' : 'Organize tracker lists'}</button>
    {status && <p role="status" aria-live="polite">{status}</p>}
    {error && <p role="alert" className="inline-error">{error}</p>}
  </details>;
}
