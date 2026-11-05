import React, { useState, useSyncExternalStore } from 'react';
import { TRACKER_USERS_GROUP } from '../lib/trackerGroups';

export function TrackerAccessSetup({ store }) {
  if (!store.scopedAccess || !store.trackerAccessJob) return null;
  return <AccessSetup store={store} job={store.trackerAccessJob} />;
}

function AccessSetup({ store, job }) {
  const { busy, status, error } = useSyncExternalStore(job.subscribe, job.getSnapshot);
  const [glossaryBusy, setGlossaryBusy] = useState(false);
  const [glossaryError, setGlossaryError] = useState('');
  const [glossaryReady, setGlossaryReady] = useState(false);
  async function setUpGlossary() {
    if (glossaryBusy) return;
    setGlossaryBusy(true); setGlossaryError(''); setGlossaryReady(false);
    try { await store.prepareGlossaryAccess(); setGlossaryReady(true); }
    catch (caught) { setGlossaryError(caught.message || 'Could not verify glossary access. Try again.'); }
    finally { setGlossaryBusy(false); }
  }
  return <details className="panel">
    <summary>Tracker permissions</summary>
    <p>Apply saved elevated roles to Tracker Project Engineers and Tracker Managers. Basic access comes from {TRACKER_USERS_GROUP}; use Invite people to enroll a new member. This setup repairs Read and manager access on projects, task/update/risk folders, and individually secured records. Reconcile assigned-engineer editing access while preserving Owners and unrelated permissions.</p>
    <p>Run once as a site owner after deployment or importing records. Existing permissions are checked in small batches; matching items are skipped. If Firepit reloads the app, run this again to finish.</p>
    <button type="button" className="button primary" disabled={busy} onClick={() => job.run()}>{busy ? 'Checking and applying permissions…' : error ? 'Retry tracker permissions' : 'Apply tracker permissions'}</button>
    {status && <p role="status" aria-live="polite">{status}</p>}
    {error && <p role="alert" className="inline-error">{error}</p>}
    {store.sessionSiteOwner && <section aria-label="Glossary access">
      <h3>Glossary access</h3>
      <p>Allow Tracker users to add acronyms. Managers can also remove them.</p>
      <button type="button" className="button secondary" disabled={glossaryBusy} onClick={setUpGlossary}>{glossaryBusy ? 'Checking glossary access…' : 'Set up glossary access'}</button>
      {glossaryReady && <p role="status">SharePoint glossary access verified.</p>}
      {glossaryError && <p role="alert" className="inline-error">{glossaryError}</p>}
    </section>}
  </details>;
}
