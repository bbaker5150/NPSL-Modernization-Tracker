import React, { useSyncExternalStore } from 'react';
import { TRACKER_USERS_GROUP } from '../lib/trackerGroups';

export function TrackerAccessSetup({ store }) {
  if (!store.scopedAccess || !store.trackerAccessJob) return null;
  return <AccessSetup job={store.trackerAccessJob} />;
}

function AccessSetup({ job }) {
  const { busy, status, error } = useSyncExternalStore(job.subscribe, job.getSnapshot);
  return <details className="panel">
    <summary>Tracker permissions</summary>
    <p>Apply saved elevated roles to Tracker Project Engineers and Tracker Managers. Basic access comes from {TRACKER_USERS_GROUP}; use Invite people to enroll a new member. This setup repairs Read and manager access on projects, task/update/risk folders, and individually secured records. Reconcile assigned-engineer editing access while preserving Owners and unrelated permissions.</p>
    <p>Run once as a site owner after deployment or importing records. Existing permissions are checked in small batches; matching items are skipped. If Firepit reloads the app, run this again to finish.</p>
    <button type="button" className="button primary" disabled={busy} onClick={() => job.run()}>{busy ? 'Checking and applying permissions…' : error ? 'Retry tracker permissions' : 'Apply tracker permissions'}</button>
    {status && <p role="status" aria-live="polite">{status}</p>}
    {error && <p role="alert" className="inline-error">{error}</p>}
  </details>;
}
