import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it, vi } from 'vitest';
import { TrackerAccessSetup } from './TrackerAccessSetup';
import { TrackerAccessJob } from '../lib/trackerAccessJob';

globalThis.IS_REACT_ACT_ENVIRONMENT = true;
it('restores the running status when returning to Users and managers without starting another job', async () => {
  let finish, progress;
  const store = { scopedAccess: true, prepareTrackerAccess: vi.fn(report => {
    progress = report; return new Promise(resolve => { finish = resolve; });
  }) };
  store.trackerAccessJob = new TrackerAccessJob(store);
  const container = document.createElement('div'); document.body.append(container);
  let root = createRoot(container);
  try {
    await act(async () => root.render(<TrackerAccessSetup store={store} />));
    await act(async () => container.querySelector('button').click());
    await act(async () => root.unmount());
    progress('Checking tasks access · 129–132 of 200');
    root = createRoot(container);
    await act(async () => root.render(<TrackerAccessSetup store={store} />));
    expect(container.querySelector('button').disabled).toBe(true);
    expect(container.querySelector('[role="status"]').textContent).toContain('129–132');
    expect(store.prepareTrackerAccess).toHaveBeenCalledOnce();
    await act(async () => { finish({ users: 2, projects: 1 }); await store.trackerAccessJob.running; });
    expect(container.querySelector('button').disabled).toBe(false);
    expect(container.querySelector('[role="status"]').textContent).toContain('Verified tracker groups');
  } finally { await act(async () => root.unmount()); container.remove(); }
});

it('exposes an explicit repair action without running it when rendered', async () => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  const node = document.createElement('div');
  const root = createRoot(node);
  const snapshot = { busy: false, status: '', error: '' };
  const job = { subscribe: () => () => {}, getSnapshot: () => snapshot, run: vi.fn() };
  try {
    await act(async () => root.render(<TrackerAccessSetup store={{ scopedAccess: true, trackerAccessJob: job }} />));
    expect(node.textContent).toContain('Repair shared Read and manager access');
    expect(job.run).not.toHaveBeenCalled();
    await act(async () => node.querySelector('button').click());
    expect(job.run).toHaveBeenCalledOnce();
  } finally { await act(async () => root.unmount()); }
});
