import { expect, it, vi } from 'vitest';
import { TrackerAccessJob } from './trackerAccessJob';

it('keeps one running job when the screen unmounts and remounts', async () => {
  let finish, report;
  const store = { prepareTrackerAccess: vi.fn(progress => {
    report = progress; return new Promise(resolve => { finish = resolve; });
  }) };
  const job = new TrackerAccessJob(store);
  const left = vi.fn(), unsubscribe = job.subscribe(left);
  const running = job.run(); unsubscribe();
  expect(job.run()).toBe(running);
  await Promise.resolve();
  report('Checking tasks access · 129–132');
  expect(job.getSnapshot()).toMatchObject({ busy: true, status: 'Checking tasks access · 129–132' });
  const returned = vi.fn(); job.subscribe(returned);
  finish({ users: 1, projects: 1 }); await running;
  expect(returned).toHaveBeenCalled();
  expect(store.prepareTrackerAccess).toHaveBeenCalledOnce();
  expect(job.state.busy).toBe(false);
});

it('retries through live permission checks after a failure or a new app instance', async () => {
  const store = { prepareTrackerAccess: vi.fn().mockRejectedValueOnce(new Error('Host interrupted')).mockResolvedValue({ users: 1, projects: 1 }) };
  const job = new TrackerAccessJob(store);
  await job.run();
  expect(job.state.error).toContain('Host interrupted');
  await job.run();
  expect(job.state).toMatchObject({ busy: false, error: '' });
  await new TrackerAccessJob(store).run();
  expect(store.prepareTrackerAccess).toHaveBeenCalledTimes(3);
});
