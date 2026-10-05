// @vitest-environment node
import { expect, it, vi } from 'vitest';
import { retryMigrationRead, reconcileMigrationWrite, migrationRequest } from './migrationRecovery';

it('retries transient reads with bounded backoff and honors Retry-After', async () => {
  const wait = vi.fn();
  const operation = vi.fn().mockRejectedValueOnce(new Error('SharePoint request timed out')).mockRejectedValueOnce(Object.assign(new Error('Throttled'), { status: 429, retryAfterMs: 5000 })).mockResolvedValue('ok');
  expect(await retryMigrationRead(operation, wait)).toBe('ok');
  expect(wait.mock.calls).toEqual([[1000], [5000]]);
  operation.mockReset().mockRejectedValue(new Error('SharePoint request timed out'));
  await expect(retryMigrationRead(operation, wait)).rejects.toThrow('timed out');
  expect(operation).toHaveBeenCalledTimes(4);
});
it('does not retry permission failures or retry before a long server cooldown', async () => {
  for (const error of [Object.assign(new Error('Denied'), { status: 403 }), Object.assign(new Error('Throttled'), { status: 429, retryAfterMs: 60000 })]) {
    const operation = vi.fn().mockRejectedValue(error), wait = vi.fn();
    await expect(retryMigrationRead(operation, wait)).rejects.toThrow(error.message);
    expect(operation).toHaveBeenCalledTimes(1); expect(wait).not.toHaveBeenCalled();
  }
});
it('reconciles a late committed write without submitting it twice', async () => {
  const operation = vi.fn().mockRejectedValue(new Error('SharePoint request timed out'));
  const verify = vi.fn().mockResolvedValueOnce(null).mockResolvedValueOnce({ value: 42 });
  expect(await reconcileMigrationWrite(operation, verify, { wait: vi.fn() })).toBe(42);
  expect(operation).toHaveBeenCalledTimes(1);
  expect(verify).toHaveBeenCalledTimes(2);
});
it('stops an unconfirmed write instead of risking a duplicate or false success', async () => {
  const operation = vi.fn().mockRejectedValue(new Error('SharePoint request timed out'));
  const verify = vi.fn().mockResolvedValue(null);
  await expect(reconcileMigrationWrite(operation, verify, { wait: vi.fn(), label: 'Tasks/task-7' })).rejects.toThrow('Tasks/task-7:');
  expect(operation).toHaveBeenCalledTimes(1); expect(verify).toHaveBeenCalledTimes(3);
});

it('bounds a hung fetch and a hung response body even when the host ignores abort', async () => {
  vi.useFakeTimers();
  try {
    for (const hangsInBody of [false, true]) {
      let release;
      const hung = new Promise(resolve => { release = resolve; });
      const transport = vi.fn(async () => hangsInBody ? { status: 200, headers: new Headers(), arrayBuffer: () => hung } : hung);
      const result = migrationRequest(transport, 'https://example.invalid/_api/list', { method: 'POST' }, 1000);
      const rejected = expect(result).rejects.toMatchObject({ status: 408 });
      await vi.advanceTimersByTimeAsync(1000); await rejected;
      expect(transport).toHaveBeenCalledOnce();
      expect(transport.mock.calls[0][1].signal.aborted).toBe(true);
      release(hangsInBody ? new ArrayBuffer(0) : new Response('{}'));
      await Promise.resolve();
      expect(vi.getTimerCount()).toBe(0);
    }
    const response = await migrationRequest(async () => new Response(null, { status: 204 }), 'https://example.invalid/_api/list');
    expect(response.status).toBe(204);
    expect(vi.getTimerCount()).toBe(0);
  } finally { vi.useRealTimers(); }
});
