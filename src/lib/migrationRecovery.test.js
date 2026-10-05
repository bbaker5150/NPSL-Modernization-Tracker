// @vitest-environment node
import { expect, it, vi } from 'vitest';
import { retryMigrationRead, reconcileMigrationWrite } from './migrationRecovery';

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
