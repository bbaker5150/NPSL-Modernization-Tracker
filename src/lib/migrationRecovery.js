export const migrationDelay = ms => new Promise(resolve => setTimeout(resolve, ms));

export function isTransientMigrationError(error) {
  if (error?.status) return [408, 429, 502, 503, 504].includes(Number(error.status));
  return /timed?\s*out|timeout|failed to fetch|networkerror|network request failed/i.test(error?.message || '');
}

// GETs are safe to repeat. Writes must instead be reconciled with destination
// state: a timed-out request may still have committed on the server.
export async function retryMigrationRead(operation, wait = migrationDelay) {
  for (let attempt = 0; ; attempt++) {
    try { return await operation(); }
    catch (error) {
      if (!isTransientMigrationError(error) || attempt >= 3) throw error;
      const delay = Math.max(1000 * 2 ** attempt, error.retryAfterMs || 0);
      // Do not ignore a long server-directed cooldown by retrying early.
      if (delay > 30000) throw error;
      await wait(delay);
    }
  }
}

export async function reconcileMigrationWrite(operation, verify, { onProgress = () => {}, wait = migrationDelay, label = 'record' } = {}) {
  try { return await operation(); }
  catch (error) {
    if (!isTransientMigrationError(error)) throw error;
    for (let attempt = 0; attempt < 3; attempt++) {
      const delay = Math.max(1000 * 2 ** attempt, error.retryAfterMs || 0);
      if (delay > 30000) break;
      onProgress(`Checking whether ${label} was saved after a timeout (${attempt + 1}/3)…`);
      await wait(delay);
      const confirmed = await verify();
      if (confirmed) return confirmed.value;
    }
    throw new Error(`${label}: ${error.message} The write result could not be confirmed; it was not sent again. Refresh the preview to check the destination before resuming.`);
  }
}
