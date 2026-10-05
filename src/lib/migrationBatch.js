export function makeMigrationBatch(webUrl, listRoot, operations) {
  if (!operations.length || operations.length > 25) throw new Error('A migration batch must contain 1–25 records.');
  const boundary = `batch_${crypto.randomUUID()}`;
  const lines = [];
  for (const [index, op] of operations.entries()) {
    const change = `changeset_${crypto.randomUUID()}`;
    const item = op.id == null ? '' : `/items(${Number(op.id)})`;
    if (op.id != null && (!Number.isSafeInteger(op.id) || op.id <= 0 || !/^"\d+"$/.test(op.etag || ''))) throw new Error('Existing records require an item ID and a valid ETag. Refresh preview.');
    lines.push(`--${boundary}`, `Content-Type: multipart/mixed; boundary=${change}`, '', `--${change}`, 'Content-Type: application/http', 'Content-Transfer-Encoding: binary', `Content-ID: ${index + 1}`, '',
      `POST ${webUrl}${listRoot}${item || '/items'} HTTP/1.1`, 'Accept: application/json;odata=nometadata', 'Content-Type: application/json;odata=nometadata',
      ...(op.id == null ? [] : ['X-HTTP-Method: MERGE', `If-Match: ${op.etag}`]), '', JSON.stringify(op.fields), `--${change}--`, '');
  }
  lines.push(`--${boundary}--`, '');
  const body = lines.join('\r\n');
  if (new TextEncoder().encode(body).byteLength > 4 * 1024 * 1024) throw new Error('Batch is too large; no records in this batch were sent.');
  return { body, contentType: `multipart/mixed; boundary=${boundary}` };
}

export function checkMigrationBatch(body, expected) {
  const statuses = [...body.matchAll(/^HTTP\/1\.[01]\s+(\d{3})\b/gm)].map(match => Number(match[1]));
  if (statuses.length !== expected || statuses.some(status => status < 200 || status >= 300)) {
    throw new Error(`SharePoint did not confirm every batch operation (${statuses.join(', ') || 'unrecognized response'}). Some records may be saved. Preview again to resume.`);
  }
}
