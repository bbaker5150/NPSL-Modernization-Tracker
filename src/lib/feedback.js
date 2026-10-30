// Feedback is deliberately separate from shared tracker data and exports.
const unwrap = value => value?.d || value;
const rows = value => value?.value || value?.d?.results || [];
export function validateFeedback(input) {
  const title = String(input.title || '').trim();
  const details = String(input.details || '').trim();
  if (!['Bug', 'Feature'].includes(input.kind)) throw new Error('Choose a bug report or feature request.');
  if (!title || title.length > 200) throw new Error('Enter a summary of 1–200 characters.');
  if (!details || details.length > 10000) throw new Error('Enter details of 1–10,000 characters.');
  return { kind: input.kind, title, details };
}
export function createFeedback(store) {
  const title = `${String(store.prefix || 'Modernization').replace(/[^A-Za-z0-9]/g, '')}Feedback`;
  const path = `/_api/web/lists/getbytitle('${title}')`;
  // Local mode is a disposable demonstration, never a private data store.
  const preview = [];
  const local = !store.webUrl;
  async function isOwner() {
    const user = await store.currentUser();
    if (user.isSiteAdmin) return true;
    if (local) return false;
    // Match the tracker's site-owner definition: Manage Permissions on the WEB.
    // Effective permissions include direct grants and security-group membership;
    // currentuser/groups alone can omit these. List/project permissions and app
    // Manager roles must never qualify a user as a site owner.
    let permissionError;
    try {
      const response = unwrap(await store.get('/_api/web/effectivebasepermissions'));
      const bits = response?.EffectiveBasePermissions || response;
      if ((BigInt(bits?.Low || 0) & 33554432n) !== 0n) return true;
    } catch (error) { permissionError = error; }
    const group = unwrap(await store.get('/_api/web/associatedownergroup?$select=Id'));
    if (Number.isInteger(group?.Id)) {
      const groups = rows(await store.get('/_api/web/currentuser/groups?$select=Id&$top=5000'));
      if (groups.some(entry => entry.Id === group.Id)) return true;
    }
    if (permissionError) throw permissionError;
    return false;
  }
  return {
    local,
    isOwner,
    async submit(input) {
      const value = validateFeedback(input);
      if (local) { preview.unshift({ ...value, id: crypto.randomUUID(), created: new Date().toISOString(), author: 'Local preview' }); return; }
      // No read-before-write or return-body display: submitters have Add Items only.
      await store.post(`${path}/items`, { body: { Title: value.title, FeedbackKind: value.kind, FeedbackDetails: value.details } });
    },
    async list() {
      if (!await isOwner()) throw new Error('Only site owners can view submissions.');
      if (local) return [...preview];
      const result = []; let next = `${path}/items?$select=Id,Title,FeedbackKind,FeedbackDetails,Created,Author/Title&$expand=Author&$orderby=Created desc&$top=100`;
      while (next) {
        const body = await store.get(next);
        result.push(...rows(body).map(row => ({ id: row.Id, title: row.Title, kind: row.FeedbackKind, details: row.FeedbackDetails, created: row.Created, author: row.Author?.Title || 'Unknown' })));
        const link = body['@odata.nextLink'] || body['odata.nextLink'] || body.d?.__next;
        if (!link) break;
        const url = new URL(link, `${store.webUrl}/`);
        if (!url.href.startsWith(`${store.webUrl}/_api/`)) throw new Error('Unexpected feedback pagination address.');
        next = url.href.slice(store.webUrl.length);
      }
      return result;
    },
  };
}
