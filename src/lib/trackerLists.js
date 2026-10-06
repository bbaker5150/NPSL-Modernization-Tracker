import { SharePointError } from './spContext';

export const TRACKER_LISTS = [
  ['projects', 'Projects', 'Projects'], ['tasks', 'Tasks', 'Tasks'],
  ['updates', 'Updates', 'Updates'], ['risks', 'Risks', 'Risks'],
  ['acronyms', 'Acronyms', 'Acronyms'], ['users', 'Users', 'Users'],
  ['references', 'ReferenceDocuments', 'Reference Documents'],
];
const namespace = prefix => String(prefix || 'Modernization').replace(/[^A-Za-z0-9]/g, '') || 'Modernization';
const definition = key => {
  const entry = TRACKER_LISTS.find(([name]) => name === key);
  if (!entry) throw new Error(`Unknown tracker list: ${key}`);
  return entry;
};
export const legacyListTitle = (prefix, key) => namespace(prefix) + definition(key)[1];
export const displayListTitle = (prefix, key) => `${namespace(prefix) === 'Modernization' ? 'NPSL Tracker' : `${namespace(prefix)} Tracker`} - ${definition(key)[2]}`;
const literal = value => `'${String(value).replace(/'/g, "''")}'`;
const guidPath = id => {
  if (!/^[\da-f]{8}(-[\da-f]{4}){3}-[\da-f]{12}$/i.test(id || '')) throw new Error('SharePoint returned an invalid tracker list ID.');
  return `/_api/web/lists(guid'${id}')`;
};

// Resolve both names, then pin all requests to the existing list GUID. Renaming
// display titles does not move records, attachments, folders, URLs, or ACLs.
export class TrackerLists {
  constructor(prefix, get, post) { this.prefix = prefix; this.get = get; this.post = post; this.cache = new Map(); }
  async resolve(key) {
    if (!this.cache.has(key)) {
      const promise = (async () => {
        const oldTitle = legacyListTitle(this.prefix, key), title = displayListTitle(this.prefix, key);
        const filter = `Title eq ${literal(oldTitle)} or Title eq ${literal(title)}`;
        const body = await this.get(`/_api/web/lists?$select=Id,Title,Hidden,BaseTemplate&$filter=${encodeURIComponent(filter)}`);
        const matches = body.value || body.d?.results;
        if (!Array.isArray(matches)) throw new Error(`Cannot read the tracker list catalog for ${title}.`);
        if (matches.length > 1) throw new Error(`Both ${oldTitle} and ${title} exist. Resolve the duplicate lists before continuing; no list was selected.`);
        if (!matches.length) throw new SharePointError(`Tracker list not found: ${title}`, 404);
        const row = matches[0];
        if (![oldTitle, title].includes(row.Title) || Number(row.BaseTemplate) !== 100) throw new Error(`Unexpected list returned for ${title}.`);
        return { ...row, path: guidPath(row.Id) };
      })();
      this.cache.set(key, promise);
      promise.catch(() => { if (this.cache.get(key) === promise) this.cache.delete(key); });
    }
    return this.cache.get(key);
  }
  async path(key) { return (await this.resolve(key)).path; }
  async rewrite(path) {
    const match = String(path).match(/\/_api\/web\/lists\/getbytitle\('((?:''|[^'])*)'\)/i);
    if (!match) return path;
    const title = decodeURIComponent(match[1]).replace(/''/g, "'");
    const key = TRACKER_LISTS.find(([key]) => [legacyListTitle(this.prefix, key), displayListTitle(this.prefix, key)].includes(title))?.[0];
    return key ? path.replace(match[0], await this.path(key)) : path;
  }
  async organize(onProgress = () => {}) {
    const body = await this.get('/_api/web/EffectiveBasePermissions');
    const permissions = body.d || body;
    if ((BigInt((permissions.EffectiveBasePermissions || permissions).Low || 0) & 2048n) === 0n) throw new Error('A site owner with Manage Lists permission must organize the tracker lists.');
    this.cache.clear();
    // Preflight every list before the first write. Never create substitutes for
    // missing/ambiguous lists, and support rerunning after a partial completion.
    const rows = await Promise.all(TRACKER_LISTS.map(async ([key]) => [key, await this.resolve(key)]));
    for (const [key, row] of rows) {
      const title = displayListTitle(this.prefix, key);
      onProgress(`Making ${title} visible…`);
      if (row.Title !== title || row.Hidden !== false) {
        await this.post(row.path, { body: { Title: title, Hidden: false }, headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' } });
      }
      const response = await this.get(`${row.path}?$select=Id,Title,Hidden`);
      const verified = response.d || response;
      if (String(verified.Id).toLowerCase() !== row.Id.toLowerCase() || verified.Title !== title || verified.Hidden !== false) throw new Error(`Could not verify ${title}. Run organization again to finish.`);
      this.cache.set(key, Promise.resolve({ ...row, ...verified }));
    }
    return rows.length;
  }
}
