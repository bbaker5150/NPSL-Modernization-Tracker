import { describe, expect, it } from 'vitest';
import { decodeProjectEngineers, engineerFields, projectEngineers } from './projectEngineers';
import { isOwnedByUser, normalizeDirectory } from './identity';

const first = { name: 'First Engineer', email: 'first@example.test', loginName: 'claims|first@example.test' };
const second = { name: 'Second Engineer', email: 'second@example.test', loginName: 'claims|second@example.test' };

describe('project team identities and legacy storage', () => {
  it('retains both engineers through serialization and mirrors the first for legacy consumers', () => {
    const row = engineerFields([first, second]);
    expect(row).toMatchObject({ engineers: [first, second], ownerName: first.name, ownerKey: first.loginName, ownerEmail: first.email });
    expect(decodeProjectEngineers(JSON.stringify(row.engineers), { ownerKey: 'stale' })).toEqual([first, second]);
    expect(isOwnedByUser(row, second)).toBe(true);
    expect(isOwnedByUser(row, { loginName: 'someone-else' })).toBe(false);
  });
  it('does not grant access through another claims account sharing the same email or a matching name', () => {
    const row = engineerFields([first, second]);
    expect(isOwnedByUser(row, { ...second, loginName: 'claims|different-account' })).toBe(false);
    expect(isOwnedByUser({ ownerName: second.name }, { title: second.name })).toBe(false);
    expect(projectEngineers({ ownerName: second.name })).toEqual([{ name: second.name, email: '', loginName: '' }]);
  });
  it('deduplicates identities without collapsing distinct resolved accounts sharing a mailbox', () => {
    const other = { ...first, loginName: 'claims|other' };
    expect(engineerFields([first, { ...first, loginName: first.loginName.toUpperCase() }, other]).engineers).toEqual([first, other]);
  });
  it('treats absent old columns as legacy, but an explicit empty team clears all stale owners', () => {
    const legacy = { ownerName: first.name, ownerEmail: first.email, ownerKey: first.loginName };
    for (const value of [undefined, null, '']) expect(decodeProjectEngineers(value, legacy)).toEqual([first]);
    expect(decodeProjectEngineers('[]', legacy)).toEqual([]);
    expect(isOwnedByUser({ ...legacy, engineers: [] }, first)).toBe(false);
    expect(engineerFields([])).toEqual({ engineers: [], ownerName: 'Unassigned', ownerEmail: '', ownerKey: '' });
  });
  it.each(['not json', '{}', 'null', '[null]', '[{}]', '[{"loginName":12}]'])('stops malformed storage %s instead of falling back to the first owner', value => {
    expect(() => decodeProjectEngineers(value, { ownerKey: first.loginName })).toThrow('EngineersJson');
  });
  it('migrates legacy team-member roles but never promotes an explicit Viewer implicitly', () => {
    const project = engineerFields([first, second]);
    const old = { ...second, role: 'Contributor' };
    expect(normalizeDirectory([old], [project])[0].role).toBe('Project Engineer');
    expect(normalizeDirectory([{ ...old, role: 'Viewer' }], [project])[0].role).toBe('Viewer');
  });
});
