// @vitest-environment node
import { describe, expect, it, vi } from 'vitest';
import { fixture, record } from './backupTestFixtures';
import { previewMigration, runMigration } from './siteMigration';
import { exportTrackerBackup, importTrackerBackup, previewTrackerImport } from './trackerBackup';

const first = { name: 'Engineer One', email: 'one@example.invalid', loginName: 'i:0#.f|membership|one@example.invalid' };
const second = { name: 'Engineer Two', email: 'two@example.invalid', loginName: 'i:0#.f|membership|two@example.invalid' };
const legacyOwner = { OwnerName: first.name, OwnerEmail: first.email, OwnerKey: first.loginName };

function withoutTeamColumn(site) {
  const schema = site.schema.getMockImplementation();
  site.schema.mockImplementation(async container => {
    const fields = await schema(container);
    if (container.key === 'projects') fields.delete('EngineersJson');
    return fields;
  });
  const read = site.readRows.getMockImplementation();
  site.readRows.mockImplementation((key, fields, ...args) => {
    if (key === 'projects' && fields.includes('EngineersJson')) throw new Error('The legacy source has no EngineersJson column.');
    return read(key, fields, ...args);
  });
}

function batchTarget(target) {
  target.writeBatch = vi.fn(async (key, operations) => {
    for (const operation of operations) {
      if (operation.id) await target.update(key, operation.id, operation.fields, operation.etag);
      else await target.create(key, operation.fields);
    }
  });
}

describe('project teams in site migration and backup import', () => {
  it('reviews and replaces a newer destination team when importing a legacy single assignment', async () => {
    const { source, target } = fixture();
    withoutTeamColumn(source);
    Object.assign(source.rows.projects[0], legacyOwner);
    target.rows.projects.push(record('p', { Id: 20, ProjectKey: 'project-key', ...legacyOwner, EngineersJson: JSON.stringify([first, second]) }));
    const before = structuredClone(source.rows.projects);
    const plan = await previewMigration(source, target);
    const projects = plan.lists.find(list => list.key === 'projects');
    expect(projects.fields).toContain('EngineersJson');
    expect(projects.sourceFields).not.toContain('EngineersJson');
    expect(projects.entries[0].conflict).toBe(true);
    expect(JSON.parse(projects.entries[0].source.EngineersJson)).toEqual([first]);
    await expect(runMigration(source, target, plan)).rejects.toThrow('explicitly approve');
    expect(target.update).not.toHaveBeenCalled();
    expect((await runMigration(source, target, plan, { replaceConflicts: true })).verified).toBe(true);
    expect(JSON.parse(target.rows.projects[0].EngineersJson)).toEqual([first]);
    const writes = target.update.mock.calls.length;
    await runMigration(source, target, await previewMigration(source, target));
    expect(target.update).toHaveBeenCalledTimes(writes);
    expect(source.rows.projects).toEqual(before);
    expect(source.update).not.toHaveBeenCalled();
    expect(source.readRows.mock.calls.filter(([key]) => key === 'projects').every(([, fields]) => !fields.includes('EngineersJson'))).toBe(true);
  });

  it('restores an old backup assignment without retaining an unrelated destination teammate', async () => {
    const { source, target } = fixture();
    withoutTeamColumn(source);
    Object.assign(source.rows.projects[0], legacyOwner);
    const backup = JSON.parse(await exportTrackerBackup(source));
    expect(backup.payload.lists.find(list => list.key === 'projects').fields).not.toContain('EngineersJson');
    target.rows.projects.push(record('p', { Id: 20, ProjectKey: 'project-key', ...legacyOwner, EngineersJson: JSON.stringify([first, second]) }));
    batchTarget(target);
    source.readRows.mockRejectedValue(new Error('Source is offline'));
    const plan = await previewTrackerImport(backup, target);
    await expect(importTrackerBackup(backup, target, plan)).rejects.toThrow('approve');
    expect(target.writeBatch).not.toHaveBeenCalled();
    expect((await importTrackerBackup(backup, target, plan, { replaceConflicts: true })).verified).toBe(true);
    expect(JSON.parse(target.rows.projects[0].EngineersJson)).toEqual([first]);
    const writes = target.writeBatch.mock.calls.length;
    await importTrackerBackup(backup, target, await previewTrackerImport(backup, target));
    expect(target.writeBatch).toHaveBeenCalledTimes(writes);
  });

  it('clears the destination team when a legacy source is unassigned', async () => {
    const { source, target } = fixture();
    withoutTeamColumn(source);
    source.rows.projects[0].OwnerName = 'Unassigned';
    target.rows.projects.push(record('p', { Id: 20, ProjectKey: 'project-key', ...legacyOwner, EngineersJson: JSON.stringify([first, second]) }));
    await runMigration(source, target, await previewMigration(source, target), { replaceConflicts: true });
    expect(target.rows.projects[0]).toMatchObject({ OwnerName: 'Unassigned', EngineersJson: '[]' });
  });

  it('roundtrips the complete team and rejects a destination that cannot store it', async () => {
    const { source, target } = fixture();
    Object.assign(source.rows.projects[0], legacyOwner, { EngineersJson: JSON.stringify([first, second]) });
    batchTarget(target);
    const backup = JSON.parse(await exportTrackerBackup(source));
    await importTrackerBackup(backup, target, await previewTrackerImport(backup, target));
    expect(JSON.parse(target.rows.projects[0].EngineersJson)).toEqual([first, second]);
    withoutTeamColumn(target);
    await expect(previewMigration(source, target)).rejects.toThrow('destination needs fields EngineersJson');
  });

  it('rejects malformed source teams before migration or backup import can write', async () => {
    const { source, target } = fixture();
    Object.assign(source.rows.projects[0], legacyOwner, { EngineersJson: '{broken' });
    await expect(previewMigration(source, target)).rejects.toThrow('Projects/p');
    await expect(exportTrackerBackup(source)).rejects.toThrow('Projects/p');
    expect(target.create).not.toHaveBeenCalled();
    expect(target.update).not.toHaveBeenCalled();
  });

  it('leaves legacy-to-legacy transfers compatible without adding an unavailable field', async () => {
    const { source, target } = fixture();
    withoutTeamColumn(source);
    withoutTeamColumn(target);
    Object.assign(source.rows.projects[0], legacyOwner);
    const plan = await previewMigration(source, target);
    expect(plan.lists.find(list => list.key === 'projects').fields).not.toContain('EngineersJson');
    expect((await runMigration(source, target, plan)).verified).toBe(true);
    expect(target.rows.projects[0]).toMatchObject(legacyOwner);
    expect(target.rows.projects[0]).not.toHaveProperty('EngineersJson');
  });
});
