import { describe, expect, it } from 'vitest';
import { seedData } from './seed';

describe('POA&M seed data', () => {
  it('preserves every workbook project and WBS task', () => {
    expect(seedData.projects).toHaveLength(14);
    expect(seedData.tasks).toHaveLength(14 * 34);
    expect(new Set(seedData.tasks.map((task) => task.projectKey))).toEqual(new Set(seedData.projects.map((project) => project.projectKey)));
  });

  it('maps every task to the seven-stage pipeline', () => {
    const phases = new Set(seedData.phases.map((phase) => phase.key));
    expect(seedData.phases).toHaveLength(7);
    expect(seedData.tasks.every((task) => phases.has(task.phaseKey))).toBe(true);
  });

  it('flags impossible source date orderings instead of silently correcting them', () => {
    const flagged = seedData.tasks.filter((task) => task.dataIssue);
    expect(flagged.length).toBeGreaterThan(0);
    expect(flagged.every((task) => task.startDate > task.finishDate)).toBe(true);
  });
});
