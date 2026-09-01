import { describe, expect, it } from 'vitest';
import { seedData } from './seed';

describe('August portfolio seed data', () => {
  it('preserves every workbook project and its phase/action tasks', () => {
    expect(seedData.projects).toHaveLength(28);
    expect(seedData.tasks).toHaveLength(161);
    expect(new Set(seedData.tasks.map((task) => task.projectKey))).toEqual(new Set(seedData.projects.map((project) => project.projectKey)));
    expect(seedData.projects.find((project) => project.title === 'Electrical Modernization 04')).toMatchObject({ ownerName: 'Electrical Engineer A', targetFinish: '2026-09-30' });
  });

  it('maps every task to the five supplied pipeline phases', () => {
    const phases = new Set(seedData.phases.map((phase) => phase.key));
    expect(seedData.phases.map((phase) => phase.short)).toEqual(['RP', 'DP', 'AP', 'O&S', 'PP']);
    expect(seedData.tasks.every((task) => phases.has(task.phaseKey))).toBe(true);
  });

  it('includes the supplied acronym reference without legacy organization labels', () => {
    expect(seedData.glossary).toHaveLength(37);
    expect(seedData.glossary.find((entry) => entry.acronym === 'CSS')?.term).toBe('Calibration Standard Specification');
    expect(JSON.stringify(seedData)).not.toMatch(/NPSL/i);
  });
});
