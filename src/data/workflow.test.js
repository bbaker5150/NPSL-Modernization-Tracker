import { describe, expect, it } from 'vitest';
import { workflowData, normalizePhaseKey, projectProgress } from './workflow';
import { defaultAcronyms } from './defaultAcronyms';

describe('modernization workflow definition', () => {
  it('contains only the four pipeline phases and clean starter task templates', () => {
    expect(workflowData.phases.map((phase) => phase.short)).toEqual(['MSA', 'TMRR', 'EMD', 'P&D']);
    expect(workflowData.taskTemplates).toHaveLength(4);
    expect(workflowData.taskTemplates.map((task) => task.phaseKey)).toEqual(workflowData.phases.map((phase) => phase.key));
    expect(JSON.stringify(workflowData)).not.toMatch(/mock|sample|NPSL/i);
  });

  it('retains the supplied acronym reference independently from mock portfolio data', () => {
    expect(defaultAcronyms).toHaveLength(37);
    expect(defaultAcronyms.find((entry) => entry.acronym === 'CSS')?.term).toBe('Calibration Standard Specification');
  });
});


describe('progress calculations and phase compatibility', () => {
  it('counts complete tasks across simultaneous phases, rather than requiring phase closure', () => {
    const tasks = [{ phaseKey: 'requirement', status: 'Complete' }, { phaseKey: 'requirement', status: 'In Progress' }, { phaseKey: 'acquisition', status: 'Complete' }, { phaseKey: 'acquisition', status: 'Not Required' }];
    expect(projectProgress(tasks, 'tasks').percentComplete).toBe(50);
    expect(projectProgress(tasks, 'phases').percentComplete).toBe(25);
    expect(projectProgress([], 'tasks').percentComplete).toBe(0);
    expect(projectProgress([], 'phases').percentComplete).toBe(0);
    expect(projectProgress([...tasks, { phaseKey: 'deployment', status: 'Not Started' }], 'tasks').percentComplete).toBe(40);
  });
  it('maps legacy stages without losing tasks', () => {
    expect(['requirement', 'acquisition', 'development', 'production-procurement', 'operation-sustainment'].map(normalizePhaseKey)).toEqual(['requirement', 'acquisition', 'procurement', 'procurement', 'deployment']);
  });
});

it('maps previous program-office names to METCAL without dropping records', async () => {
  const { normalizeOrganization, normalizeTaskOrganizations } = await import('./workflow');
  for (const name of ['PMA 260B', 'Program Office', 'CHENG Team', 'NAVAIR METCAL 260B']) expect(normalizeOrganization(name)).toBe('NAVAIR METCAL 260B');
  expect(normalizeTaskOrganizations([{ id: 'a', projectKey: 'p', subAssignment: 'Team 2 Lead' }], [{ projectKey: 'p', organization: 'PMA 260B' }])[0]).toMatchObject({ organization: 'NAVAIR METCAL 260B', subAssignment: 'Team 2 Lead' });
});
