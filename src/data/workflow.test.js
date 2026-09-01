import { describe, expect, it } from 'vitest';
import { workflowData } from './workflow';
import { defaultAcronyms } from './defaultAcronyms';

describe('modernization workflow definition', () => {
  it('contains only the five pipeline phases and clean starter task templates', () => {
    expect(workflowData.phases.map((phase) => phase.short)).toEqual(['RP', 'DP', 'AP', 'O&S', 'PP']);
    expect(workflowData.taskTemplates).toHaveLength(5);
    expect(workflowData.taskTemplates.map((task) => task.phaseKey)).toEqual(workflowData.phases.map((phase) => phase.key));
    expect(JSON.stringify(workflowData)).not.toMatch(/mock|sample|NPSL/i);
  });

  it('retains the supplied acronym reference independently from mock portfolio data', () => {
    expect(defaultAcronyms).toHaveLength(37);
    expect(defaultAcronyms.find((entry) => entry.acronym === 'CSS')?.term).toBe('Calibration Standard Specification');
  });
});
