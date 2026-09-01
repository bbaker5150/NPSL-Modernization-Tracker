import { describe, expect, it } from 'vitest';
import { workflowData } from './workflow';

describe('modernization workflow definition', () => {
  it('contains only the five pipeline phases and clean starter task templates', () => {
    expect(workflowData.phases.map((phase) => phase.short)).toEqual(['RP', 'DP', 'AP', 'O&S', 'PP']);
    expect(workflowData.taskTemplates).toHaveLength(5);
    expect(workflowData.taskTemplates.map((task) => task.phaseKey)).toEqual(workflowData.phases.map((phase) => phase.key));
    expect(JSON.stringify(workflowData)).not.toMatch(/mock|sample|NPSL/i);
  });
});
