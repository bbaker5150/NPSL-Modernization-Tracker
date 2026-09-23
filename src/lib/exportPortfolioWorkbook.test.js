import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { workflowData } from '../data/workflow';
import { createPortfolioWorkbook } from './exportPortfolioWorkbook';

describe('portfolio Excel export', () => {
  it('creates a styled, typed, multi-sheet workbook containing the complete portfolio', async () => {
    const projects = [{ id: 'project-1', projectKey: 'project-1', title: 'Calibration modernization', measurementArea: 'Electrical', description: 'Replace an aging standard.', ownerName: 'Test Engineer', ownerEmail: 'engineer@example.invalid', ownerKey: 'engineer', health: 'On Track', priority: 'High', status: 'In Progress', currentStageKey: 'requirement', percentComplete: 20, targetFinish: '2027-01-15', nextMilestone: 'Approve requirements', nextMilestoneDate: '2026-10-01', blockedCount: 0, overdueCount: 0 }];
    const tasks = workflowData.taskTemplates.map((task, index) => ({ ...task, id: `task-${index + 1}`, projectKey: 'project-1', status: index === 0 ? 'Complete' : 'Not Started', ownerName: 'Test Engineer', ownerEmail: 'engineer@example.invalid', ownerKey: 'engineer', startDate: '', dueDate: '', finishDate: '', blockedReason: '', notes: '', dataIssue: '' }));
    const glossary = [{ id: 'acronym-1', acronym: 'CSS', term: 'Calibration Standard Specification', definition: 'Technical requirements for a calibration standard.' }];
    const workbook = createPortfolioWorkbook({
      projects,
      tasks,
      updates: [],
      risks: [],
      phases: workflowData.phases,
      glossary,
      user: { title: 'Test Engineer', email: 'engineer@example.invalid' },
      sourceLabel: 'Live SharePoint portfolio',
    });
    const bytes = await workbook.xlsx.writeBuffer();
    expect(bytes.byteLength).toBeGreaterThan(18_000);
    const projectsSheet = workbook.getWorksheet('Projects');
    expect(projectsSheet.autoFilter).toBeNull();
    expect(projectsSheet.getTable('ProjectsTable').table.totalsRow).toBe(false);

    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(bytes);
    expect(reopened.worksheets.map((sheet) => sheet.name)).toEqual([
      'Portfolio Summary', 'Projects', 'Tasks', 'Risks', 'Updates', 'Pipeline Reference', 'Acronym Glossary',
    ]);
    expect(reopened.getWorksheet('Projects').rowCount).toBe(projects.length + 1);
    expect(reopened.getWorksheet('Tasks').rowCount).toBe(tasks.length + 1);
    expect(reopened.getWorksheet('Projects').getCell('A1').font.bold).toBe(true);
    expect(reopened.getWorksheet('Projects').getCell('A1').fill.fgColor.argb).toBe('0B2942');
    expect(reopened.getWorksheet('Projects').getCell('M2').numFmt).toBe('0%');
    expect(reopened.getWorksheet('Pipeline Reference').rowCount).toBe(workflowData.phases.length + 1);
    expect(reopened.getWorksheet('Acronym Glossary').rowCount).toBe(glossary.length + 1);
    expect(reopened.getWorksheet('Portfolio Summary').getCell('A1').value).toBe('Modernization Project Tracker');
  }, 20_000);
});
