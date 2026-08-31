import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { seedData } from '../data/seed';
import { createPortfolioWorkbook } from './exportPortfolioWorkbook';

describe('portfolio Excel export', () => {
  it('creates a styled, typed, multi-sheet workbook containing the complete portfolio', async () => {
    const workbook = createPortfolioWorkbook({
      projects: seedData.projects,
      tasks: seedData.tasks,
      updates: seedData.updates,
      risks: seedData.risks,
      phases: seedData.phases,
      user: { title: 'Test Engineer', email: 'engineer@navy.mil' },
      sourceLabel: 'Read-only mock portfolio preview',
    });
    const bytes = await workbook.xlsx.writeBuffer();
    expect(bytes.byteLength).toBeGreaterThan(25_000);

    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(bytes);
    expect(reopened.worksheets.map((sheet) => sheet.name)).toEqual([
      'Portfolio Summary', 'Projects', 'WBS Tasks', 'Risks', 'Updates', 'Pipeline Reference',
    ]);
    expect(reopened.getWorksheet('Projects').rowCount).toBe(seedData.projects.length + 1);
    expect(reopened.getWorksheet('WBS Tasks').rowCount).toBe(seedData.tasks.length + 1);
    expect(reopened.getWorksheet('Projects').getCell('A1').font.bold).toBe(true);
    expect(reopened.getWorksheet('Projects').getCell('A1').fill.fgColor.argb).toBe('0B2942');
    expect(reopened.getWorksheet('Projects').getCell('L2').numFmt).toBe('0%');
    expect(reopened.getWorksheet('Pipeline Reference').rowCount).toBe(seedData.phases.length + 1);
    expect(reopened.getWorksheet('Portfolio Summary').getCell('A1').value).toBe('Modernization Project Tracker');
  }, 20_000);
});
