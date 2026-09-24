import { displayName } from './displayName';
import { projectProgress } from '../data/workflow';
import ExcelJS from 'exceljs';

const NAVY = '0B2942';
const BLUE = '1F73C9';
const SKY = 'DDEEFF';
const GOLD = 'D6A84B';
const PALE = 'F4F7FA';
const WHITE = 'FFFFFF';
const INK = '172B3A';
const MUTED = '5F7282';
const LINE = 'D8E1E8';
const isResolvedTask = (task) => ['Complete', 'Not Required', 'Not Applicable'].includes(task.status);

const toDate = (value) => {
  if (!value) return null;
  const parsed = new Date(`${String(value).slice(0, 10)}T12:00:00`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const phaseName = (phases, key) => phases.find((phase) => phase.key === key)?.name || key || '';

function safeSheetName(name) {
  return String(name).replace(/[\\/?*:[\]]/g, ' ').slice(0, 31);
}

function styleDataSheet(sheet, columns, statusColumns = []) {
  sheet.views = [{ state: 'frozen', ySplit: 1, xSplit: 1 }];
  sheet.properties.defaultRowHeight = 20;
  sheet.getRow(1).height = 28;
  sheet.getRow(1).eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
    cell.font = { name: 'Aptos Display', size: 10, bold: true, color: { argb: WHITE } };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
    cell.border = { bottom: { style: 'medium', color: { argb: GOLD } } };
  });
  columns.forEach((column, index) => {
    const excelColumn = sheet.getColumn(index + 1);
    excelColumn.width = column.width;
    excelColumn.hidden = /(?: ID|Identity Key)$/.test(column.header);
    excelColumn.alignment = { vertical: 'top', wrapText: !!column.wrap };
    if (column.numFmt) excelColumn.numFmt = column.numFmt;
  });
  for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex += 1) {
    const row = sheet.getRow(rowIndex);
    const wrapLines = columns.reduce((max, column, index) => {
      if (!column.wrap) return max;
      const length = String(row.getCell(index + 1).value || '').length;
      return Math.max(max, Math.ceil(length / Math.max(12, column.width * 1.35)));
    }, 1);
    row.height = Math.min(72, Math.max(21, wrapLines * 13));
    row.font = { name: 'Aptos', size: 9, color: { argb: INK } };
    row.eachCell((cell) => {
      cell.border = { bottom: { style: 'hair', color: { argb: LINE } } };
      cell.alignment = { ...cell.alignment, vertical: 'top' };
    });
  }
  statusColumns.forEach((columnIndex) => {
    for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex += 1) {
      const cell = sheet.getCell(rowIndex, columnIndex);
      const value = String(cell.value || '');
      let fill;
      let color;
      if (/complete|not required|not applicable|on track|closed/i.test(value)) { fill = 'E3F5EE'; color = '18795B'; }
      else if (/blocked|critical|at risk|program office/i.test(value)) { fill = 'FCE8E8'; color = 'B4232B'; }
      else if (/review|high|progress|possible|medium/i.test(value)) { fill = 'FFF2D9'; color = '9A5B00'; }
      if (fill) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
        cell.font = { ...cell.font, bold: true, color: { argb: color } };
      }
    }
  });
  sheet.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.25, right: 0.25, top: 0.45, bottom: 0.45, header: 0.2, footer: 0.2 } };
  sheet.headerFooter.oddFooter = '&LModernization Project Tracker&CPage &P of &N&RExported &D';
}

function addDataSheet(workbook, { name, tableName, columns, rows, statusColumns = [] }) {
  const sheet = workbook.addWorksheet(safeSheetName(name), { properties: { tabColor: { argb: BLUE } } });
  const headers = columns.map((column) => column.header);
  sheet.addRow(headers);
  rows.forEach((row) => sheet.addRow(row));
  if (rows.length) {
    sheet.addTable({
      name: tableName,
      ref: 'A1',
      headerRow: true,
      totalsRow: false,
      style: { theme: 'TableStyleMedium2', showRowStripes: true },
      columns: headers.map((header) => ({ name: header, filterButton: true })),
      rows,
    });
  }
  styleDataSheet(sheet, columns, statusColumns);
  return sheet;
}

function addSummary(workbook, { projects, tasks, updates, risks, phases, user, sourceLabel }) {
  const sheet = workbook.addWorksheet('Portfolio Summary', { properties: { tabColor: { argb: GOLD } } });
  sheet.views = [{ showGridLines: false, state: 'frozen', ySplit: 6 }];
  sheet.mergeCells('A1:H2');
  const title = sheet.getCell('A1');
  title.value = 'Modernization Project Tracker';
  title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
  title.font = { name: 'Aptos Display', size: 22, bold: true, color: { argb: WHITE } };
  title.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  sheet.getRow(1).height = 25;
  sheet.getRow(2).height = 25;

  const generated = new Date();
  const metadata = [
    ['Export source', sourceLabel],
    ['Generated for', displayName(user?.title || user?.email || 'SharePoint user')],
    ['Generated on', generated],
  ];
  metadata.forEach(([label, value], index) => {
    const row = index + 3;
    sheet.getCell(row, 1).value = label;
    sheet.getCell(row, 1).font = { name: 'Aptos', size: 9, bold: true, color: { argb: MUTED } };
    sheet.mergeCells(row, 2, row, 4);
    sheet.getCell(row, 2).value = value;
    sheet.getCell(row, 2).font = { name: 'Aptos', size: 9, color: { argb: INK } };
  });
  sheet.getCell('B5').numFmt = 'mmm d, yyyy h:mm AM/PM';

  const completedTasks = tasks.filter(isResolvedTask).length;
  const openRisks = risks.filter((risk) => risk.status !== 'Closed').length;
  const blockedTasks = tasks.filter((task) => task.status === 'Blocked').length;
  const avg = projects.length ? Math.round(projects.reduce((sum, project) => sum + Number(project.percentComplete || 0), 0) / projects.length) : 0;
  const cards = [
    ['Projects', projects.length, 'Projects in this export scope'],
    ['Portfolio progress', `${avg}%`, `${completedTasks} of ${tasks.length} tasks resolved`],
    ['Open risks', openRisks, `${risks.length} risks captured`],
    ['Blocked tasks', blockedTasks, `${updates.length} status updates`],
  ];
  cards.forEach(([label, value, detail], index) => {
    const start = 1 + index * 2;
    sheet.mergeCells(7, start, 7, start + 1);
    sheet.mergeCells(8, start, 9, start + 1);
    sheet.mergeCells(10, start, 10, start + 1);
    const labelCell = sheet.getCell(7, start);
    const valueCell = sheet.getCell(8, start);
    const detailCell = sheet.getCell(10, start);
    [labelCell, valueCell, detailCell].forEach((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: PALE } };
      cell.border = { left: { style: 'thin', color: { argb: LINE } }, right: { style: 'thin', color: { argb: LINE } } };
    });
    labelCell.border.top = { style: 'thin', color: { argb: LINE } };
    detailCell.border.bottom = { style: 'thin', color: { argb: LINE } };
    labelCell.value = label;
    labelCell.font = { name: 'Aptos', size: 9, bold: true, color: { argb: BLUE } };
    valueCell.value = value;
    valueCell.font = { name: 'Aptos Display', size: 20, bold: true, color: { argb: NAVY } };
    detailCell.value = detail;
    detailCell.font = { name: 'Aptos', size: 8, color: { argb: MUTED } };
    [labelCell, valueCell, detailCell].forEach((cell) => { cell.alignment = { vertical: 'middle', horizontal: 'center' }; });
  });

  sheet.mergeCells('A12:H12');
  sheet.getCell('A12').value = 'PIPELINE DISTRIBUTION';
  sheet.getCell('A12').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SKY } };
  sheet.getCell('A12').font = { name: 'Aptos', size: 10, bold: true, color: { argb: NAVY } };
  sheet.getCell('A12').alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  sheet.addRow(['Stage', 'Pipeline step', 'Projects', 'Average progress', 'Blocked tasks', 'Open risks', 'Next milestone due', 'Notes']);
  phases.forEach((phase, index) => {
    const inStage = projects.filter((project) => project.currentStageKey === phase.key);
    const keys = new Set(inStage.map((project) => project.projectKey));
    const stageTasks = tasks.filter((task) => keys.has(task.projectKey));
    const stageRisks = risks.filter((risk) => keys.has(risk.projectKey) && risk.status !== 'Closed');
    const dates = inStage.map((project) => toDate(project.nextMilestoneDate)).filter(Boolean).sort((a, b) => a - b);
    sheet.addRow([
      index + 1,
      phase.name,
      inStage.length,
      inStage.length ? inStage.reduce((sum, project) => sum + Number(project.percentComplete || 0), 0) / inStage.length / 100 : 0,
      stageTasks.filter((task) => task.status === 'Blocked').length,
      stageRisks.length,
      dates[0] || null,
      phase.short || '',
    ]);
  });
  [9, 28, 12, 18, 15, 12, 20, 20].forEach((width, index) => { sheet.getColumn(index + 1).width = width; });
  const pipelineHeader = sheet.getRow(13);
  pipelineHeader.height = 27;
  pipelineHeader.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
    cell.font = { name: 'Aptos', size: 9, bold: true, color: { argb: WHITE } };
  });
  for (let row = 14; row <= 13 + phases.length; row += 1) {
    sheet.getRow(row).font = { name: 'Aptos', size: 9, color: { argb: INK } };
    sheet.getRow(row).eachCell((cell) => {
      cell.border = { bottom: { style: 'hair', color: { argb: LINE } } };
      cell.alignment = { vertical: 'top', wrapText: false };
    });
    sheet.getCell(row, 4).numFmt = '0%';
    sheet.getCell(row, 7).numFmt = 'mmm d, yyyy';
  }
  sheet.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 1 };
  return sheet;
}

export function createPortfolioWorkbook({ projects, tasks, updates, risks, phases, glossary = [], user, sourceLabel = 'Live SharePoint portfolio' }) {
  const keys = new Set(projects.map((project) => project.projectKey));
  tasks = tasks.filter((task) => keys.has(task.projectKey));
  risks = risks.filter((risk) => keys.has(risk.projectKey));
  updates = updates.filter((update) => keys.has(update.projectKey));
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Modernization Project Tracker';
  workbook.lastModifiedBy = displayName(user?.title || user?.email || 'SharePoint user');
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.subject = sourceLabel;
  workbook.title = 'Modernization Portfolio Export';
  workbook.description = 'Portfolio, project, pipeline task, risk, update, and pipeline reference data.';
  workbook.calcProperties.fullCalcOnLoad = true;

  addSummary(workbook, { projects, tasks, updates, risks, phases, user, sourceLabel });
  addDataSheet(workbook, {
    name: 'Projects', tableName: 'ProjectsTable', statusColumns: [4],
    columns: [
      { header: 'Project', width: 34, wrap: true }, { header: 'Owner', width: 24 },
      { header: 'Stage', width: 26 }, { header: 'Health', width: 16 },
      { header: 'Progress', width: 12, numFmt: '0%' }, { header: 'Progress View', width: 18 },
      { header: 'Next Milestone', width: 40, wrap: true }, { header: 'Milestone Date', width: 16, numFmt: 'mmm d, yyyy' },
      { header: 'Target Finish', width: 16, numFmt: 'mmm d, yyyy' }, { header: 'Owner Email', width: 30 },
      { header: 'Phase Progress', width: 15, numFmt: '0%' }, { header: 'Task Progress', width: 15, numFmt: '0%' },
    ],
    rows: projects.map((project) => [project.title, displayName(project.ownerName || 'Unassigned'), phaseName(phases, project.currentStageKey), project.health,
      Number(project.percentComplete || 0) / 100, project.progressMode === 'tasks' ? 'By task' : 'By phase', project.nextMilestone || '', toDate(project.nextMilestoneDate), toDate(project.targetFinish), project.ownerEmail || '',
      projectProgress(tasks.filter((task) => task.projectKey === project.projectKey), 'phases').percentComplete / 100,
      projectProgress(tasks.filter((task) => task.projectKey === project.projectKey), 'tasks').percentComplete / 100]),
  });
  addDataSheet(workbook, {
    name: 'Needs Attention', tableName: 'AttentionTable', statusColumns: [1],
    columns: [{ header: 'Status', width: 32, wrap: true }, { header: 'Project', width: 32, wrap: true }, { header: 'Task', width: 44, wrap: true }, { header: 'Owner', width: 24 }, { header: 'Due Date', width: 16, numFmt: 'mmm d, yyyy' }, { header: 'Deferred Date', width: 16, numFmt: 'mmm d, yyyy' }, { header: 'Deferral Reason', width: 40, wrap: true }],
    rows: tasks.filter((task) => !isResolvedTask(task)).sort((a, b) => a.status.localeCompare(b.status)).map((task) => [task.status, projects.find((p) => p.projectKey === task.projectKey)?.title || '', task.title, displayName(displayName(task.ownerName || 'Unassigned')), toDate(task.dueDate), toDate(task.deferredDate), task.deferredJustification || '']),
  });
  addDataSheet(workbook, {
    name: 'Tasks', tableName: 'TasksTable', statusColumns: [7],
    columns: [
      { header: 'Task ID', width: 22 }, { header: 'Project ID', width: 22 }, { header: 'Project', width: 30 },
      { header: 'Task', width: 46, wrap: true }, { header: 'Pipeline Stage', width: 26 }, { header: 'Sort Order', width: 11 },
      { header: 'Status', width: 16 }, { header: 'Owner', width: 22 }, { header: 'Owner Email', width: 30 }, { header: 'Owner Identity Key', width: 36 },
      { header: 'Start Date', width: 15, numFmt: 'mmm d, yyyy' }, { header: 'Due Date', width: 15, numFmt: 'mmm d, yyyy' },
      { header: 'Finish Date', width: 15, numFmt: 'mmm d, yyyy' }, { header: 'Blocked Reason', width: 36, wrap: true },
      { header: 'Notes / Data Issue', width: 44, wrap: true },
      { header: 'Deferred Date', width: 15, numFmt: 'mmm d, yyyy' }, { header: 'Deferral Justification', width: 44, wrap: true }, { header: 'Not Required Justification', width: 44, wrap: true },
    ],
    rows: tasks.map((task) => [
      task.id, task.projectKey, projects.find((project) => project.projectKey === task.projectKey)?.title || '', task.title,
      phaseName(phases, task.phaseKey), Number(task.order || 0), task.status, displayName(task.ownerName || 'Unassigned'), task.ownerEmail || '', task.ownerKey || '',
      toDate(task.startDate), toDate(task.dueDate), toDate(task.finishDate), task.blockedReason || '', [task.notes, task.dataIssue].filter(Boolean).join(' | '), toDate(task.deferredDate), task.deferredJustification || '', task.notRequiredJustification || '',
    ]),
  });
  addDataSheet(workbook, {
    name: 'Risks', tableName: 'RisksTable', statusColumns: [5, 7],
    columns: [
      { header: 'Risk ID', width: 22 }, { header: 'Project ID', width: 22 }, { header: 'Project', width: 30 }, { header: 'Risk / Issue', width: 38, wrap: true },
      { header: 'Severity', width: 13 }, { header: 'Probability', width: 14 }, { header: 'Status', width: 13 }, { header: 'Owner', width: 22 },
      { header: 'Owner Identity Key', width: 36 }, { header: 'Due Date', width: 16, numFmt: 'mmm d, yyyy' }, { header: 'Mitigation', width: 46, wrap: true },
    ],
    rows: risks.map((risk) => [risk.id, risk.projectKey, projects.find((project) => project.projectKey === risk.projectKey)?.title || '', risk.title, risk.severity, risk.probability, risk.status, displayName(risk.ownerName || 'Unassigned'), risk.ownerKey || '', toDate(risk.dueDate), risk.mitigation || '']),
  });
  addDataSheet(workbook, {
    name: 'Updates', tableName: 'UpdatesTable',
    columns: [
      { header: 'Update ID', width: 22 }, { header: 'Project ID', width: 22 }, { header: 'Project', width: 30 }, { header: 'Date', width: 16, numFmt: 'mmm d, yyyy' },
      { header: 'Type', width: 14 }, { header: 'Author', width: 22 }, { header: 'Author Email', width: 30 }, { header: 'Author Identity Key', width: 36 },
      { header: 'Summary', width: 64, wrap: true },
    ],
    rows: updates.map((update) => [update.id, update.projectKey, projects.find((project) => project.projectKey === update.projectKey)?.title || '', toDate(update.entryDate), update.type, displayName(update.authorName), update.authorEmail, update.authorKey || '', update.summary]),
  });
  addDataSheet(workbook, {
    name: 'Pipeline Reference', tableName: 'PipelineTable',
    columns: [{ header: 'Stage', width: 10 }, { header: 'Key', width: 24 }, { header: 'Pipeline Step', width: 34 }, { header: 'Acronym', width: 14 }, { header: 'Description', width: 54, wrap: true }],
    rows: phases.map((phase, index) => [index + 1, phase.key, phase.name, phase.short || '', phase.description || '']),
  });
  addDataSheet(workbook, {
    name: 'Acronym Glossary', tableName: 'AcronymGlossaryTable',
    columns: [{ header: 'Acronym', width: 18 }, { header: 'Full Term', width: 44, wrap: true }, { header: 'Definition', width: 76, wrap: true }],
    rows: glossary.map((entry) => [entry.acronym, entry.term, entry.definition]),
  });
  const summary = workbook.getWorksheet('Portfolio Summary');
  let linkRow = summary.rowCount + 3;
  for (const sheet of workbook.worksheets.filter((sheet) => sheet !== summary)) {
    summary.getCell(linkRow++, 1).value = { text: `Open ${sheet.name}`, hyperlink: `#'${sheet.name}'!A1` };
  }
  return workbook;
}

export async function downloadPortfolioWorkbook(options) {
  const workbook = createPortfolioWorkbook(options);
  const bytes = await workbook.xlsx.writeBuffer();
  const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `modernization-portfolio-${new Date().toISOString().slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(link.href), 0);
}
