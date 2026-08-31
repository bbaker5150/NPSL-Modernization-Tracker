import fs from 'node:fs/promises';
import path from 'node:path';

const sourcePath = path.resolve(process.argv[2] || '../workbook_analysis/sheets.json');
const outputPath = path.resolve(process.argv[3] || 'src/data/seed.js');
const sheets = JSON.parse(await fs.readFile(sourcePath, 'utf8'));

const phases = [
  { key: 'need-scope', name: 'Need & Scope', short: 'Scope', prefixes: ['1.3.2'] },
  { key: 'research-tds', name: 'Research & TDS', short: 'TDS', prefixes: ['1.3.3'] },
  { key: 'requirements-acquisition', name: 'Requirements & Acquisition', short: 'Acquisition', prefixes: ['1.3.5'] },
  { key: 'procurement-support', name: 'Procurement Support', short: 'Procurement', prefixes: ['1.3.6'] },
  { key: 'test-evaluation', name: 'Test Planning & Evaluation', short: 'Test & Eval', prefixes: ['1.4', '1.5'] },
  { key: 'integration-deployment', name: 'Integration & Deployment', short: 'Integration', prefixes: ['1.6'] },
  { key: 'sustainment-closeout', name: 'Sustainment & Closeout', short: 'Closeout', prefixes: ['1.7'] },
];

function slug(value) {
  return String(value).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
}

function excelDate(value, isStart = false) {
  if (typeof value === 'number') {
    const millis = Date.UTC(1899, 11, 30) + Math.round(value) * 86400000;
    return new Date(millis).toISOString().slice(0, 10);
  }
  const match = /^FY(\d{2})$/i.exec(String(value || '').trim());
  if (match) {
    const endYear = 2000 + Number(match[1]);
    return isStart ? `${endYear - 1}-10-01` : `${endYear}-09-30`;
  }
  return null;
}

function phaseFor(wbs) {
  return phases.find((phase) => phase.prefixes.some((prefix) => wbs === prefix || wbs.startsWith(`${prefix}.`))) || phases[0];
}

const tasks = [];
const projects = sheets.map((sheet, projectIndex) => {
  const projectKey = slug(sheet.name);
  const rows = sheet.values.slice(1).filter((row) => row[0] && row[1]);
  const projectTasks = rows.map((row, taskIndex) => {
    const [wbs, title, startRaw, finishRaw, note] = row;
    const startDate = excelDate(startRaw, true);
    const finishDate = excelDate(finishRaw, false);
    const status = finishDate ? 'Complete' : startDate ? 'In Progress' : 'Not Started';
    const phase = phaseFor(String(wbs));
    const dataIssue = startDate && finishDate && startDate > finishDate ? 'Start date is after finish date in source workbook.' : '';
    const task = {
      id: `${projectKey}-${String(taskIndex + 1).padStart(2, '0')}`,
      projectKey,
      wbs: String(wbs),
      title: String(title).trim(),
      phaseKey: phase.key,
      order: taskIndex + 1,
      status,
      startDate,
      finishDate,
      dueDate: finishDate,
      ownerName: 'Unassigned',
      ownerEmail: '',
      notes: note ? String(note).trim() : '',
      blockedReason: '',
      sourceStartLabel: typeof startRaw === 'string' ? startRaw : '',
      dataIssue,
    };
    tasks.push(task);
    return task;
  });

  const completed = projectTasks.filter((task) => task.status === 'Complete').length;
  const active = projectTasks.filter((task) => task.status === 'In Progress').length;
  const latestPhase = [...phases].reverse().find((phase) => projectTasks.some((task) => task.phaseKey === phase.key && task.status !== 'Not Started')) || phases[0];
  const dated = projectTasks.flatMap((task) => [task.startDate, task.finishDate]).filter(Boolean).sort();
  const nextTask = projectTasks.find((task) => task.status !== 'Complete');
  const sourceNotes = projectTasks.filter((task) => task.notes).map((task) => task.notes);
  const dateIssues = projectTasks.filter((task) => task.dataIssue).length;

  return {
    id: `project-${String(projectIndex + 1).padStart(2, '0')}`,
    projectKey,
    title: sheet.name,
    measurementArea: sheet.name,
    description: `${sheet.name} measurement-area modernization project imported from POAM_NPSL_Mod_Projects.xlsx.`,
    ownerName: 'Unassigned',
    ownerEmail: '',
    managerName: 'Unassigned',
    priority: 'Medium',
    health: 'Needs Review',
    status: completed === projectTasks.length ? 'Complete' : active ? 'In Progress' : 'Planned',
    currentStageKey: latestPhase.key,
    percentComplete: Math.round((completed / projectTasks.length) * 100),
    targetFinish: dated.at(-1) || '',
    nextMilestone: nextTask?.title || 'Final Report',
    nextMilestoneDate: nextTask?.dueDate || '',
    sourceNotes: sourceNotes.join(' • '),
    importedBaseline: true,
    dataIssueCount: dateIssues,
    tags: [],
  };
});

const seed = { source: 'POAM_NPSL_Mod_Projects.xlsx', importedAt: '2026-08-31', phases, projects, tasks, updates: [], risks: [] };
const contents = `// Generated from POAM_NPSL_Mod_Projects.xlsx. Do not edit by hand.\nexport const seedData = ${JSON.stringify(seed, null, 2)};\n`;
await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.writeFile(outputPath, contents, 'utf8');
console.log(JSON.stringify({ projects: projects.length, tasks: tasks.length, outputPath }, null, 2));
