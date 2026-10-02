export const normalizePhaseKey = (key) => ({
  'need-scope': 'requirement', 'research-tds': 'procurement',
  'requirements-acquisition': 'acquisition', 'procurement-support': 'procurement',
  'test-evaluation': 'procurement', 'integration-deployment': 'deployment',
  'sustainment-closeout': 'deployment', development: 'procurement',
  'production-procurement': 'procurement', 'operation-sustainment': 'deployment',
}[key] || key || 'requirement');

export const ORGANIZATIONS = ['CHENG Team', 'NPSL', 'NPSL Metrology Engineering'];
export const normalizeOrganization = (value) => ({ 'Program Office': 'CHENG Team', ISE: 'NPSL Metrology Engineering' }[value] || (ORGANIZATIONS.includes(value) ? value : 'NPSL'));
export const ROLES = ['Viewer', 'Project Engineer', 'Manager'];
export const normalizeRole = (value) => ROLES.includes(value) ? value : 'Viewer';
export function normalizeTaskOrganizations(tasks, projects) {
  const byKey = new Map(projects.map((project) => [project.projectKey, project]));
  return tasks.map((task) => ({ ...task, organization: normalizeOrganization(task.organization || byKey.get(task.projectKey)?.organization) }));
}
export const normalizeTaskStatus = (value) => value === 'In Progress – At Program Office' ? 'In Progress' : (value || 'Not Started');

export const workflowData = {
  phases: [
    { key: 'requirement', name: 'Requirement (MSA)', short: 'MSA', acronym: 'MSA', description: 'Requirement / Materiel Solution Analysis. ICP and workload analysis.' },
    { key: 'acquisition', name: 'Acquisition (TMRR)', short: 'TMRR', acronym: 'TMRR', description: 'Acquisition / Technology Maturation and Risk Reduction.' },
    { key: 'procurement', name: 'Procurement (EMD)', short: 'EMD', acronym: 'EMD', description: 'Procurement / Engineering and Manufacturing Development. Includes legacy Development and Production/Procurement tasks.' },
    { key: 'deployment', name: 'Deployment (P&D)', short: 'P&D', acronym: 'P&D', description: 'Deployment / Production and Deployment. Includes legacy Operation and Sustainment tasks.' },
  ],
  taskTemplates: [
    { title: 'Complete ICP and workload analysis', phaseKey: 'requirement', order: 1 },
    { title: 'Complete acquisition and risk-reduction activities', phaseKey: 'acquisition', order: 2 },
    { title: 'Complete development, procurement, and acceptance testing', phaseKey: 'procurement', order: 3 },
    { title: 'Publish PMSD/METBUL and transition to deployment', phaseKey: 'deployment', order: 4 },
  ],
};

export function projectProgress(tasks, mode = 'phases') {
  const completedTasks = tasks.filter((task) => task.status === 'Complete').length;
  const completedPhases = workflowData.phases.filter((phase) => {
    const rows = tasks.filter((task) => normalizePhaseKey(task.phaseKey) === phase.key);
    return rows.length > 0 && rows.every((task) => ['Complete', 'Not Required', 'Not Applicable'].includes(task.status));
  }).length;
  const completed = mode === 'tasks' ? completedTasks : completedPhases;
  const total = mode === 'tasks' ? tasks.length : workflowData.phases.length;
  return { percentComplete: total ? Math.round(completed / total * 100) : 0, progressDetail: `${completed} of ${total} ${mode === 'tasks' ? 'tasks completed' : 'phases resolved'}` };
}
