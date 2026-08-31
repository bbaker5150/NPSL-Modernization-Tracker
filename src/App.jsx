import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Icon } from './components/Icon';
import { createRepository, seedData } from './lib/repository';

const NAV = [
  ['overview', 'Portfolio', 'overview'],
  ['board', 'Pipeline board', 'board'],
  ['projects', 'All projects', 'projects'],
  ['my-work', 'My work', 'user'],
];
const TASK_STATUSES = ['Not Started', 'In Progress', 'Blocked', 'Complete', 'Not Applicable'];
const HEALTHS = ['On Track', 'Needs Review', 'At Risk', 'Blocked'];
const PRIORITIES = ['Low', 'Medium', 'High', 'Critical'];
const todayIso = () => new Date().toISOString().slice(0, 10);
const titleCase = (value) => String(value || '').replace(/(^|[-_])([a-z])/g, (_, space, char) => `${space ? ' ' : ''}${char.toUpperCase()}`);
const displayDate = (value) => value ? new Intl.DateTimeFormat(undefined, { month: 'short', day: 'numeric', year: 'numeric' }).format(new Date(`${String(value).slice(0, 10)}T12:00:00`)) : 'Not set';
const compactDate = (value) => value ? new Intl.DateTimeFormat(undefined, { month: 'short', day: 'numeric' }).format(new Date(`${String(value).slice(0, 10)}T12:00:00`)) : '—';
const uid = (prefix) => `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

function phaseIndex(key) { return Math.max(0, seedData.phases.findIndex((phase) => phase.key === key)); }
function isOverdue(task) { return task.status !== 'Complete' && task.status !== 'Not Applicable' && task.dueDate && task.dueDate < todayIso(); }
function statusTone(value) {
  if (/complete|on track/i.test(value)) return 'good';
  if (/blocked|critical/i.test(value)) return 'bad';
  if (/risk|high|review|progress/i.test(value)) return 'warn';
  return 'neutral';
}

function enrichProject(project, allTasks) {
  const tasks = allTasks.filter((task) => task.projectKey === project.projectKey);
  if (!tasks.length) return { ...project, tasks, blockedCount: 0, overdueCount: 0 };
  const applicable = tasks.filter((task) => task.status !== 'Not Applicable');
  const completed = applicable.filter((task) => task.status === 'Complete').length;
  const latest = [...seedData.phases].reverse().find((phase) => tasks.some((task) => task.phaseKey === phase.key && task.status !== 'Not Started'));
  const nextTask = tasks.sort((a, b) => a.order - b.order).find((task) => !['Complete', 'Not Applicable'].includes(task.status));
  return {
    ...project,
    tasks,
    percentComplete: applicable.length ? Math.round((completed / applicable.length) * 100) : 0,
    currentStageKey: latest?.key || project.currentStageKey,
    nextMilestone: nextTask?.title || project.nextMilestone || 'Final Report',
    nextMilestoneDate: nextTask?.dueDate || project.nextMilestoneDate,
    blockedCount: tasks.filter((task) => task.status === 'Blocked').length,
    overdueCount: tasks.filter(isOverdue).length,
    dataIssueCount: tasks.filter((task) => task.dataIssue).length,
  };
}

function Badge({ children, tone = 'neutral', dot = false }) {
  return <span className={`badge badge-${tone}`}>{dot && <span className="badge-dot" />}{children}</span>;
}

function Progress({ value, compact = false }) {
  return <div className={`progress ${compact ? 'progress-compact' : ''}`}><span style={{ width: `${Math.min(100, Math.max(0, value || 0))}%` }} /></div>;
}

function EmptyState({ title, message, action }) {
  return <div className="empty-state"><div className="empty-icon"><Icon name="projects" size={24} /></div><h3>{title}</h3><p>{message}</p>{action}</div>;
}

function Skeleton() {
  return <div className="loading-shell"><div className="loading-mark">M</div><div><strong>Modernization Tracker</strong><span>Loading portfolio data…</span></div></div>;
}

function SetupGate({ repo, onReady }) {
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState(null);
  const [error, setError] = useState('');

  async function setup(importBaseline) {
    setBusy(true); setError('');
    try {
      await repo.store.provision();
      if (importBaseline) await repo.store.seed(seedData, setProgress);
      await onReady();
    } catch (caught) {
      setError(caught.message || 'SharePoint setup failed.');
      setBusy(false);
    }
  }

  return <div className="gate-shell"><div className="gate-card">
    <div className="brand-mark large">M</div>
    <Badge tone="info">First-time SharePoint setup</Badge>
    <h1>Connect the modernization portfolio</h1>
    <p>This site needs four SharePoint Lists for projects, WBS tasks, updates, and risks. Setup is additive and safe to rerun.</p>
    <div className="list-preview">
      {['ModernizationProjects', 'ModernizationTasks', 'ModernizationUpdates', 'ModernizationRisks'].map((name) => <div key={name}><Icon name="database" /><span>{name}</span><Badge>Missing</Badge></div>)}
    </div>
    {progress && <div className="seed-progress"><div><span>Importing workbook baseline</span><strong>{progress.completed} / {progress.total}</strong></div><Progress value={(progress.completed / progress.total) * 100} /></div>}
    {error && <div className="inline-error"><Icon name="alert" />{error}</div>}
    <div className="gate-actions">
      <button className="button primary" disabled={busy} onClick={() => setup(true)}>{busy ? 'Setting up…' : 'Create lists & import baseline'}</button>
      <button className="button secondary" disabled={busy} onClick={() => setup(false)}>Create empty lists</button>
    </div>
    <p className="fine-print">Requires Edit or Full Control on the current SharePoint site.</p>
  </div></div>;
}

function KpiCard({ icon, label, value, detail, tone = 'blue' }) {
  return <article className={`kpi-card kpi-${tone}`}><div className="kpi-icon"><Icon name={icon} /></div><div><span>{label}</span><strong>{value}</strong><small>{detail}</small></div></article>;
}

function PipelineRail({ projects, selected, onSelect }) {
  const max = Math.max(1, ...seedData.phases.map((phase) => projects.filter((project) => project.currentStageKey === phase.key).length));
  return <div className="pipeline-rail">
    {seedData.phases.map((phase, index) => {
      const count = projects.filter((project) => project.currentStageKey === phase.key).length;
      return <button key={phase.key} className={`phase-node ${selected === phase.key ? 'selected' : ''}`} onClick={() => onSelect?.(selected === phase.key ? '' : phase.key)}>
        <div className="phase-top"><span className="phase-index">{String(index + 1).padStart(2, '0')}</span><strong>{count}</strong></div>
        <span className="phase-label">{phase.name}</span>
        <div className="phase-bar"><span style={{ width: `${Math.max(8, (count / max) * 100)}%` }} /></div>
      </button>;
    })}
  </div>;
}

function ProjectCard({ project, onOpen }) {
  const phase = seedData.phases.find((item) => item.key === project.currentStageKey);
  return <button className="project-card" onClick={() => onOpen(project)}>
    <div className="project-card-top"><Badge tone={statusTone(project.health)} dot>{project.health}</Badge><span className="project-menu">•••</span></div>
    <div><span className="eyebrow">{project.measurementArea}</span><h3>{project.title}</h3></div>
    <div className="project-owner"><span className="avatar small">{project.ownerName === 'Unassigned' ? '?' : project.ownerName.slice(0, 1)}</span><span>{project.ownerName}</span></div>
    <div className="project-phase"><span>Stage {phaseIndex(project.currentStageKey) + 1} of {seedData.phases.length}</span><strong>{phase?.short}</strong></div>
    <Progress value={project.percentComplete} compact />
    <div className="project-card-bottom"><span>{project.percentComplete}% complete</span><span><Icon name="calendar" size={14} /> {compactDate(project.nextMilestoneDate)}</span></div>
  </button>;
}

function Overview({ projects, tasks, risks, onOpen, phaseFilter, setPhaseFilter }) {
  const active = projects.filter((project) => project.status !== 'Complete').length;
  const completedTasks = tasks.filter((task) => task.status === 'Complete').length;
  const attention = projects.filter((project) => ['At Risk', 'Blocked', 'Needs Review'].includes(project.health) || project.blockedCount).length;
  const avg = projects.length ? Math.round(projects.reduce((sum, project) => sum + project.percentComplete, 0) / projects.length) : 0;
  const filtered = phaseFilter ? projects.filter((project) => project.currentStageKey === phaseFilter) : projects;
  const attentionProjects = [...projects].sort((a, b) => (b.blockedCount + b.overdueCount + (b.health === 'At Risk' ? 2 : 0)) - (a.blockedCount + a.overdueCount + (a.health === 'At Risk' ? 2 : 0))).slice(0, 5);

  return <div className="page-stack">
    <section className="hero-row"><div><span className="section-kicker">Portfolio command center</span><h1>Modernization at a glance</h1><p>One view of every measurement area, milestone, and handoff across the NPSL pipeline.</p></div><div className="source-pill"><Icon name="database" /><div><span>Source baseline</span><strong>POAM_NPSL_Mod_Projects.xlsx</strong></div></div></section>
    <section className="kpi-grid">
      <KpiCard icon="projects" label="Active projects" value={active} detail={`${projects.length} total measurement areas`} tone="blue" />
      <KpiCard icon="trend" label="Portfolio progress" value={`${avg}%`} detail={`${completedTasks} of ${tasks.length} tasks complete`} tone="mint" />
      <KpiCard icon="alert" label="Needs attention" value={attention} detail={`${risks.filter((risk) => risk.status !== 'Closed').length} open risks`} tone="amber" />
      <KpiCard icon="clock" label="Blocked tasks" value={tasks.filter((task) => task.status === 'Blocked').length} detail="Across all project teams" tone="rose" />
    </section>
    <section className="panel pipeline-panel"><div className="panel-heading"><div><span className="section-kicker">Portfolio flow</span><h2>Projects by pipeline stage</h2></div>{phaseFilter && <button className="text-button" onClick={() => setPhaseFilter('')}>Clear filter</button>}</div><PipelineRail projects={projects} selected={phaseFilter} onSelect={setPhaseFilter} /></section>
    <section className="dashboard-grid">
      <div className="panel project-showcase"><div className="panel-heading"><div><span className="section-kicker">In motion</span><h2>{phaseFilter ? `${seedData.phases.find((p) => p.key === phaseFilter)?.name} projects` : 'Portfolio projects'}</h2></div><span className="count-label">{filtered.length} projects</span></div>
        <div className="project-card-grid">{filtered.slice(0, 6).map((project) => <ProjectCard key={project.id} project={project} onOpen={onOpen} />)}</div>
      </div>
      <div className="panel attention-panel"><div className="panel-heading"><div><span className="section-kicker">Focus queue</span><h2>Attention needed</h2></div></div>
        <div className="attention-list">{attentionProjects.map((project) => <button key={project.id} onClick={() => onOpen(project)}><span className={`health-pip ${statusTone(project.health)}`} /><div><strong>{project.title}</strong><span>{project.nextMilestone}</span></div><div className="attention-meta"><Badge tone={statusTone(project.health)}>{project.health}</Badge><Icon name="chevron" size={16} /></div></button>)}</div>
      </div>
    </section>
  </div>;
}

function Board({ projects, onOpen }) {
  return <div className="page-stack"><section className="page-heading"><div><span className="section-kicker">End-to-end flow</span><h1>Pipeline board</h1><p>Scan where every project sits and open a card to update its work breakdown.</p></div></section>
    <div className="board-scroll"><div className="kanban-board">{seedData.phases.map((phase, index) => {
      const rows = projects.filter((project) => project.currentStageKey === phase.key);
      return <section className="kanban-column" key={phase.key}><header><div><span>{String(index + 1).padStart(2, '0')}</span><strong>{phase.name}</strong></div><Badge>{rows.length}</Badge></header><div className="kanban-list">{rows.map((project) => <ProjectCard key={project.id} project={project} onOpen={onOpen} />)}{!rows.length && <div className="kanban-empty">No projects in this stage</div>}</div></section>;
    })}</div></div>
  </div>;
}

function ProjectsTable({ projects, onOpen }) {
  const [health, setHealth] = useState('');
  const rows = health ? projects.filter((project) => project.health === health) : projects;
  return <div className="page-stack"><section className="page-heading inline"><div><span className="section-kicker">Portfolio register</span><h1>All projects</h1><p>Detailed ownership, health, stage, and milestone visibility.</p></div><label className="select-wrap"><Icon name="filter" /><select value={health} onChange={(event) => setHealth(event.target.value)}><option value="">All health states</option>{HEALTHS.map((item) => <option key={item}>{item}</option>)}</select></label></section>
    <section className="panel table-panel"><div className="project-table" role="table"><div className="project-table-row table-header" role="row"><span>Project</span><span>Owner</span><span>Stage</span><span>Health</span><span>Progress</span><span>Next milestone</span><span /></div>{rows.map((project) => <button className="project-table-row" role="row" key={project.id} onClick={() => onOpen(project)}><span><strong>{project.title}</strong><small>{project.measurementArea}</small></span><span><span className="avatar tiny">{project.ownerName === 'Unassigned' ? '?' : project.ownerName[0]}</span>{project.ownerName}</span><span>{seedData.phases.find((phase) => phase.key === project.currentStageKey)?.short}</span><span><Badge tone={statusTone(project.health)} dot>{project.health}</Badge></span><span><strong>{project.percentComplete}%</strong><Progress value={project.percentComplete} compact /></span><span><strong>{project.nextMilestone}</strong><small>{displayDate(project.nextMilestoneDate)}</small></span><span><Icon name="chevron" /></span></button>)}</div></section>
  </div>;
}

function MyWork({ projects, tasks, user, onOpen, onTask }) {
  const email = user?.email?.toLowerCase();
  const assigned = projects.filter((project) => project.ownerEmail?.toLowerCase() === email);
  const myTasks = tasks.filter((task) => task.ownerEmail?.toLowerCase() === email && !['Complete', 'Not Applicable'].includes(task.status)).sort((a, b) => (a.dueDate || '9999').localeCompare(b.dueDate || '9999'));
  return <div className="page-stack"><section className="page-heading"><div><span className="section-kicker">Personal workspace</span><h1>My work</h1><p>Projects and actions assigned to {user?.title || user?.email || 'you'}.</p></div></section>
    <section className="dashboard-grid my-work-grid"><div className="panel"><div className="panel-heading"><div><span className="section-kicker">Ownership</span><h2>My projects</h2></div><Badge>{assigned.length}</Badge></div>{assigned.length ? <div className="project-card-grid one-column">{assigned.map((project) => <ProjectCard key={project.id} project={project} onOpen={onOpen} />)}</div> : <EmptyState title="No assigned projects" message="Open an unassigned project and choose Claim project to make it yours." />}</div>
    <div className="panel"><div className="panel-heading"><div><span className="section-kicker">Next actions</span><h2>My task queue</h2></div><Badge>{myTasks.length}</Badge></div><div className="task-queue">{myTasks.slice(0, 14).map((task) => <button key={task.id} onClick={() => onTask(task)}><span className={`task-state ${statusTone(task.status)}`}><Icon name={task.status === 'Blocked' ? 'alert' : 'clock'} size={15} /></span><div><strong>{task.title}</strong><span>{projects.find((project) => project.projectKey === task.projectKey)?.title} · WBS {task.wbs}</span></div><div><Badge tone={isOverdue(task) ? 'bad' : 'neutral'}>{task.dueDate ? compactDate(task.dueDate) : 'No due date'}</Badge></div></button>)}{!myTasks.length && <EmptyState title="Queue clear" message="No open tasks are assigned to your SharePoint email." />}</div></div></section>
  </div>;
}

function Field({ label, children, wide = false }) { return <label className={`field ${wide ? 'field-wide' : ''}`}><span>{label}</span>{children}</label>; }

function Modal({ title, subtitle, onClose, children, actions, wide = false }) {
  useEffect(() => { const listener = (event) => event.key === 'Escape' && onClose(); window.addEventListener('keydown', listener); return () => window.removeEventListener('keydown', listener); }, [onClose]);
  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className={`modal ${wide ? 'modal-wide' : ''}`} role="dialog" aria-modal="true"><header><div><h2>{title}</h2>{subtitle && <p>{subtitle}</p>}</div><button className="icon-button" onClick={onClose} aria-label="Close"><Icon name="close" /></button></header><div className="modal-body">{children}</div>{actions && <footer>{actions}</footer>}</section></div>;
}

function ProjectEditor({ project, user, onClose, onSave }) {
  const [draft, setDraft] = useState({ title: '', measurementArea: '', description: '', ownerName: '', ownerEmail: '', priority: 'Medium', health: 'Needs Review', status: 'Planned', currentStageKey: 'need-scope', targetFinish: '', ...project });
  const set = (key) => (event) => setDraft((row) => ({ ...row, [key]: event.target.value }));
  function claim() { setDraft((row) => ({ ...row, ownerName: user.title, ownerEmail: user.email })); }
  return <Modal title={project?.id ? 'Edit project' : 'Add modernization project'} subtitle="Update the portfolio record and ownership." onClose={onClose} actions={<><button className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" disabled={!draft.title.trim()} onClick={() => onSave({ ...draft, id: draft.id || uid('project'), projectKey: draft.projectKey || draft.title.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, ''), percentComplete: draft.percentComplete || 0, nextMilestone: draft.nextMilestone || 'Define project plan', nextMilestoneDate: draft.nextMilestoneDate || '' })}>Save project</button></>}>
    <div className="field-grid"><Field label="Project name"><input value={draft.title} onChange={set('title')} /></Field><Field label="Measurement area"><input value={draft.measurementArea} onChange={set('measurementArea')} /></Field><Field label="Owner name"><div className="input-action"><input value={draft.ownerName} onChange={set('ownerName')} /><button onClick={claim}>Assign me</button></div></Field><Field label="Owner email"><input type="email" value={draft.ownerEmail} onChange={set('ownerEmail')} /></Field><Field label="Priority"><select value={draft.priority} onChange={set('priority')}>{PRIORITIES.map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Health"><select value={draft.health} onChange={set('health')}>{HEALTHS.map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Status"><select value={draft.status} onChange={set('status')}>{['Planned', 'In Progress', 'On Hold', 'Complete', 'Cancelled'].map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Current stage"><select value={draft.currentStageKey} onChange={set('currentStageKey')}>{seedData.phases.map((phase) => <option value={phase.key} key={phase.key}>{phase.name}</option>)}</select></Field><Field label="Target finish"><input type="date" value={draft.targetFinish || ''} onChange={set('targetFinish')} /></Field><Field label="Description" wide><textarea rows="4" value={draft.description} onChange={set('description')} /></Field></div>
  </Modal>;
}

function TaskEditor({ task, user, onClose, onSave }) {
  const [draft, setDraft] = useState(task);
  const set = (key) => (event) => setDraft((row) => ({ ...row, [key]: event.target.value }));
  return <Modal title="Update task" subtitle={`WBS ${task.wbs} · ${seedData.phases.find((phase) => phase.key === task.phaseKey)?.name}`} onClose={onClose} actions={<><button className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" onClick={() => onSave(draft)}>Save task</button></>}>
    <div className="field-grid"><Field label="Task" wide><input value={draft.title} onChange={set('title')} /></Field><Field label="Status"><select value={draft.status} onChange={set('status')}>{TASK_STATUSES.map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Due date"><input type="date" value={draft.dueDate || ''} onChange={set('dueDate')} /></Field><Field label="Owner name"><div className="input-action"><input value={draft.ownerName || ''} onChange={set('ownerName')} /><button onClick={() => setDraft((row) => ({ ...row, ownerName: user.title, ownerEmail: user.email }))}>Assign me</button></div></Field><Field label="Owner email"><input value={draft.ownerEmail || ''} onChange={set('ownerEmail')} /></Field><Field label="Notes" wide><textarea rows="4" value={draft.notes || ''} onChange={set('notes')} /></Field>{draft.status === 'Blocked' && <Field label="Blocked reason" wide><textarea rows="3" value={draft.blockedReason || ''} onChange={set('blockedReason')} /></Field>}</div>
  </Modal>;
}

function ProjectDrawer({ project, user, tasks, updates, risks, onClose, onEdit, onSaveTask, onAddUpdate, onAddRisk, onClaim }) {
  const [tab, setTab] = useState('overview');
  const [taskEditor, setTaskEditor] = useState(null);
  const [updateText, setUpdateText] = useState('');
  const [riskDraft, setRiskDraft] = useState({ title: '', severity: 'Medium', probability: 'Possible', mitigation: '' });
  const projectTasks = tasks.filter((task) => task.projectKey === project.projectKey).sort((a, b) => a.order - b.order);
  const projectUpdates = updates.filter((item) => item.projectKey === project.projectKey).sort((a, b) => String(b.entryDate).localeCompare(String(a.entryDate)));
  const projectRisks = risks.filter((item) => item.projectKey === project.projectKey);
  const phase = seedData.phases.find((item) => item.key === project.currentStageKey);
  return <div className="drawer-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><aside className="project-drawer" role="dialog" aria-modal="true"><header className="drawer-header"><div className="drawer-title-row"><div><span className="section-kicker">{project.measurementArea}</span><h2>{project.title}</h2><div className="drawer-badges"><Badge tone={statusTone(project.health)} dot>{project.health}</Badge><Badge>{project.priority} priority</Badge><Badge>{project.status}</Badge></div></div><button className="icon-button" onClick={onClose} aria-label="Close"><Icon name="close" /></button></div>
    <div className="drawer-actions"><button className="button secondary" onClick={() => onEdit(project)}>Edit project</button>{!project.ownerEmail && <button className="button primary" onClick={() => onClaim(project)}>Claim project</button>}</div>
    <div className="drawer-progress"><div><span>{project.percentComplete}% complete</span><strong>{phase?.name}</strong></div><Progress value={project.percentComplete} /></div>
    <div className="mini-pipeline">{seedData.phases.map((item, index) => <span key={item.key} className={index < phaseIndex(project.currentStageKey) ? 'done' : item.key === project.currentStageKey ? 'current' : ''} title={item.name}>{index + 1}</span>)}</div>
    <nav className="drawer-tabs">{[['overview', 'Overview'], ['tasks', `Work breakdown (${projectTasks.length})`], ['updates', `Updates (${projectUpdates.length})`], ['risks', `Risks (${projectRisks.length})`]].map(([key, label]) => <button className={tab === key ? 'active' : ''} key={key} onClick={() => setTab(key)}>{label}</button>)}</nav></header>
    <div className="drawer-content">
      {tab === 'overview' && <div className="drawer-section-stack"><section className="detail-grid"><div><span>Owner</span><strong>{project.ownerName}</strong><small>{project.ownerEmail || 'No email assigned'}</small></div><div><span>Target finish</span><strong>{displayDate(project.targetFinish)}</strong><small>{project.status}</small></div><div><span>Next milestone</span><strong>{project.nextMilestone}</strong><small>{displayDate(project.nextMilestoneDate)}</small></div><div><span>Open actions</span><strong>{projectTasks.filter((task) => !['Complete', 'Not Applicable'].includes(task.status)).length}</strong><small>{project.blockedCount} blocked</small></div></section><section className="detail-block"><h3>Purpose</h3><p>{project.description || 'No project description has been added.'}</p></section>{project.sourceNotes && <section className="detail-block source-note"><h3>Workbook notes</h3><p>{project.sourceNotes}</p></section>}{project.dataIssueCount > 0 && <section className="detail-block data-warning"><Icon name="alert" /><div><h3>{project.dataIssueCount} source date issue{project.dataIssueCount === 1 ? '' : 's'}</h3><p>Review tasks where the workbook start date occurs after the finish date.</p></div></section>}<section className="detail-block"><h3>Upcoming work</h3><div className="upcoming-list">{projectTasks.filter((task) => !['Complete', 'Not Applicable'].includes(task.status)).slice(0, 5).map((task) => <button key={task.id} onClick={() => setTaskEditor(task)}><span>{task.wbs}</span><div><strong>{task.title}</strong><small>{seedData.phases.find((item) => item.key === task.phaseKey)?.short}</small></div><Badge tone={statusTone(task.status)}>{task.status}</Badge></button>)}</div></section></div>}
      {tab === 'tasks' && <div className="wbs-list">{seedData.phases.map((stage) => { const stageTasks = projectTasks.filter((task) => task.phaseKey === stage.key); const done = stageTasks.filter((task) => task.status === 'Complete').length; return <section key={stage.key} className="wbs-stage"><header><div><span>{String(phaseIndex(stage.key) + 1).padStart(2, '0')}</span><strong>{stage.name}</strong></div><Badge>{done}/{stageTasks.length}</Badge></header><div>{stageTasks.map((task) => <button key={task.id} className="wbs-task" onClick={() => setTaskEditor(task)}><span className={`task-check ${task.status === 'Complete' ? 'checked' : ''}`}>{task.status === 'Complete' && <Icon name="check" size={13} />}</span><span className="wbs-code">{task.wbs}</span><div><strong>{task.title}</strong>{task.dataIssue && <small className="issue-text"><Icon name="alert" size={12} />{task.dataIssue}</small>}</div><Badge tone={isOverdue(task) ? 'bad' : statusTone(task.status)}>{task.status}</Badge><span className="task-date">{compactDate(task.dueDate)}</span></button>)}</div></section>; })}</div>}
      {tab === 'updates' && <div className="drawer-section-stack"><section className="composer"><textarea rows="3" placeholder="Add a concise status update, decision, or handoff…" value={updateText} onChange={(event) => setUpdateText(event.target.value)} /><div><span>Visible to the project team</span><button className="button primary small-button" disabled={!updateText.trim()} onClick={() => { onAddUpdate(project, updateText); setUpdateText(''); }}>Post update</button></div></section><section className="timeline-list">{projectUpdates.map((item) => <article key={item.id}><span className="timeline-dot" /><div><header><strong>{item.authorName}</strong><span>{displayDate(item.entryDate)}</span></header><Badge>{item.type}</Badge><p>{item.summary}</p></div></article>)}{!projectUpdates.length && <EmptyState title="No updates yet" message="Post the first status update to establish the project history." />}</section></div>}
      {tab === 'risks' && <div className="drawer-section-stack"><section className="risk-composer"><div className="field-grid"><Field label="Risk or issue"><input value={riskDraft.title} onChange={(event) => setRiskDraft((row) => ({ ...row, title: event.target.value }))} /></Field><Field label="Severity"><select value={riskDraft.severity} onChange={(event) => setRiskDraft((row) => ({ ...row, severity: event.target.value }))}>{['Low', 'Medium', 'High', 'Critical'].map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Mitigation" wide><textarea rows="3" value={riskDraft.mitigation} onChange={(event) => setRiskDraft((row) => ({ ...row, mitigation: event.target.value }))} /></Field></div><button className="button primary small-button" disabled={!riskDraft.title.trim()} onClick={() => { onAddRisk(project, riskDraft); setRiskDraft({ title: '', severity: 'Medium', probability: 'Possible', mitigation: '' }); }}>Add risk</button></section><section className="risk-list">{projectRisks.map((risk) => <article key={risk.id}><div><Badge tone={statusTone(risk.severity)}>{risk.severity}</Badge><Badge>{risk.status}</Badge></div><h3>{risk.title}</h3><p>{risk.mitigation || 'No mitigation has been documented.'}</p><span>{risk.ownerName || 'Unassigned'}</span></article>)}{!projectRisks.length && <EmptyState title="No open risks" message="Capture risks and mitigation actions here as the project advances." />}</section></div>}
    </div>
    {taskEditor && <TaskEditor task={taskEditor} user={user} onClose={() => setTaskEditor(null)} onSave={(task) => { onSaveTask(task); setTaskEditor(null); }} />}
  </aside></div>;
}

function Toast({ toast }) { if (!toast) return null; return <div className={`toast toast-${toast.tone || 'good'}`}><Icon name={toast.tone === 'bad' ? 'alert' : 'check'} /><span>{toast.message}</span></div>; }

export function App() {
  const repoRef = useRef(null);
  if (!repoRef.current) repoRef.current = createRepository();
  const repo = repoRef.current;
  const [theme, setTheme] = useState(() => localStorage.getItem('mod-tracker-theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'));
  const [view, setView] = useState('overview');
  const [user, setUser] = useState(null);
  const [data, setData] = useState({ projects: [], tasks: [], updates: [], risks: [] });
  const [loading, setLoading] = useState(true);
  const [needsSetup, setNeedsSetup] = useState(false);
  const [error, setError] = useState('');
  const [search, setSearch] = useState('');
  const [phaseFilter, setPhaseFilter] = useState('');
  const [openProjectId, setOpenProjectId] = useState('');
  const [projectEditor, setProjectEditor] = useState(null);
  const [taskEditor, setTaskEditor] = useState(null);
  const [toast, setToast] = useState(null);

  useEffect(() => { document.documentElement.dataset.theme = theme; localStorage.setItem('mod-tracker-theme', theme); }, [theme]);
  useEffect(() => { initialize(); }, []);
  useEffect(() => { if (!toast) return; const timer = setTimeout(() => setToast(null), 3200); return () => clearTimeout(timer); }, [toast]);

  async function initialize() {
    setLoading(true); setError('');
    try {
      const [currentUser, readiness] = await Promise.all([repo.store.currentUser(), repo.store.readiness()]);
      setUser(currentUser);
      if (!readiness.ready) { setNeedsSetup(true); setLoading(false); return; }
      setData(await repo.store.load());
      setNeedsSetup(false);
    } catch (caught) { setError(caught.message || 'Could not load portfolio data.'); }
    finally { setLoading(false); }
  }

  const managers = (repo.config.managerEmails || []).map((email) => email.toLowerCase());
  const isManager = repo.mode === 'demo' || managers.length === 0 || managers.includes(user?.email?.toLowerCase());
  const enriched = useMemo(() => data.projects.map((project) => enrichProject(project, data.tasks)), [data.projects, data.tasks]);
  const filteredProjects = useMemo(() => {
    const query = search.trim().toLowerCase();
    if (!query) return enriched;
    return enriched.filter((project) => [project.title, project.measurementArea, project.ownerName, project.nextMilestone, project.health].some((value) => String(value || '').toLowerCase().includes(query)));
  }, [enriched, search]);
  const openProject = enriched.find((project) => project.id === openProjectId);

  async function saveProject(project) {
    try {
      const saved = await repo.store.saveProject(project);
      setData((state) => ({ ...state, projects: state.projects.some((row) => row.id === saved.id) ? state.projects.map((row) => row.id === saved.id ? saved : row) : [...state.projects, saved] }));
      setOpenProjectId(saved.id); setProjectEditor(null); setToast({ message: `${saved.title} saved.` });
    } catch (caught) { setToast({ tone: 'bad', message: caught.message }); }
  }

  async function saveTask(task) {
    try {
      const normalized = { ...task, finishDate: task.status === 'Complete' ? (task.finishDate || todayIso()) : task.finishDate };
      const saved = await repo.store.saveTask(normalized);
      setData((state) => ({ ...state, tasks: state.tasks.map((row) => row.id === saved.id ? saved : row) }));
      setTaskEditor(null); setToast({ message: `WBS ${saved.wbs} updated.` });
    } catch (caught) { setToast({ tone: 'bad', message: caught.message }); }
  }

  async function addUpdate(project, summary) {
    try {
      const saved = await repo.store.saveUpdate({ id: uid('update'), projectKey: project.projectKey, type: 'Status', summary, entryDate: todayIso(), authorName: user.title, authorEmail: user.email });
      setData((state) => ({ ...state, updates: [saved, ...state.updates] })); setToast({ message: 'Project update posted.' });
    } catch (caught) { setToast({ tone: 'bad', message: caught.message }); }
  }

  async function addRisk(project, risk) {
    try {
      const saved = await repo.store.saveRisk({ id: uid('risk'), projectKey: project.projectKey, ...risk, probability: risk.probability || 'Possible', ownerName: user.title, status: 'Open', dueDate: '' });
      setData((state) => ({ ...state, risks: [saved, ...state.risks] })); setToast({ message: 'Risk added.' });
    } catch (caught) { setToast({ tone: 'bad', message: caught.message }); }
  }

  async function claimProject(project) { await saveProject({ ...project, ownerName: user.title, ownerEmail: user.email }); }

  async function importBaseline() {
    try {
      await repo.store.seed(seedData);
      setData(await repo.store.load()); setToast({ message: 'Workbook baseline imported.' });
    } catch (caught) { setToast({ tone: 'bad', message: caught.message }); }
  }

  function exportCsv() {
    const quote = (value) => `"${String(value ?? '').replace(/"/g, '""')}"`;
    const rows = [['Project', 'Measurement Area', 'Owner', 'Health', 'Priority', 'Stage', 'Percent Complete', 'Next Milestone', 'Next Milestone Date'], ...enriched.map((project) => [project.title, project.measurementArea, project.ownerName, project.health, project.priority, seedData.phases.find((phase) => phase.key === project.currentStageKey)?.name, project.percentComplete, project.nextMilestone, project.nextMilestoneDate])];
    const blob = new Blob([rows.map((row) => row.map(quote).join(',')).join('\n')], { type: 'text/csv;charset=utf-8' });
    const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = `modernization-portfolio-${todayIso()}.csv`; link.click(); URL.revokeObjectURL(link.href);
  }

  if (loading) return <Skeleton />;
  if (needsSetup) return <SetupGate repo={repo} onReady={initialize} />;
  if (error) return <div className="gate-shell"><div className="gate-card"><div className="inline-error"><Icon name="alert" />{error}</div><button className="button primary" onClick={initialize}>Try again</button></div></div>;

  return <div className="app-shell">
    <aside className="sidebar"><div className="brand"><div className="brand-mark">M</div><div><strong>MODERNIZE</strong><span>NPSL Project Tracker</span></div></div><nav>{NAV.filter(([key]) => isManager || ['my-work'].includes(key)).map(([key, label, icon]) => <button className={view === key ? 'active' : ''} key={key} onClick={() => setView(key)}><Icon name={icon} /><span>{label}</span>{key === 'my-work' && <Badge>{data.tasks.filter((task) => task.ownerEmail?.toLowerCase() === user?.email?.toLowerCase() && task.status !== 'Complete').length}</Badge>}</button>)}</nav><div className="sidebar-footer"><div className="sync-card"><span className={`sync-dot ${repo.mode}`} /><div><strong>{repo.mode === 'sharepoint' ? 'SharePoint connected' : 'Demo workspace'}</strong><span>{repo.mode === 'sharepoint' ? 'Changes save automatically' : 'Saved in this browser'}</span></div></div><div className="sidebar-user"><span className="avatar">{user?.title?.slice(0, 1) || 'U'}</span><div><strong>{user?.title || 'User'}</strong><span>{isManager ? 'Portfolio access' : 'Engineer access'}</span></div></div></div></aside>
    <div className="main-shell"><header className="topbar"><div className="mobile-brand"><div className="brand-mark">M</div><strong>MODERNIZE</strong></div><label className="search-box"><Icon name="search" /><input aria-label="Search projects" placeholder="Search projects, owners, milestones…" value={search} onChange={(event) => setSearch(event.target.value)} />{search && <button onClick={() => setSearch('')} aria-label="Clear search"><Icon name="close" size={14} /></button>}</label><div className="top-actions"><button className="icon-button" onClick={initialize} aria-label="Refresh"><Icon name="refresh" /></button><button className="icon-button" onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')} aria-label="Toggle theme"><Icon name={theme === 'dark' ? 'sun' : 'moon'} /></button><button className="button secondary export-button" onClick={exportCsv}><Icon name="download" /> Export</button>{isManager && <button className="button primary" onClick={() => setProjectEditor({})}><Icon name="plus" /> New project</button>}</div></header>
      {repo.mode === 'demo' && <div className="demo-banner"><Icon name="shield" /><span><strong>Demo workspace:</strong> edits stay in this browser. The Forge build switches to SharePoint automatically when hosted on your site.</span></div>}
      {data.projects.some((project) => project.importedBaseline) && <div className="baseline-banner"><Icon name="alert" /><span>The workbook is a historical baseline. Validate owners, health, and target dates before using it for current reporting.</span></div>}
      <main>
        {view === 'overview' && <Overview projects={filteredProjects} tasks={data.tasks} risks={data.risks} onOpen={(project) => setOpenProjectId(project.id)} phaseFilter={phaseFilter} setPhaseFilter={setPhaseFilter} />}
        {view === 'board' && <Board projects={filteredProjects} onOpen={(project) => setOpenProjectId(project.id)} />}
        {view === 'projects' && <ProjectsTable projects={filteredProjects} onOpen={(project) => setOpenProjectId(project.id)} />}
        {view === 'my-work' && <MyWork projects={filteredProjects} tasks={data.tasks} user={user} onOpen={(project) => setOpenProjectId(project.id)} onTask={setTaskEditor} />}
        {!data.projects.length && <EmptyState title="Portfolio is empty" message="Import the workbook baseline or add the first project." action={<div className="empty-actions"><button className="button primary" onClick={importBaseline}>Import workbook baseline</button><button className="button secondary" onClick={() => setProjectEditor({})}>Add project</button></div>} />}
      </main>
    </div>
    {openProject && <ProjectDrawer project={openProject} user={user} tasks={data.tasks} updates={data.updates} risks={data.risks} onClose={() => setOpenProjectId('')} onEdit={setProjectEditor} onSaveTask={saveTask} onAddUpdate={addUpdate} onAddRisk={addRisk} onClaim={claimProject} />}
    {projectEditor && <ProjectEditor project={projectEditor.id ? projectEditor : null} user={user} onClose={() => setProjectEditor(null)} onSave={saveProject} />}
    {taskEditor && <TaskEditor task={taskEditor} user={user} onClose={() => setTaskEditor(null)} onSave={saveTask} />}
    <Toast toast={toast} />
  </div>;
}
