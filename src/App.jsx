import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Icon } from './components/Icon';
import { createRepository, createStarterTasks, isOwnedByUser, workflowData, userIdentityKey } from './lib/repository';
import { downloadPortfolioWorkbook } from './lib/exportPortfolioWorkbook';
import navairSeal from './assets/navair-seal-384.webp';

const NAV = [
  ['overview', 'Portfolio', 'overview'],
  ['board', 'Pipeline board', 'board'],
  ['projects', 'All projects', 'projects'],
  ['my-work', 'My work', 'tasks'],
  ['glossary', 'Acronym glossary', 'book'],
];
const TASK_STATUSES = ['Not Started', 'In Progress', 'Blocked', 'Complete', 'Not Required'];
const CLOSED_TASK_STATUSES = ['Complete', 'Not Required', 'Not Applicable'];
const HEALTHS = ['On Track', 'Needs Review', 'At Risk', 'Blocked'];
const PRIORITIES = ['Low', 'Medium', 'High', 'Critical'];
const todayIso = () => new Date().toISOString().slice(0, 10);
const titleCase = (value) => String(value || '').replace(/(^|[-_])([a-z])/g, (_, space, char) => `${space ? ' ' : ''}${char.toUpperCase()}`);
const displayDate = (value) => value ? new Intl.DateTimeFormat(undefined, { month: 'short', day: 'numeric', year: 'numeric' }).format(new Date(`${String(value).slice(0, 10)}T12:00:00`)) : 'Not set';
const compactDate = (value) => value ? new Intl.DateTimeFormat(undefined, { month: 'short', day: 'numeric' }).format(new Date(`${String(value).slice(0, 10)}T12:00:00`)) : '—';
const uid = (prefix) => `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
export const userInitials = (user) => {
  const name = String(user?.title || '').replace(/\([^)]*\)/g, '').trim();
  if (name) {
    if (name.includes(',')) {
      const [lastName, givenNames = ''] = name.split(',', 2).map((part) => part.trim());
      return `${lastName[0] || ''}${givenNames[0] || lastName[1] || ''}`.toUpperCase();
    }
    const parts = name.split(/\s+/).filter(Boolean);
    return `${parts[0]?.[0] || ''}${parts.length > 1 ? parts[parts.length - 1][0] : parts[0]?.[1] || ''}`.toUpperCase();
  }
  return String(user?.email || user?.loginName || 'U').slice(0, 2).toUpperCase();
};

function phaseIndex(key) { return Math.max(0, workflowData.phases.findIndex((phase) => phase.key === key)); }
function isClosedTask(task) { return CLOSED_TASK_STATUSES.includes(task.status); }
function isOverdue(task) { return !isClosedTask(task) && task.dueDate && task.dueDate < todayIso(); }
function statusTone(value) {
  if (/complete|not required|not applicable|on track/i.test(value)) return 'good';
  if (/blocked|critical/i.test(value)) return 'bad';
  if (/risk|high|review|progress/i.test(value)) return 'warn';
  return 'neutral';
}

function enrichProject(project, allTasks) {
  const tasks = allTasks.filter((task) => task.projectKey === project.projectKey);
  if (!tasks.length) return { ...project, tasks, blockedCount: 0, overdueCount: 0 };
  const completed = tasks.filter(isClosedTask).length;
  const orderedTasks = [...tasks].sort((a, b) => phaseIndex(a.phaseKey) - phaseIndex(b.phaseKey) || a.order - b.order);
  const nextTask = orderedTasks.find((task) => !isClosedTask(task));
  const finalPhase = [...workflowData.phases].reverse().find((phase) => tasks.some((task) => task.phaseKey === phase.key && isClosedTask(task)));
  return {
    ...project,
    tasks,
    percentComplete: Math.round((completed / tasks.length) * 100),
    currentStageKey: nextTask?.phaseKey || finalPhase?.key || project.currentStageKey,
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

function KpiCard({ icon, label, value, detail, tone = 'blue' }) {
  return <article className={`kpi-card kpi-${tone}`}><div className="kpi-icon"><Icon name={icon} /></div><div><span>{label}</span><strong>{value}</strong><small>{detail}</small></div></article>;
}

function PipelineRail({ projects, selected, onSelect }) {
  const max = Math.max(1, ...workflowData.phases.map((phase) => projects.filter((project) => project.currentStageKey === phase.key).length));
  return <div className="pipeline-rail">
    {workflowData.phases.map((phase, index) => {
      const count = projects.filter((project) => project.currentStageKey === phase.key).length;
      return <button key={phase.key} className={`phase-node ${selected === phase.key ? 'selected' : ''}`} onClick={() => onSelect?.(selected === phase.key ? '' : phase.key)}>
        <div className="phase-top"><span className="phase-index">{String(index + 1).padStart(2, '0')}</span><strong>{count}</strong></div>
        <span className="phase-label">{phase.name}</span>
        <div className="phase-bar"><span style={{ width: `${count ? Math.max(8, (count / max) * 100) : 0}%` }} /></div>
      </button>;
    })}
  </div>;
}

function ProjectCard({ project, onOpen, onDelete }) {
  const [menuOpen, setMenuOpen] = useState(false);
  const phase = workflowData.phases.find((item) => item.key === project.currentStageKey);
  const openProject = () => onOpen(project);
  return <article className="project-card" role="button" tabIndex="0" onClick={openProject} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); openProject(); } }}>
    <div className="project-card-top"><Badge tone={statusTone(project.health)} dot>{project.health}</Badge><div className="project-menu-wrap"><button className="project-menu" aria-label={`Project actions for ${project.title}`} aria-expanded={menuOpen} onClick={(event) => { event.stopPropagation(); setMenuOpen((open) => !open); }}><Icon name="more" size={16} /></button>{menuOpen && <div className="project-menu-popover"><button onClick={(event) => { event.stopPropagation(); setMenuOpen(false); onDelete(project); }}>Delete project</button></div>}</div></div>
    <div><span className="eyebrow">{project.measurementArea}</span><h3>{project.title}</h3></div>
    <div className="project-owner"><span className="avatar small">{project.ownerName === 'Unassigned' ? '?' : project.ownerName.slice(0, 1)}</span><span>{project.ownerName}</span></div>
    <div className="project-phase"><span>Stage {phaseIndex(project.currentStageKey) + 1} of {workflowData.phases.length}</span><strong>{phase?.short}</strong></div>
    <Progress value={project.percentComplete} compact />
    <div className="project-card-bottom"><span>{project.percentComplete}% complete</span><span title={`Target completion: ${displayDate(project.targetFinish)}`}><Icon name="calendar" size={14} /> {project.targetFinish ? compactDate(project.targetFinish) : 'No target date'}</span></div>
  </article>;
}

function Overview({ projects, tasks, risks, onOpen, onDelete, phaseFilter, setPhaseFilter }) {
  const active = projects.filter((project) => project.status !== 'Complete').length;
  const completedTasks = tasks.filter(isClosedTask).length;
  const attention = projects.filter((project) => ['At Risk', 'Blocked', 'Needs Review'].includes(project.health) || project.blockedCount).length;
  const avg = projects.length ? Math.round(projects.reduce((sum, project) => sum + project.percentComplete, 0) / projects.length) : 0;
  const filtered = phaseFilter ? projects.filter((project) => project.currentStageKey === phaseFilter) : projects;
  const attentionProjects = [...projects].sort((a, b) => (b.blockedCount + b.overdueCount + (b.health === 'At Risk' ? 2 : 0)) - (a.blockedCount + a.overdueCount + (a.health === 'At Risk' ? 2 : 0))).slice(0, 5);

  return <div className="page-stack">
    <section className="hero-row"><div><span className="section-kicker">Portfolio command center</span><h1>Modernization at a glance</h1><p>One view of every measurement area, milestone, and handoff across the modernization pipeline.</p></div></section>
    <section className="kpi-grid">
      <KpiCard icon="projects" label="Active projects" value={active} detail={`${projects.length} total measurement areas`} tone="blue" />
      <KpiCard icon="trend" label="Portfolio progress" value={`${avg}%`} detail={`${completedTasks} of ${tasks.length} tasks resolved`} tone="mint" />
      <KpiCard icon="alert" label="Needs attention" value={attention} detail={`${risks.filter((risk) => risk.status !== 'Closed').length} open risks`} tone="amber" />
      <KpiCard icon="clock" label="Blocked tasks" value={tasks.filter((task) => task.status === 'Blocked').length} detail="Across all project teams" tone="rose" />
    </section>
    <section className="panel pipeline-panel"><div className="panel-heading"><div><span className="section-kicker">Portfolio flow</span><h2>Projects by pipeline stage</h2></div>{phaseFilter && <button className="text-button" onClick={() => setPhaseFilter('')}>Clear filter</button>}</div><PipelineRail projects={projects} selected={phaseFilter} onSelect={setPhaseFilter} /></section>
    <section className="dashboard-grid">
      <div className="panel project-showcase"><div className="panel-heading"><div><span className="section-kicker">In motion</span><h2>{phaseFilter ? `${workflowData.phases.find((p) => p.key === phaseFilter)?.name} projects` : 'Portfolio projects'}</h2></div><span className="count-label">{filtered.length} projects</span></div>
        <div className="project-card-grid">{filtered.slice(0, 6).map((project) => <ProjectCard key={project.id} project={project} onOpen={onOpen} onDelete={onDelete} />)}</div>
      </div>
      <div className="panel attention-panel"><div className="panel-heading"><div><span className="section-kicker">Focus queue</span><h2>Attention needed</h2></div></div>
        <div className="attention-list">{attentionProjects.map((project) => <button key={project.id} onClick={() => onOpen(project)}><span className={`health-pip ${statusTone(project.health)}`} /><div><strong>{project.title}</strong><span className="attention-owner">Owner: {project.ownerName || 'Unassigned'}</span><small>{project.nextMilestone}</small></div><div className="attention-meta"><Badge tone={statusTone(project.health)}>{project.health}</Badge><Icon name="chevron" size={16} /></div></button>)}</div>
      </div>
    </section>
  </div>;
}

function Board({ projects, onOpen, onDelete }) {
  return <div className="page-stack"><section className="page-heading"><div><span className="section-kicker">End-to-end flow</span><h1>Pipeline board</h1><p>Scan where every project sits and open a card to update its work breakdown.</p></div></section>
    <div className="board-scroll"><div className="kanban-board">{workflowData.phases.map((phase, index) => {
      const rows = projects.filter((project) => project.currentStageKey === phase.key);
      return <section className="kanban-column" key={phase.key}><header><div><span>{String(index + 1).padStart(2, '0')}</span><strong>{phase.name}</strong></div><Badge>{rows.length}</Badge></header><div className="kanban-list">{rows.map((project) => <ProjectCard key={project.id} project={project} onOpen={onOpen} onDelete={onDelete} />)}{!rows.length && <div className="kanban-empty">No projects in this stage</div>}</div></section>;
    })}</div></div>
  </div>;
}

function ProjectsTable({ projects, onOpen }) {
  const [health, setHealth] = useState('');
  const rows = health ? projects.filter((project) => project.health === health) : projects;
  return <div className="page-stack"><section className="page-heading inline"><div><span className="section-kicker">Portfolio register</span><h1>All projects</h1><p>Detailed ownership, health, stage, and milestone visibility.</p></div><label className="select-wrap"><Icon name="filter" /><select value={health} onChange={(event) => setHealth(event.target.value)}><option value="">All health states</option>{HEALTHS.map((item) => <option key={item}>{item}</option>)}</select></label></section>
    <section className="panel table-panel"><div className="project-table" role="table"><div className="project-table-row table-header" role="row"><span>Project</span><span>Owner</span><span>Stage</span><span>Health</span><span>Progress</span><span>Next milestone</span><span /></div>{rows.map((project) => <button className="project-table-row" role="row" key={project.id} onClick={() => onOpen(project)}><span><strong>{project.title}</strong><small>{project.measurementArea}</small></span><span><span className="avatar tiny">{project.ownerName === 'Unassigned' ? '?' : project.ownerName[0]}</span>{project.ownerName}</span><span>{workflowData.phases.find((phase) => phase.key === project.currentStageKey)?.short}</span><span><Badge tone={statusTone(project.health)} dot>{project.health}</Badge></span><span><strong>{project.percentComplete}%</strong><Progress value={project.percentComplete} compact /></span><span><strong>{project.nextMilestone}</strong><small>{displayDate(project.nextMilestoneDate)}</small></span><span><Icon name="chevron" /></span></button>)}</div></section>
  </div>;
}

function MyWork({ projects, tasks, user, onOpen, onTask, onDelete }) {
  const assigned = projects.filter((project) => isOwnedByUser(project, user));
  const myTasks = tasks.filter((task) => isOwnedByUser(task, user) && !isClosedTask(task)).sort((a, b) => (a.dueDate || '9999').localeCompare(b.dueDate || '9999'));
  return <div className="page-stack"><section className="page-heading"><div><span className="section-kicker">Personal workspace</span><h1>My work</h1><p>Projects and actions assigned to {user?.title || user?.email || 'you'}.</p></div></section>
    <section className="dashboard-grid my-work-grid"><div className="panel"><div className="panel-heading"><div><span className="section-kicker">Ownership</span><h2>My projects</h2></div><Badge>{assigned.length}</Badge></div>{assigned.length ? <div className="project-card-grid one-column">{assigned.map((project) => <ProjectCard key={project.id} project={project} onOpen={onOpen} onDelete={onDelete} />)}</div> : <EmptyState title="No assigned projects" message="Open an unassigned project and choose Claim project to make it yours." />}</div>
    <div className="panel"><div className="panel-heading"><div><span className="section-kicker">Next actions</span><h2>My task queue</h2></div><Badge>{myTasks.length}</Badge></div><div className="task-queue">{myTasks.slice(0, 14).map((task) => <button key={task.id} onClick={() => onTask(task)}><span className={`task-state ${statusTone(task.status)}`}><Icon name={task.status === 'Blocked' ? 'alert' : 'task'} size={16} /></span><div><strong>{task.title}</strong><span>{projects.find((project) => project.projectKey === task.projectKey)?.title} · WBS {task.wbs}</span></div><div><Badge tone={isOverdue(task) ? 'bad' : 'neutral'}>{task.dueDate ? compactDate(task.dueDate) : 'No due date'}</Badge></div></button>)}{!myTasks.length && <EmptyState title="Queue clear" message="No open tasks are assigned to your signed-in SharePoint identity." />}</div></div></section>
  </div>;
}

function Glossary({ acronyms, onAdd, onDelete }) {
  const [query, setQuery] = useState('');
  const [draft, setDraft] = useState({ acronym: '', term: '', definition: '' });
  const normalizedQuery = query.trim().toLowerCase();
  const rows = [...acronyms]
    .filter((entry) => [entry.acronym, entry.term, entry.definition].some((value) => String(value || '').toLowerCase().includes(normalizedQuery)))
    .sort((a, b) => a.acronym.localeCompare(b.acronym));
  const set = (key) => (event) => setDraft((entry) => ({ ...entry, [key]: event.target.value }));
  const submit = async (event) => {
    event.preventDefault();
    if (!draft.acronym.trim() || !draft.term.trim()) return;
    const saved = await onAdd(draft);
    if (saved) setDraft({ acronym: '', term: '', definition: '' });
  };
  return <div className="page-stack"><section className="page-heading inline"><div><span className="section-kicker">Modernization reference</span><h1>Acronym glossary</h1><p>A shared SharePoint reference that everyone with access can maintain.</p></div><label className="select-wrap glossary-search"><Icon name="search" /><input aria-label="Search acronym glossary" placeholder="Search terms or definitions…" value={query} onChange={(event) => setQuery(event.target.value)} /></label></section>
    <form className="panel glossary-composer" onSubmit={submit}><label><span>Acronym</span><input aria-label="Acronym" placeholder="e.g., CSS" value={draft.acronym} onChange={set('acronym')} /></label><label><span>Full term</span><input aria-label="Full term" placeholder="Expanded term" value={draft.term} onChange={set('term')} /></label><label className="glossary-definition"><span>Definition</span><input aria-label="Definition" placeholder="Optional plain-language definition" value={draft.definition} onChange={set('definition')} /></label><button className="button primary" type="submit" disabled={!draft.acronym.trim() || !draft.term.trim()}><Icon name="plus" /> Add acronym</button></form>
    <section className="panel glossary-panel"><div className="glossary-table" role="table"><div className="glossary-row glossary-header" role="row"><span>Acronym</span><span>Full term</span><span>Definition</span><span>Actions</span></div>{rows.map((entry) => <div className="glossary-row" role="row" key={entry.id}><strong>{entry.acronym}</strong><span>{entry.term}</span><p>{entry.definition || '—'}</p><button className="glossary-delete" type="button" aria-label={`Remove ${entry.acronym}`} onClick={() => onDelete(entry)}><Icon name="trash" size={16} /></button></div>)}</div>{!rows.length && <EmptyState title={acronyms.length ? 'No matching acronym' : 'No acronyms yet'} message={acronyms.length ? 'Try a shorter acronym, term, or keyword.' : 'Add the first acronym using the fields above.'} />}</section>
  </div>;
}

function Field({ label, children, wide = false }) { return <label className={`field ${wide ? 'field-wide' : ''}`}><span>{label}</span>{children}</label>; }

function Modal({ title, subtitle, onClose, children, actions, wide = false }) {
  useEffect(() => { const listener = (event) => event.key === 'Escape' && onClose(); window.addEventListener('keydown', listener); return () => window.removeEventListener('keydown', listener); }, [onClose]);
  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className={`modal ${wide ? 'modal-wide' : ''}`} role="dialog" aria-modal="true"><header><div><h2>{title}</h2>{subtitle && <p>{subtitle}</p>}</div><button className="icon-button" onClick={onClose} aria-label="Close"><Icon name="close" /></button></header><div className="modal-body">{children}</div>{actions && <footer>{actions}</footer>}</section></div>;
}

function ProjectEditor({ project, user, onClose, onSave }) {
  const [draft, setDraft] = useState({ title: '', measurementArea: '', description: '', ownerName: 'Unassigned', ownerEmail: '', ownerKey: '', priority: 'Medium', health: 'Needs Review', status: 'Planned', currentStageKey: 'requirement', targetFinish: '', ...project });
  const set = (key) => (event) => setDraft((row) => ({ ...row, [key]: event.target.value }));
  function claim() { setDraft((row) => ({ ...row, ownerName: user.title, ownerEmail: user.email, ownerKey: userIdentityKey(user) })); }
  function setOwnerIdentity(event) {
    const entered = event.target.value.trim();
    const isLoginKey = !!entered && (entered.includes('|') || !entered.includes('@'));
    const ownerKey = isLoginKey ? entered : '';
    const ownerEmail = entered.includes('@') ? entered.split('|').pop() : '';
    setDraft((row) => ({ ...row, ownerKey, ownerEmail }));
  }
  return <Modal title={project?.id ? 'Edit project' : 'Add modernization project'} subtitle="Update the portfolio record and ownership." onClose={onClose} actions={<><button className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" disabled={!draft.title.trim()} onClick={() => onSave({ ...draft, id: draft.id || uid('project'), projectKey: draft.projectKey || draft.title.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, ''), percentComplete: draft.percentComplete || 0, nextMilestone: draft.nextMilestone || 'Define project plan', nextMilestoneDate: draft.nextMilestoneDate || '' })}>Save project</button></>}>
    <div className="field-grid"><Field label="Project name"><input value={draft.title} onChange={set('title')} /></Field><Field label="Measurement area"><input value={draft.measurementArea} onChange={set('measurementArea')} /></Field><Field label="Project owner"><input value={draft.ownerName || ''} placeholder="Unassigned" onChange={set('ownerName')} /></Field><Field label="Owner email or SharePoint login"><div className="input-action"><input value={draft.ownerEmail || draft.ownerKey || ''} placeholder="Unassigned" onChange={setOwnerIdentity} /><button type="button" onClick={claim}>Assign me</button></div></Field><Field label="Priority"><select value={draft.priority} onChange={set('priority')}>{PRIORITIES.map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Health"><select value={draft.health} onChange={set('health')}>{HEALTHS.map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Status"><select value={draft.status} onChange={set('status')}>{['Planned', 'In Progress', 'On Hold', 'Complete', 'Cancelled'].map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Current stage"><select value={draft.currentStageKey} onChange={set('currentStageKey')}>{workflowData.phases.map((phase) => <option value={phase.key} key={phase.key}>{phase.name}</option>)}</select></Field><Field label="Target completion"><input type="date" value={draft.targetFinish || ''} onChange={set('targetFinish')} /></Field><Field label="Description" wide><textarea rows="4" value={draft.description} onChange={set('description')} /></Field></div>
  </Modal>;
}

function TaskEditor({ task, user, onClose, onSave }) {
  const [draft, setDraft] = useState(task);
  const set = (key) => (event) => setDraft((row) => ({ ...row, [key]: event.target.value }));
  function claim() { setDraft((row) => ({ ...row, ownerName: user.title, ownerEmail: user.email, ownerKey: userIdentityKey(user) })); }
  function setOwnerIdentity(event) {
    const entered = event.target.value.trim();
    const isLoginKey = !!entered && (entered.includes('|') || !entered.includes('@'));
    const ownerKey = isLoginKey ? entered : '';
    const ownerEmail = entered.includes('@') ? entered.split('|').pop() : '';
    setDraft((row) => ({ ...row, ownerKey, ownerEmail }));
  }
  const status = draft.status === 'Not Applicable' ? 'Not Required' : draft.status;
  return <Modal title="Update task" subtitle={`${workflowData.phases.find((phase) => phase.key === task.phaseKey)?.name || 'Modernization work breakdown'}`} onClose={onClose} actions={<><button className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" disabled={!draft.wbs?.trim() || !draft.title?.trim()} onClick={() => onSave({ ...draft, status })}>Save task</button></>}>
    <div className="field-grid"><Field label="WBS code"><input aria-label="WBS code" value={draft.wbs || ''} onChange={set('wbs')} /></Field><Field label="Pipeline stage"><select aria-label="Pipeline stage" value={draft.phaseKey} onChange={set('phaseKey')}>{workflowData.phases.map((phase) => <option value={phase.key} key={phase.key}>{phase.name}</option>)}</select></Field><Field label="Task name" wide><input aria-label="Task name" value={draft.title} onChange={set('title')} /></Field><Field label="Status"><select value={status} onChange={set('status')}>{TASK_STATUSES.map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Due date"><input type="date" value={status === 'Not Required' ? '' : (draft.dueDate || '')} disabled={status === 'Not Required'} title={status === 'Not Required' ? 'A due date is not needed for a task marked Not Required.' : ''} onChange={set('dueDate')} />{status === 'Not Required' && <small className="field-hint">No due date is needed.</small>}</Field><Field label="Task owner"><input value={draft.ownerName || ''} placeholder="Unassigned" onChange={set('ownerName')} /></Field><Field label="Owner email or SharePoint login"><div className="input-action"><input value={draft.ownerEmail || draft.ownerKey || ''} placeholder="Unassigned" onChange={setOwnerIdentity} /><button type="button" onClick={claim}>Assign me</button></div></Field><Field label="Notes" wide><textarea rows="4" value={draft.notes || ''} onChange={set('notes')} /></Field>{status === 'Blocked' && <Field label="Blocked reason" wide><textarea rows="3" value={draft.blockedReason || ''} onChange={set('blockedReason')} /></Field>}</div>
  </Modal>;
}

function ProjectDrawer({ project, user, tasks, updates, risks, onClose, onEdit, onSaveTask, onToggleTask, onSetPhaseRequired, onAddUpdate, onAddRisk, onClaim }) {
  const [tab, setTab] = useState('overview');
  const [taskEditor, setTaskEditor] = useState(null);
  const [updateText, setUpdateText] = useState('');
  const [riskDraft, setRiskDraft] = useState({ title: '', severity: 'Medium', probability: 'Possible', mitigation: '' });
  const projectTasks = tasks.filter((task) => task.projectKey === project.projectKey).sort((a, b) => a.order - b.order);
  const projectUpdates = updates.filter((item) => item.projectKey === project.projectKey).sort((a, b) => String(b.entryDate).localeCompare(String(a.entryDate)));
  const projectRisks = risks.filter((item) => item.projectKey === project.projectKey);
  const phase = workflowData.phases.find((item) => item.key === project.currentStageKey);
  return <div className="drawer-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><aside className="project-drawer" role="dialog" aria-modal="true"><header className="drawer-header"><div className="drawer-title-row"><div><span className="section-kicker">{project.measurementArea}</span><h2>{project.title}</h2><div className="drawer-badges"><Badge tone={statusTone(project.health)} dot>{project.health}</Badge><Badge>{project.priority} priority</Badge><Badge>{project.status}</Badge></div></div><button className="icon-button" onClick={onClose} aria-label="Close"><Icon name="close" /></button></div>
    <div className="drawer-actions"><button className="button secondary" onClick={() => onEdit(project)}>Edit project</button>{!project.ownerKey && <button className="button primary" onClick={() => onClaim(project)}>Claim project</button>}</div>
    <div className="drawer-progress"><div><span>{project.percentComplete}% complete</span><strong>{phase?.name}</strong></div><Progress value={project.percentComplete} /></div>
    <div className="mini-pipeline">{workflowData.phases.map((item, index) => <span key={item.key} className={index < phaseIndex(project.currentStageKey) ? 'done' : item.key === project.currentStageKey ? 'current' : ''} title={item.name}>{index + 1}</span>)}</div>
    <nav className="drawer-tabs">{[['overview', 'Overview'], ['tasks', `Work breakdown (${projectTasks.length})`], ['updates', `Updates (${projectUpdates.length})`], ['risks', `Risks (${projectRisks.length})`]].map(([key, label]) => <button className={tab === key ? 'active' : ''} key={key} onClick={() => setTab(key)}>{label}</button>)}</nav></header>
    <div className="drawer-content">
      {tab === 'overview' && <div className="drawer-section-stack"><section className="detail-grid"><div><span>Owner</span><strong>{project.ownerName}</strong><small>{project.ownerEmail || 'No SharePoint identity assigned'}</small></div><div><span>Target completion</span><strong>{displayDate(project.targetFinish)}</strong><small>{project.status}</small></div><div><span>Next milestone</span><strong>{project.nextMilestone}</strong><small>{displayDate(project.nextMilestoneDate)}</small></div><div><span>Open actions</span><strong>{projectTasks.filter((task) => !isClosedTask(task)).length}</strong><small>{project.blockedCount} blocked</small></div></section><section className="detail-block"><h3>Purpose</h3><p>{project.description || 'No project description has been added.'}</p></section><section className="detail-block"><h3>Upcoming work</h3><div className="upcoming-list">{projectTasks.filter((task) => !isClosedTask(task)).slice(0, 5).map((task) => <button key={task.id} onClick={() => setTaskEditor(task)}><span>{task.wbs}</span><div><strong>{task.title}</strong><small>{workflowData.phases.find((item) => item.key === task.phaseKey)?.short}</small></div><Badge tone={statusTone(task.status)}>{task.status}</Badge></button>)}</div></section></div>}
      {tab === 'tasks' && <div className="wbs-list">{workflowData.phases.map((stage) => { const stageTasks = projectTasks.filter((task) => task.phaseKey === stage.key); const done = stageTasks.filter(isClosedTask).length; const allNotRequired = stageTasks.length > 0 && stageTasks.every((task) => ['Not Required', 'Not Applicable'].includes(task.status)); return <section key={stage.key} className={`wbs-stage ${allNotRequired ? 'phase-not-required' : ''}`}><header><div><span>{String(phaseIndex(stage.key) + 1).padStart(2, '0')}</span><strong>{stage.name}</strong></div><div className="phase-header-actions"><Badge>{done}/{stageTasks.length}</Badge><button type="button" className={`phase-required-button ${allNotRequired ? 'active' : ''}`} disabled={!stageTasks.length} aria-label={`${allNotRequired ? 'Restore' : 'Mark'} ${stage.name} ${allNotRequired ? '' : 'not required'}`.trim()} onClick={() => onSetPhaseRequired(project, stage.key, !allNotRequired)}><Icon name={allNotRequired ? 'refresh' : 'ban'} size={13} />{allNotRequired ? 'Restore phase' : 'Not required'}</button></div></header><div>{stageTasks.map((task) => { const resolved = isClosedTask(task); return <div key={task.id} className="wbs-task" role="button" tabIndex="0" onClick={() => setTaskEditor(task)} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); setTaskEditor(task); } }}><button className={`task-check ${resolved ? 'checked' : ''}`} aria-label={`${resolved ? 'Reopen' : 'Complete'} WBS ${task.wbs}`} onClick={(event) => { event.stopPropagation(); onToggleTask(task); }}>{resolved && <Icon name="check" size={13} />}</button><span className="wbs-code">{task.wbs}</span><div><strong>{task.title}</strong>{task.dataIssue && <small className="issue-text"><Icon name="alert" size={12} />{task.dataIssue}</small>}</div><Badge tone={isOverdue(task) ? 'bad' : statusTone(task.status)}>{task.status === 'Not Applicable' ? 'Not Required' : task.status}</Badge><span className="task-date">{isClosedTask(task) ? 'Resolved' : compactDate(task.dueDate)}</span></div>; })}</div></section>; })}</div>}
      {tab === 'updates' && <div className="drawer-section-stack"><section className="composer"><textarea rows="3" placeholder="Add a concise status update, decision, or handoff…" value={updateText} onChange={(event) => setUpdateText(event.target.value)} /><div><span>Visible to everyone with access to this SharePoint workspace</span><button className="button primary small-button" disabled={!updateText.trim()} onClick={() => { onAddUpdate(project, updateText); setUpdateText(''); }}>Post update</button></div></section><section className="timeline-list">{projectUpdates.map((item) => <article key={item.id}><span className="timeline-dot" /><div><header><strong>{item.authorName}</strong><span>{displayDate(item.entryDate)}</span></header><Badge>{item.type}</Badge><p>{item.summary}</p></div></article>)}{!projectUpdates.length && <EmptyState title="No updates yet" message="Post the first status update to establish the project history." />}</section></div>}
      {tab === 'risks' && <div className="drawer-section-stack"><section className="risk-composer"><div className="field-grid"><Field label="Risk or issue"><input value={riskDraft.title} onChange={(event) => setRiskDraft((row) => ({ ...row, title: event.target.value }))} /></Field><Field label="Severity"><select value={riskDraft.severity} onChange={(event) => setRiskDraft((row) => ({ ...row, severity: event.target.value }))}>{['Low', 'Medium', 'High', 'Critical'].map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Mitigation" wide><textarea rows="3" value={riskDraft.mitigation} onChange={(event) => setRiskDraft((row) => ({ ...row, mitigation: event.target.value }))} /></Field></div><button className="button primary small-button" disabled={!riskDraft.title.trim()} onClick={() => { onAddRisk(project, riskDraft); setRiskDraft({ title: '', severity: 'Medium', probability: 'Possible', mitigation: '' }); }}>Add risk</button></section><section className="risk-list">{projectRisks.map((risk) => <article key={risk.id}><div><Badge tone={statusTone(risk.severity)}>{risk.severity}</Badge><Badge>{risk.status}</Badge></div><h3>{risk.title}</h3><p>{risk.mitigation || 'No mitigation has been documented.'}</p><span>{risk.ownerName || 'Unassigned'}</span></article>)}{!projectRisks.length && <EmptyState title="No open risks" message="Capture risks and mitigation actions here as the project advances." />}</section></div>}
    </div>
    {taskEditor && <TaskEditor task={taskEditor} user={user} onClose={() => setTaskEditor(null)} onSave={(task) => { onSaveTask(task); setTaskEditor(null); }} />}
  </aside></div>;
}

function Toast({ toast, onClose }) { if (!toast) return null; return <div className={`toast toast-${toast.tone || 'good'}`} role={toast.tone === 'bad' ? 'alert' : 'status'} aria-live={toast.tone === 'bad' ? 'assertive' : 'polite'}><Icon name={toast.tone === 'bad' ? 'alert' : toast.tone === 'info' ? 'clock' : 'check'} /><span>{toast.message}</span><button type="button" aria-label="Dismiss notification" onClick={onClose}><Icon name="close" size={13} /></button></div>; }

export function App() {
  const repoRef = useRef(null);
  if (!repoRef.current) repoRef.current = createRepository();
  const repo = repoRef.current;
  const [theme, setTheme] = useState(() => localStorage.getItem('mod-tracker-theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'));
  const [view, setView] = useState('overview');
  const [user, setUser] = useState(null);
  const [data, setData] = useState({ projects: [], tasks: [], updates: [], risks: [], acronyms: [] });
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [search, setSearch] = useState('');
  const [phaseFilter, setPhaseFilter] = useState('');
  const [openProjectId, setOpenProjectId] = useState('');
  const [projectEditor, setProjectEditor] = useState(null);
  const [taskEditor, setTaskEditor] = useState(null);
  const [toast, setToast] = useState(null);
  const [exporting, setExporting] = useState(false);

  useEffect(() => { document.documentElement.dataset.theme = theme; localStorage.setItem('mod-tracker-theme', theme); }, [theme]);
  useEffect(() => { initialize(); }, []);
  useEffect(() => { if (!toast) return; const timer = setTimeout(() => setToast(null), 3200); return () => clearTimeout(timer); }, [toast]);

  async function initialize() {
    setLoading(true); setError('');
    try {
      const [currentUser, readiness] = await Promise.all([repo.store.currentUser(), repo.store.readiness()]);
      setUser(currentUser);
      if (!readiness.ready) await repo.store.provision();
      setData(await repo.store.load());
    } catch (caught) { setError(caught.message || 'Could not load portfolio data.'); }
    finally { setLoading(false); }
  }

  const enriched = useMemo(() => data.projects.map((project) => enrichProject(project, data.tasks)), [data.projects, data.tasks]);
  const filteredProjects = useMemo(() => {
    const query = search.trim().toLowerCase();
    if (!query) return enriched;
    return enriched.filter((project) => [project.title, project.measurementArea, project.ownerName, project.nextMilestone, project.health].some((value) => String(value || '').toLowerCase().includes(query)));
  }, [enriched, search]);
  const openProject = enriched.find((project) => project.id === openProjectId);

  async function saveProject(project) {
    const source = data;
    const prepared = { ...project, ownerName: project.ownerName?.trim() || 'Unassigned' };
    const creating = !prepared.spId && !source.projects.some((row) => row.id === prepared.id);
    const template = creating ? createStarterTasks(prepared.projectKey, prepared) : [];
    const updateProjects = (state, saved, starterTasks = []) => ({
      ...state,
      projects: state.projects.some((row) => row.id === saved.id) ? state.projects.map((row) => row.id === saved.id ? saved : row) : [...state.projects, saved],
      tasks: starterTasks.length ? [...state.tasks.filter((row) => !starterTasks.some((task) => task.id === row.id)), ...starterTasks] : state.tasks,
    });

    setData((state) => updateProjects(state, prepared));
    setOpenProjectId(prepared.id);
    setProjectEditor(null);
    setToast({ tone: 'info', message: creating ? `${prepared.title} saved. Setting up its work breakdown…` : `Saving ${prepared.title}…` });
    try {
      const saved = await repo.store.saveProject(prepared);
      const starterTasks = [];
      if (creating) {
        for (let index = 0; index < template.length; index += 10) {
          starterTasks.push(...await Promise.all(template.slice(index, index + 10).map((task) => repo.store.saveTask(task))));
        }
      }
      setData((state) => updateProjects(state, saved, starterTasks));
      setToast({ message: `${saved.title} is ready.` });
    } catch (caught) {
      try { setData(await repo.store.load()); } catch { setData(source); }
      setOpenProjectId('');
      setToast({ tone: 'bad', message: caught.message || `${prepared.title} could not be saved.` });
    }
  }

  async function saveTask(task) {
    try {
      const status = task.status === 'Not Applicable' ? 'Not Required' : task.status;
      const resolved = CLOSED_TASK_STATUSES.includes(status);
      const normalized = { ...task, status, ownerName: task.ownerName?.trim() || 'Unassigned', dueDate: status === 'Not Required' ? '' : task.dueDate, finishDate: resolved ? (task.finishDate || todayIso()) : '' };
      const saved = await repo.store.saveTask(normalized);
      const updateState = (state) => ({ ...state, tasks: state.tasks.map((row) => row.id === saved.id ? saved : row) });
      setData(updateState);
      setTaskEditor(null); setToast({ message: `WBS ${saved.wbs} updated.` });
    } catch (caught) { setToast({ tone: 'bad', message: caught.message }); }
  }

  async function addUpdate(project, summary) {
    try {
      const update = { id: uid('update'), projectKey: project.projectKey, type: 'Status', summary, entryDate: todayIso(), authorName: user.title, authorEmail: user.email, authorKey: userIdentityKey(user) };
      const saved = await repo.store.saveUpdate(update);
      const updateState = (state) => ({ ...state, updates: [saved, ...state.updates] });
      setData(updateState);
      setToast({ message: 'Project update posted.' });
    } catch (caught) { setToast({ tone: 'bad', message: caught.message }); }
  }

  async function addRisk(project, risk) {
    try {
      const nextRisk = { id: uid('risk'), projectKey: project.projectKey, ...risk, probability: risk.probability || 'Possible', ownerName: user.title, ownerKey: userIdentityKey(user), status: 'Open', dueDate: '' };
      const saved = await repo.store.saveRisk(nextRisk);
      const updateState = (state) => ({ ...state, risks: [saved, ...state.risks] });
      setData(updateState);
      setToast({ message: 'Risk added.' });
    } catch (caught) { setToast({ tone: 'bad', message: caught.message }); }
  }

  async function claimProject(project) { await saveProject({ ...project, ownerName: user.title, ownerEmail: user.email, ownerKey: userIdentityKey(user) }); }

  async function toggleTaskComplete(task) {
    await saveTask({ ...task, status: isClosedTask(task) ? 'Not Started' : 'Complete' });
  }

  async function setPhaseRequired(project, phaseKey, notRequired) {
    const source = data;
    const affected = source.tasks.filter((task) => task.projectKey === project.projectKey && task.phaseKey === phaseKey);
    if (!affected.length) return;
    const changed = affected.map((task) => ({
      ...task,
      status: notRequired ? 'Not Required' : 'Not Started',
      dueDate: notRequired ? '' : task.dueDate,
      finishDate: notRequired ? (task.finishDate || todayIso()) : '',
      blockedReason: notRequired ? '' : task.blockedReason,
    }));
    const apply = (state, rows) => {
      const byId = new Map(rows.map((row) => [row.id, row]));
      return { ...state, tasks: state.tasks.map((task) => byId.get(task.id) || task) };
    };
    const phase = workflowData.phases.find((item) => item.key === phaseKey);
    setData((state) => apply(state, changed));
    setToast({ tone: 'info', message: `Updating ${phase?.short || 'phase'} tasks…` });
    try {
      const saved = repo.store.saveTasks ? await repo.store.saveTasks(changed) : await Promise.all(changed.map((task) => repo.store.saveTask(task)));
      setData((state) => apply(state, saved));
      setToast({ message: `${phase?.short || 'Phase'} ${notRequired ? 'marked not required' : 'restored'}.` });
    } catch (caught) {
      try { setData(await repo.store.load()); } catch { setData(source); }
      setToast({ tone: 'bad', message: caught.message || 'The phase could not be updated.' });
    }
  }

  async function deleteProject(project) {
    const source = data;
    const belongsToProject = (row) => row.projectKey === project.projectKey;
    const removeProject = (state) => ({
      ...state,
      projects: state.projects.filter((row) => row.id !== project.id),
      tasks: state.tasks.filter((row) => !belongsToProject(row)),
      updates: state.updates.filter((row) => !belongsToProject(row)),
      risks: state.risks.filter((row) => !belongsToProject(row)),
    });
    setData(removeProject);
    setOpenProjectId('');
    setToast({ tone: 'info', message: `Deleting ${project.title}…` });
    try {
      const related = [
        ...source.tasks.filter(belongsToProject).map((row) => ['tasks', row]),
        ...source.updates.filter(belongsToProject).map((row) => ['updates', row]),
        ...source.risks.filter(belongsToProject).map((row) => ['risks', row]),
      ];
      for (let index = 0; index < related.length; index += 10) {
        await Promise.all(related.slice(index, index + 10).map(([collection, row]) => repo.store.recycle(collection, row.spId, row.id)));
      }
      await repo.store.recycle('projects', project.spId, project.id);
      setToast({ message: `${project.title} deleted.` });
    } catch (caught) {
      try { setData(await repo.store.load()); } catch { setData(source); }
      setToast({ tone: 'bad', message: caught.message || 'Project could not be deleted.' });
    }
  }

  async function addAcronym(draft) {
    const acronym = draft.acronym.trim().toUpperCase();
    if (data.acronyms.some((entry) => entry.acronym.toUpperCase() === acronym)) {
      setToast({ tone: 'bad', message: `${acronym} is already in the glossary.` });
      return null;
    }
    const row = { id: uid('acronym'), acronym, term: draft.term.trim(), definition: draft.definition.trim() };
    try {
      const saved = await repo.store.saveAcronym(row);
      setData((state) => ({ ...state, acronyms: [...state.acronyms, saved] }));
      setToast({ message: `${acronym} added to the glossary.` });
      return saved;
    } catch (caught) {
      setToast({ tone: 'bad', message: caught.message || `${acronym} could not be added.` });
      return null;
    }
  }

  async function deleteAcronym(entry) {
    if (!window.confirm(`Remove ${entry.acronym} from the shared glossary?`)) return;
    const source = data;
    setData((state) => ({ ...state, acronyms: state.acronyms.filter((row) => row.id !== entry.id) }));
    setToast({ tone: 'info', message: `Removing ${entry.acronym}…` });
    try {
      await repo.store.recycle('acronyms', entry.spId, entry.id);
      setToast({ message: `${entry.acronym} removed.` });
    } catch (caught) {
      try { setData(await repo.store.load()); } catch { setData(source); }
      setToast({ tone: 'bad', message: caught.message || `${entry.acronym} could not be removed.` });
    }
  }

  async function exportWorkbook() {
    setExporting(true);
    try {
      await downloadPortfolioWorkbook({
        projects: enriched,
        tasks: data.tasks,
        updates: data.updates,
        risks: data.risks,
        phases: workflowData.phases,
        glossary: data.acronyms,
        user,
        sourceLabel: 'Live SharePoint portfolio',
      });
      setToast({ message: 'Portfolio Excel report downloaded.' });
    } catch (caught) {
      setToast({ tone: 'bad', message: caught.message || 'Excel export failed.' });
    } finally {
      setExporting(false);
    }
  }

  if (loading) return <Skeleton />;
  if (error) return <div className="gate-shell"><div className="gate-card"><div className="inline-error"><Icon name="alert" />{error}</div><button className="button primary" onClick={initialize}>Try again</button></div></div>;

  return <div className="app-shell">
    <aside className="sidebar"><div className="brand"><img className="brand-logo" src={navairSeal} alt="NAVAIR" /><div><strong>MODERNIZATION</strong><span>Project Tracker</span></div></div><nav>{NAV.map(([key, label, icon]) => <button className={view === key ? 'active' : ''} key={key} onClick={() => setView(key)}><Icon name={icon} /><span>{label}</span>{key === 'my-work' && <Badge>{enriched.filter((project) => isOwnedByUser(project, user)).length}</Badge>}</button>)}</nav><div className="sidebar-footer"><div className="sidebar-user"><span className="avatar">{userInitials(user)}</span><div><strong>{user?.title || 'SharePoint user'}</strong><span>{user?.email || user?.loginName || 'Full portfolio access'}</span></div></div></div></aside>
    <div className="main-shell"><header className="topbar"><div className="mobile-brand"><img className="brand-logo" src={navairSeal} alt="NAVAIR" /><strong>MODERNIZATION</strong></div><label className="search-box"><Icon name="search" /><input aria-label="Search projects" placeholder="Search projects, owners, milestones…" value={search} onChange={(event) => setSearch(event.target.value)} />{search && <button onClick={() => setSearch('')} aria-label="Clear search"><Icon name="close" size={14} /></button>}</label><div className="top-actions"><button className="icon-button" onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')} aria-label={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}><Icon name={theme === 'dark' ? 'sun' : 'moon'} /></button><button className="button secondary export-button" disabled={exporting} onClick={exportWorkbook}><Icon name="download" /> {exporting ? 'Building Excel…' : 'Export Excel'}</button><button className="button primary" onClick={() => setProjectEditor({})}><Icon name="plus" /> New project</button></div></header>
      <main>
        {view === 'overview' && <Overview projects={filteredProjects} tasks={data.tasks} risks={data.risks} onOpen={(project) => setOpenProjectId(project.id)} onDelete={deleteProject} phaseFilter={phaseFilter} setPhaseFilter={setPhaseFilter} />}
        {view === 'board' && <Board projects={filteredProjects} onOpen={(project) => setOpenProjectId(project.id)} onDelete={deleteProject} />}
        {view === 'projects' && <ProjectsTable projects={filteredProjects} onOpen={(project) => setOpenProjectId(project.id)} />}
        {view === 'my-work' && <MyWork projects={filteredProjects} tasks={data.tasks} user={user} onOpen={(project) => setOpenProjectId(project.id)} onTask={setTaskEditor} onDelete={deleteProject} />}
        {view === 'glossary' && <Glossary acronyms={data.acronyms} onAdd={addAcronym} onDelete={deleteAcronym} />}
        {!data.projects.length && view === 'overview' && <EmptyState title="Your modernization portfolio is ready" message="Add the first project to begin tracking modernization work." action={<div className="empty-actions"><button className="button primary" onClick={() => setProjectEditor({})}>Add first project</button></div>} />}
      </main>
    </div>
    {openProject && <ProjectDrawer project={openProject} user={user} tasks={data.tasks} updates={data.updates} risks={data.risks} onClose={() => setOpenProjectId('')} onEdit={setProjectEditor} onSaveTask={saveTask} onToggleTask={toggleTaskComplete} onSetPhaseRequired={setPhaseRequired} onAddUpdate={addUpdate} onAddRisk={addRisk} onClaim={claimProject} />}
    {projectEditor && <ProjectEditor project={projectEditor.id ? projectEditor : null} user={user} onClose={() => setProjectEditor(null)} onSave={saveProject} />}
    {taskEditor && <TaskEditor task={taskEditor} user={user} onClose={() => setTaskEditor(null)} onSave={saveTask} />}
    <Toast toast={toast} onClose={() => setToast(null)} />
  </div>;
}
