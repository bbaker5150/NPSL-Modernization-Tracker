import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Icon } from './components/Icon';
import { createRepository, createStarterTasks, isOwnedByUser, workflowData, userIdentityKey } from './lib/repository';
import { downloadPortfolioWorkbook } from './lib/exportPortfolioWorkbook';
import { authorizedStore, isManager, validateTask } from './lib/access';
import navairSeal from './assets/navair-seal-384.webp';

const NAV = [
  ['overview', 'Portfolio', 'overview'],
  ['board', 'Pipeline board', 'board'],
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

function KpiCard({ icon, label, value, detail, tone = 'blue', onClick }) {
  return <article className={`kpi-card kpi-${tone}`} role={onClick ? "button" : undefined} tabIndex={onClick ? 0 : undefined} onClick={onClick} onKeyDown={(event) => { if (onClick && ['Enter', ' '].includes(event.key)) { event.preventDefault(); onClick(); } }}><div className="kpi-icon"><Icon name={icon} /></div><div><span>{label}</span><strong>{value}</strong><small>{detail}</small></div></article>;
}

function PipelineRail({ projects, selected, onSelect }) {
  const max = Math.max(1, ...workflowData.phases.map((phase) => projects.filter((project) => project.currentStageKey === phase.key).length));
  return <div className="pipeline-rail">
    {workflowData.phases.map((phase, index) => {
      const count = projects.filter((project) => project.currentStageKey === phase.key).length;
      return <button key={phase.key} className={`phase-node ${selected.includes(phase.key) ? 'selected' : ''}`} aria-pressed={selected.includes(phase.key)} onClick={(event) => onSelect?.(event.ctrlKey || event.metaKey ? (selected.includes(phase.key) ? selected.filter((key) => key !== phase.key) : [...selected, phase.key]) : (selected.length === 1 && selected[0] === phase.key ? [] : [phase.key]))}>
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
    <div className="project-card-top"><Badge tone={statusTone(project.health)} dot>{project.health}</Badge><div className="project-menu-wrap" hidden={!onDelete}><button className="project-menu" aria-label={`Project actions for ${project.title}`} aria-expanded={menuOpen} onClick={(event) => { event.stopPropagation(); setMenuOpen((open) => !open); }}><Icon name="more" size={16} /></button>{menuOpen && onDelete && <div className="project-menu-popover"><button onClick={(event) => { event.stopPropagation(); setMenuOpen(false); onDelete(project); }}>Delete project</button></div>}</div></div>
    <div><span className="eyebrow">{project.measurementArea}</span><h3>{project.title}</h3></div>
    <div className="project-owner"><span className="avatar small">{project.ownerName === 'Unassigned' ? '?' : project.ownerName.slice(0, 1)}</span><span>{project.ownerName}</span></div>
    <div className="project-phase"><span>Stage {phaseIndex(project.currentStageKey) + 1} of {workflowData.phases.length}</span><strong>{phase?.short}</strong></div>
    <Progress value={project.percentComplete} compact />
    <div className="project-card-bottom"><span>{project.percentComplete}% complete</span><span title={`Target completion: ${displayDate(project.targetFinish)}`}><Icon name="calendar" size={14} /> {project.targetFinish ? compactDate(project.targetFinish) : 'No target date'}</span></div>
  </article>;
}

function Overview({ projects, tasks, onOpen, phaseFilter, setPhaseFilter, onAttention }) {
  const active = projects.filter((project) => project.status !== 'Complete');
  const avg = projects.length ? Math.round(projects.reduce((sum, project) => sum + (project.percentComplete || 0), 0) / projects.length) : 0;
  const filtered = phaseFilter.length ? projects.filter((project) => phaseFilter.includes(project.currentStageKey)) : projects;
  return <div className="page-stack">
    <section className="hero-row"><div><span className="section-kicker">Portfolio command center</span><h1>Modernization at a glance</h1><p>Measurement areas, milestones, and handoffs across the modernization pipeline.</p></div></section>
    <section className="kpi-grid">
      <KpiCard icon="projects" label="Active projects" value={active.length} detail={`${projects.length} total measurement areas`} />
      <KpiCard icon="trend" label="Portfolio progress" value={`${avg}%`} detail={`${tasks.filter(isClosedTask).length} of ${tasks.length} tasks resolved`} tone="mint" />
      <KpiCard icon="alert" label="Needs attention" value={active.length} detail="Review unfinished projects and task exceptions" tone="amber" onClick={onAttention} />
    </section>
    <section className="panel pipeline-panel"><div className="panel-heading"><h2>Projects by pipeline stage</h2>{phaseFilter.length > 0 && <button className="text-button" onClick={() => setPhaseFilter([])}>Clear filter</button>}</div><p className="pipeline-hint">Click a stage to filter. Ctrl-click (or ⌘-click) to select multiple stages.</p><PipelineRail projects={projects} selected={phaseFilter} onSelect={setPhaseFilter} /></section>
    {projects.length > 0 && <ProjectsTable projects={filtered} onOpen={onOpen} />}
  </div>;
}

function Attention({ projects, onOpen }) {
  const rows = projects.filter((project) => project.status !== 'Complete');
  return <div className="page-stack"><section className="page-heading"><h1>Needs attention</h1><p>All unfinished projects, pending tasks, and documented exceptions.</p></section>{rows.map((project) => <section className="panel attention-project" key={project.id}>
    <button className="text-button" onClick={() => onOpen(project)}><h2>{project.title}</h2></button><p>Owner: {project.ownerName || 'Unassigned'}</p>
    {project.tasks.filter((task) => !isClosedTask(task) || task.deferredDate || ['Not Required', 'Not Applicable'].includes(task.status)).map((task) => <article className="attention-task" key={task.id}>
      <div><strong>{task.title}</strong><p>Task owner: {task.ownerName || project.ownerName || 'Unassigned'}</p></div>
      <Badge tone={statusTone(task.status)}>{task.status}</Badge> {isOverdue(task) && <Badge tone="bad">Overdue · {displayDate(task.dueDate)}</Badge>}
      {task.deferredDate && <div><Badge tone="warn">Deferred · {displayDate(task.deferredDate)}</Badge><p>{task.deferredJustification}</p></div>}
      {['Not Required', 'Not Applicable'].includes(task.status) && <div><Badge tone="warn">Not required</Badge><p>{task.notRequiredJustification || 'Legacy task — justification not recorded'}</p></div>}
    </article>)}{!project.tasks.some((task) => !isClosedTask(task)) && <p>No pending tasks. Review project closeout.</p>}
  </section>)}{!rows.length && <EmptyState title="No unfinished projects" message="All projects have been closed out." />}</div>;
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
  const [search, setSearch] = useState('');
  const query = search.trim().toLowerCase();
  const rows = projects.filter((project) => (!health || project.health === health) && (!query || [project.title, project.measurementArea, project.ownerName, project.nextMilestone, project.health].some((value) => String(value || '').toLowerCase().includes(query))));
  return <div className="panel portfolio-register"><section className="portfolio-heading"><div><span className="section-kicker">Portfolio register</span><h2>Portfolio projects</h2><p>Detailed ownership, health, stage, and milestone visibility.</p></div><div className="portfolio-filters"><label className="search-box"><Icon name="search" /><input aria-label="Search portfolio projects" placeholder="Search projects, owners, milestones…" value={search} onChange={(event) => setSearch(event.target.value)} />{search && <button onClick={() => setSearch('')} aria-label="Clear search"><Icon name="close" size={14} /></button>}</label><label className="select-wrap"><Icon name="filter" /><select aria-label="Filter by health" value={health} onChange={(event) => setHealth(event.target.value)}><option value="">All health states</option>{HEALTHS.map((item) => <option key={item}>{item}</option>)}</select></label></div></section>
    <section className="table-panel"><div className="project-table" role="table"><div className="project-table-row table-header" role="row"><span>Project</span><span>Owner</span><span>Stage</span><span>Health</span><span>Progress</span><span>Next milestone</span><span /></div>{rows.map((project) => <button className="project-table-row" role="row" key={project.id} onClick={() => onOpen(project)}><span><strong>{project.title}</strong><small>{project.measurementArea}</small></span><span><span className="avatar tiny">{project.ownerName === 'Unassigned' ? '?' : project.ownerName[0]}</span>{project.ownerName}</span><span>{workflowData.phases.find((phase) => phase.key === project.currentStageKey)?.short}</span><span><Badge tone={statusTone(project.health)} dot>{project.health}</Badge></span><span><strong>{project.percentComplete}%</strong><Progress value={project.percentComplete} compact /></span><span><strong>{project.nextMilestone}</strong><small>{displayDate(project.nextMilestoneDate)}</small></span><span><Icon name="chevron" /></span></button>)}</div>{!rows.length && <p className="portfolio-no-results">No projects match the selected filters.</p>}</section>
  </div>;
}

function MyWork({ projects, tasks, user, onOpen, onTask, onDelete }) {
  const assigned = projects.filter((project) => isOwnedByUser(project, user));
  const myTasks = tasks.filter((task) => isOwnedByUser(task, user) && !isClosedTask(task)).sort((a, b) => (a.dueDate || '9999').localeCompare(b.dueDate || '9999'));
  return <div className="page-stack"><section className="page-heading"><div><span className="section-kicker">Personal workspace</span><h1>My work</h1><p>Projects and actions assigned to {user?.title || user?.email || 'you'}.</p></div></section>
    <section className="dashboard-grid my-work-grid"><div className="panel"><div className="panel-heading"><div><span className="section-kicker">Ownership</span><h2>My projects</h2></div><Badge>{assigned.length}</Badge></div>{assigned.length ? <div className="project-card-grid one-column">{assigned.map((project) => <ProjectCard key={project.id} project={project} onOpen={onOpen} onDelete={onDelete} />)}</div> : <EmptyState title="No assigned projects" message="Your manager can assign projects to you." />}</div>
    <div className="panel"><div className="panel-heading"><div><span className="section-kicker">Next actions</span><h2>My task queue</h2></div><Badge>{myTasks.length}</Badge></div><div className="task-queue">{myTasks.slice(0, 14).map((task) => <button key={task.id} onClick={() => onTask(task)}><span className={`task-state ${statusTone(task.status)}`}><Icon name={task.status === 'Blocked' ? 'alert' : 'task'} size={16} /></span><div><strong>{task.title}</strong><span>{projects.find((project) => project.projectKey === task.projectKey)?.title}</span></div><div><Badge tone={isOverdue(task) ? 'bad' : 'neutral'}>{task.dueDate ? compactDate(task.dueDate) : 'No due date'}</Badge></div></button>)}{!myTasks.length && <EmptyState title="Queue clear" message="No open tasks are assigned to your signed-in SharePoint identity." />}</div></div></section>
  </div>;
}

function Glossary({ manager, acronyms, onAdd, onDelete }) {
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
  return <div className="page-stack"><section className="page-heading inline"><div><span className="section-kicker">Modernization reference</span><h1>Acronym glossary</h1><p>A shared modernization reference maintained by managers.</p></div><label className="select-wrap glossary-search"><Icon name="search" /><input aria-label="Search acronym glossary" placeholder="Search terms or definitions…" value={query} onChange={(event) => setQuery(event.target.value)} /></label></section>
    {manager && <form className="panel glossary-composer" onSubmit={submit}><label><span>Acronym</span><input aria-label="Acronym" placeholder="e.g., CSS" value={draft.acronym} onChange={set('acronym')} /></label><label><span>Full term</span><input aria-label="Full term" placeholder="Expanded term" value={draft.term} onChange={set('term')} /></label><label className="glossary-definition"><span>Definition</span><input aria-label="Definition" placeholder="Optional plain-language definition" value={draft.definition} onChange={set('definition')} /></label><button className="button primary" type="submit" disabled={!draft.acronym.trim() || !draft.term.trim()}><Icon name="plus" /> Add acronym</button></form>}
    <section className="panel glossary-panel"><div className="glossary-table" role="table"><div className="glossary-row glossary-header" role="row"><span>Acronym</span><span>Full term</span><span>Definition</span><span>Actions</span></div>{rows.map((entry) => <div className="glossary-row" role="row" key={entry.id}><strong>{entry.acronym}</strong><span>{entry.term}</span><p>{entry.definition || '—'}</p>{manager && <button className="glossary-delete" type="button" aria-label={`Remove ${entry.acronym}`} onClick={() => onDelete(entry)}><Icon name="trash" size={16} /></button>}</div>)}</div>{!rows.length && <EmptyState title={acronyms.length ? 'No matching acronym' : 'No acronyms yet'} message={acronyms.length ? 'Try a shorter acronym, term, or keyword.' : 'Add the first acronym using the fields above.'} />}</section>
  </div>;
}

function Field({ label, children, wide = false }) { return <label className={`field ${wide ? 'field-wide' : ''}`}><span>{label}</span>{children}</label>; }

function Modal({ title, subtitle, onClose, children, actions, wide = false }) {
  useEffect(() => { const listener = (event) => event.key === 'Escape' && onClose(); window.addEventListener('keydown', listener); return () => window.removeEventListener('keydown', listener); }, [onClose]);
  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className={`modal ${wide ? 'modal-wide' : ''}`} role="dialog" aria-modal="true"><header><div><h2>{title}</h2>{subtitle && <p>{subtitle}</p>}</div><button className="icon-button" onClick={onClose} aria-label="Close"><Icon name="close" /></button></header><div className="modal-body">{children}</div>{actions && <footer>{actions}</footer>}</section></div>;
}

function ProjectEditor({ project, user, users, onClose, onSave }) {
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
  return <Modal title={project?.id ? 'Edit project' : 'Add modernization project'} subtitle="Update the portfolio record and ownership." onClose={onClose} actions={<><button className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" disabled={!draft.title.trim()} onClick={() => onSave({ ...draft, id: draft.id || uid('project'), projectKey: draft.projectKey || uid('project-key'), percentComplete: draft.percentComplete || 0, nextMilestone: draft.nextMilestone || 'Define project plan', nextMilestoneDate: draft.nextMilestoneDate || '' })}>Save project</button></>}>
    <UserPicker users={users} label="Assign project owner" onSelect={(entry) => setDraft((row) => ({ ...row, ownerName: entry.title, ownerEmail: entry.email, ownerKey: userIdentityKey(entry) }))} /><UserPicker users={users.filter((entry) => entry.role === 'Manager')} label="Project manager" onSelect={(entry) => setDraft((row) => ({ ...row, managerName: entry.title, managerEmail: entry.email || entry.loginName }))} /><p>Manager: {draft.managerName || 'Unassigned'}</p><div className="field-grid"><Field label="Project name"><input value={draft.title} onChange={set('title')} /></Field><Field label="Measurement area"><input value={draft.measurementArea} onChange={set('measurementArea')} /></Field><Field label="Project owner"><input value={draft.ownerName || ''} placeholder="Unassigned" onChange={set('ownerName')} /></Field><Field label="Owner email or SharePoint login"><div className="input-action"><input value={draft.ownerEmail || draft.ownerKey || ''} placeholder="Unassigned" onChange={setOwnerIdentity} /><button type="button" onClick={claim}>Assign me</button></div></Field><Field label="Priority"><select value={draft.priority} onChange={set('priority')}>{PRIORITIES.map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Health"><select value={draft.health} onChange={set('health')}>{HEALTHS.map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Status"><select value={draft.status} onChange={set('status')}>{['Planned', 'In Progress', 'On Hold', 'Complete', 'Cancelled'].map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Current stage"><select value={draft.currentStageKey} onChange={set('currentStageKey')}>{workflowData.phases.map((phase) => <option value={phase.key} key={phase.key}>{phase.name}</option>)}</select></Field><Field label="Target completion"><input type="date" value={draft.targetFinish || ''} onChange={set('targetFinish')} /></Field><Field label="Description" wide><textarea rows="4" value={draft.description} onChange={set('description')} /></Field></div>
  </Modal>;
}

function TaskEditor({ task, manager, users = [], onClose, onSave, onDelete }) {
  const [draft, setDraft] = useState(task);
  const [error, setError] = useState('');
  const [saving, setSaving] = useState(false);
  const set = (key) => (event) => setDraft((row) => ({ ...row, [key]: event.target.value }));
  const status = draft.status === 'Not Applicable' ? 'Not Required' : draft.status;
  async function save() {
    try { validateTask({ ...draft, status }); setSaving(true); if (await onSave({ ...draft, status })) onClose(); }
    catch (caught) { setError(caught.message); }
    finally { setSaving(false); }
  }
  return <Modal title={task.title ? 'Update task' : 'Add task'} onClose={onClose} actions={<>{manager && onDelete && <button className="button secondary" disabled={saving} onClick={async () => { if (await onDelete(task)) onClose(); }}>Delete task</button>}<button className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" disabled={saving} onClick={save}>{saving ? 'Saving…' : 'Save task'}</button></>}>
    {error && <p className="inline-error" role="alert">{error}</p>}
    {manager && <UserPicker users={users} label="Assign task owner" onSelect={(entry) => setDraft((row) => ({ ...row, ownerName: entry.title, ownerEmail: entry.email, ownerKey: userIdentityKey(entry) }))} />}
    <div className="field-grid"><Field label="Pipeline stage"><select disabled={!manager} value={draft.phaseKey} onChange={set('phaseKey')}>{workflowData.phases.map((phase) => <option value={phase.key} key={phase.key}>{phase.name}</option>)}</select></Field>
    <Field label="Task name" wide><input aria-label="Task name" disabled={!manager} value={draft.title} onChange={set('title')} /></Field>
    <Field label="Status"><select value={status} onChange={set('status')}>{TASK_STATUSES.map((item) => <option key={item}>{item}</option>)}</select></Field>
    <Field label="Original due date"><input type="date" disabled={!manager} value={draft.dueDate || ''} onChange={set('dueDate')} /></Field>
    <Field label="Deferred date"><input type="date" value={draft.deferredDate || ''} onChange={set('deferredDate')} /></Field>
    <Field label="Deferral justification" wide><textarea value={draft.deferredJustification || ''} onChange={set('deferredJustification')} /></Field>
    {status === 'Not Required' && <Field label="Not required justification" wide><textarea value={draft.notRequiredJustification || ''} onChange={set('notRequiredJustification')} /></Field>}
    <Field label="Task owner"><input disabled={!manager} value={draft.ownerName || ''} onChange={set('ownerName')} /></Field>
    {manager && <Field label="Notes" wide><textarea value={draft.notes || ''} onChange={set('notes')} /></Field>}
    </div>
  </Modal>;
}

function UserPicker({ users, label, onSelect }) {
  return <Field label={label}><select value="" onChange={(event) => { const entry = users.find((row) => row.id === event.target.value); if (entry) onSelect(entry); }}><option value="">Select a saved user…</option>{users.map((entry) => <option key={entry.id} value={entry.id}>{entry.title} — {entry.email || entry.loginName} ({entry.role})</option>)}</select></Field>;
}

function UserDirectory({ users, manager, testingEnabled, onActivate, onSave }) {
  const [password, setPassword] = useState('');
  const [activationError, setActivationError] = useState('');
  async function activate(event) {
    event.preventDefault();
    if (saving || !password) return;
    setSaving(true); setActivationError('');
    try { await onActivate(password); setPassword(''); }
    catch (error) { setActivationError(error.message || 'Manager activation failed. Please try again.'); }
    finally { setSaving(false); }
  }
  const empty = { title: '', email: '', loginName: '', role: 'User' };
  const [draft, setDraft] = useState(empty);
  const [saving, setSaving] = useState(false);
  return <section className="page-stack"><h1>Users and managers</h1><p>Signed-in users and their application roles.</p>{!manager && testingEnabled && <form className="panel directory-form" onSubmit={activate}><h2>Testing manager access</h2><p>Testing only: enter the shared testing password to grant your signed-in account manager access.</p><Field label="Testing password"><input type="password" autoComplete="off" value={password} onChange={(event) => setPassword(event.target.value)} /></Field>{activationError && <p role="alert" className="inline-error">{activationError}</p>}<button type="button" className="button primary" onClick={activate} disabled={saving || !password}>{saving ? 'Saving and verifying role…' : 'Enable manager access'}</button>{saving && <p role="status">Please wait while your manager role is saved and verified.</p>}</form>}{manager && <form className="panel directory-form" onSubmit={async (event) => { event.preventDefault(); setSaving(true); try { if (await onSave({ ...draft, id: draft.id || uid('user'), loginName: draft.loginName.trim() || draft.email.trim() })) setDraft(empty); } finally { setSaving(false); } }}>
    {['title', 'email', 'loginName'].map((key) => <Field key={key} label={{ title: 'Name', email: 'Email', loginName: 'SharePoint login' }[key]}><input required={key === 'title'} value={draft[key]} onChange={(event) => setDraft((row) => ({ ...row, [key]: event.target.value }))} /></Field>)}
    <Field label="Role"><select value={draft.role} onChange={(event) => setDraft((row) => ({ ...row, role: event.target.value }))}><option>User</option><option>Manager</option></select></Field><button className="button primary" disabled={saving || !(draft.loginName.trim() || draft.email.trim())}>Save user</button>{draft.id && <button type="button" className="button secondary" onClick={() => setDraft(empty)}>Cancel edit</button>}
    </form>}{users.map((entry) => <article className="panel directory-row" key={entry.id}><strong>{entry.title}</strong><span>{entry.email || entry.loginName}</span><Badge>{entry.role}</Badge>{manager && <button className="button secondary" onClick={() => setDraft(entry)}>Edit</button>}</article>)}</section>;
}

function ProjectDrawer({ project, user, manager, users, onDeleteTask, tasks, updates, risks, onClose, onEdit, onSaveTask, onToggleTask, onSetPhaseRequired, onAddUpdate, onAddRisk, onClaim }) {
  const [tab, setTab] = useState('overview');
  const [taskEditor, setTaskEditor] = useState(null);
  const [updateText, setUpdateText] = useState('');
  const [riskDraft, setRiskDraft] = useState({ title: '', severity: 'Medium', probability: 'Possible', mitigation: '' });
  const projectTasks = tasks.filter((task) => task.projectKey === project.projectKey).sort((a, b) => a.order - b.order);
  const projectUpdates = updates.filter((item) => item.projectKey === project.projectKey).sort((a, b) => String(b.entryDate).localeCompare(String(a.entryDate)));
  const projectRisks = risks.filter((item) => item.projectKey === project.projectKey);
  const phase = workflowData.phases.find((item) => item.key === project.currentStageKey);
  return <div className="drawer-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><aside className="project-drawer" role="dialog" aria-modal="true"><header className="drawer-header"><div className="drawer-title-row"><div><span className="section-kicker">{project.measurementArea}</span><h2>{project.title}</h2><div className="drawer-badges"><Badge tone={statusTone(project.health)} dot>{project.health}</Badge><Badge>{project.priority} priority</Badge><Badge>{project.status}</Badge></div></div><button className="icon-button" onClick={onClose} aria-label="Close"><Icon name="close" /></button></div>
    {manager && <div className="drawer-actions"><button className="button secondary" onClick={() => onEdit(project)}>Edit project</button>{!project.ownerKey && <button className="button primary" onClick={() => onClaim(project)}>Claim project</button>}</div>}
    <div className="drawer-progress"><div><span>{project.percentComplete}% complete</span><strong>{phase?.name}</strong></div><Progress value={project.percentComplete} /></div>
    <div className="mini-pipeline">{workflowData.phases.map((item, index) => <span key={item.key} className={index < phaseIndex(project.currentStageKey) ? 'done' : item.key === project.currentStageKey ? 'current' : ''} title={item.name}>{index + 1}</span>)}</div>
    <nav className="drawer-tabs">{[['overview', 'Overview'], ['tasks', `Work breakdown (${projectTasks.length})`], ['updates', `Updates (${projectUpdates.length})`], ['risks', `Risks (${projectRisks.length})`]].map(([key, label]) => <button className={tab === key ? 'active' : ''} key={key} onClick={() => setTab(key)}>{label}</button>)}</nav></header>
    <div className="drawer-content">
      {tab === 'overview' && <div className="drawer-section-stack"><section className="detail-grid"><div><span>Owner</span><strong>{project.ownerName}</strong><small>{project.ownerEmail || 'No SharePoint identity assigned'}</small></div><div><span>Target completion</span><strong>{displayDate(project.targetFinish)}</strong><small>{project.status}</small></div><div><span>Next milestone</span><strong>{project.nextMilestone}</strong><small>{displayDate(project.nextMilestoneDate)}</small></div><div><span>Open actions</span><strong>{projectTasks.filter((task) => !isClosedTask(task)).length}</strong><small>{project.blockedCount} blocked</small></div></section><section className="detail-block"><h3>Purpose</h3><p>{project.description || 'No project description has been added.'}</p></section><section className="detail-block"><h3>Upcoming work</h3><div className="upcoming-list">{projectTasks.filter((task) => !isClosedTask(task)).slice(0, 5).map((task) => <button key={task.id} onClick={() => setTaskEditor(task)}><div><strong>{task.title}</strong><small>{workflowData.phases.find((item) => item.key === task.phaseKey)?.short}</small></div><Badge tone={statusTone(task.status)}>{task.status}</Badge></button>)}</div></section></div>}
      {tab === 'tasks' && <div className="phase-task-list">{workflowData.phases.map((stage) => { const stageTasks = projectTasks.filter((task) => task.phaseKey === stage.key); const done = stageTasks.filter(isClosedTask).length; const allNotRequired = stageTasks.length > 0 && stageTasks.every((task) => ['Not Required', 'Not Applicable'].includes(task.status)); return <section key={stage.key} className={`task-stage ${allNotRequired ? 'phase-not-required' : ''}`}><header><div><span>{String(phaseIndex(stage.key) + 1).padStart(2, '0')}</span><strong>{stage.name}</strong></div><div className="phase-header-actions"><Badge>{done}/{stageTasks.length}</Badge>{manager && <button className="button secondary small-button" onClick={() => setTaskEditor({ id: uid('task'), projectKey: project.projectKey, title: '', phaseKey: stage.key, order: Math.max(0, ...projectTasks.map((task) => task.order || 0)) + 1, status: 'Not Started', ownerName: project.ownerName, ownerKey: project.ownerKey, ownerEmail: project.ownerEmail, dueDate: '' })}>Add task</button>}{manager && <button type="button" className={`phase-required-button ${allNotRequired ? 'active' : ''}`} disabled={!stageTasks.length} aria-label={`${allNotRequired ? 'Restore' : 'Mark'} ${stage.name} ${allNotRequired ? '' : 'not required'}`.trim()} onClick={() => onSetPhaseRequired(project, stage.key, !allNotRequired)}><Icon name={allNotRequired ? 'refresh' : 'ban'} size={13} />{allNotRequired ? 'Restore phase' : 'Not required'}</button>}</div></header><div>{stageTasks.map((task) => { const resolved = isClosedTask(task); return <div key={task.id} className="phase-task" role="button" tabIndex="0" onClick={() => setTaskEditor(task)} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); setTaskEditor(task); } }}><button className={`task-check ${resolved ? 'checked' : ''}`} aria-label={`${resolved ? 'Reopen' : 'Complete'} ${task.title}`} onClick={(event) => { event.stopPropagation(); onToggleTask(task); }}>{resolved && <Icon name="check" size={13} />}</button><div><strong>{task.title}</strong>{task.dataIssue && <small className="issue-text"><Icon name="alert" size={12} />{task.dataIssue}</small>}</div><Badge tone={isOverdue(task) ? 'bad' : statusTone(task.status)}>{task.status === 'Not Applicable' ? 'Not Required' : task.status}</Badge><span className="task-date">{isClosedTask(task) ? 'Resolved' : compactDate(task.dueDate)}</span></div>; })}</div></section>; })}</div>}
      {tab === 'updates' && <div className="drawer-section-stack">{manager && <section className="composer"><textarea rows="3" placeholder="Add a concise status update, decision, or handoff…" value={updateText} onChange={(event) => setUpdateText(event.target.value)} /><div><span>Visible to everyone with access to this SharePoint workspace</span><button className="button primary small-button" disabled={!updateText.trim()} onClick={() => { onAddUpdate(project, updateText); setUpdateText(''); }}>Post update</button></div></section>}<section className="timeline-list">{projectUpdates.map((item) => <article key={item.id}><span className="timeline-dot" /><div><header><strong>{item.authorName}</strong><span>{displayDate(item.entryDate)}</span></header><Badge>{item.type}</Badge><p>{item.summary}</p></div></article>)}{!projectUpdates.length && <EmptyState title="No updates yet" message="Post the first status update to establish the project history." />}</section></div>}
      {tab === 'risks' && <div className="drawer-section-stack">{manager && <section className="risk-composer"><div className="field-grid"><Field label="Risk or issue"><input value={riskDraft.title} onChange={(event) => setRiskDraft((row) => ({ ...row, title: event.target.value }))} /></Field><Field label="Severity"><select value={riskDraft.severity} onChange={(event) => setRiskDraft((row) => ({ ...row, severity: event.target.value }))}>{['Low', 'Medium', 'High', 'Critical'].map((item) => <option key={item}>{item}</option>)}</select></Field><Field label="Mitigation" wide><textarea rows="3" value={riskDraft.mitigation} onChange={(event) => setRiskDraft((row) => ({ ...row, mitigation: event.target.value }))} /></Field></div><button className="button primary small-button" disabled={!riskDraft.title.trim()} onClick={() => { onAddRisk(project, riskDraft); setRiskDraft({ title: '', severity: 'Medium', probability: 'Possible', mitigation: '' }); }}>Add risk</button></section>}<section className="risk-list">{projectRisks.map((risk) => <article key={risk.id}><div><Badge tone={statusTone(risk.severity)}>{risk.severity}</Badge><Badge>{risk.status}</Badge></div><h3>{risk.title}</h3><p>{risk.mitigation || 'No mitigation has been documented.'}</p><span>{risk.ownerName || 'Unassigned'}</span></article>)}{!projectRisks.length && <EmptyState title="No open risks" message="Capture risks and mitigation actions here as the project advances." />}</section></div>}
    </div>
    {taskEditor && <TaskEditor task={taskEditor} manager={manager} users={users} onClose={() => setTaskEditor(null)} onSave={onSaveTask} onDelete={projectTasks.some((row) => row.id === taskEditor.id) ? onDeleteTask : undefined} />}
  </aside></div>;
}

function Toast({ toast, onClose }) { if (!toast) return null; return <div className={`toast toast-${toast.tone || 'good'}`} role={toast.tone === 'bad' ? 'alert' : 'status'} aria-live={toast.tone === 'bad' ? 'assertive' : 'polite'}><Icon name={toast.tone === 'bad' ? 'alert' : toast.tone === 'info' ? 'clock' : 'check'} /><span>{toast.message}</span><button type="button" aria-label="Dismiss notification" onClick={onClose}><Icon name="close" size={13} /></button></div>; }

export function App() {
  const repoRef = useRef(null);
  if (!repoRef.current) { const repository = createRepository(); repoRef.current = { ...repository, store: authorizedStore(repository.store, repository.config) }; }
  const repo = repoRef.current;
  const [theme, setTheme] = useState(() => localStorage.getItem('mod-tracker-theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'));
  const [view, setView] = useState('my-work');
  const [user, setUser] = useState(null);
  const [data, setData] = useState({ projects: [], tasks: [], updates: [], risks: [], acronyms: [] });
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [phaseFilter, setPhaseFilter] = useState([]);
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
      await repo.store.registerCurrentUser();
      const loaded = await repo.store.load();
      setData(loaded);
      setView(isManager(currentUser, loaded.users) ? 'overview' : 'my-work');
    } catch (caught) { setError(caught.message || 'Could not load portfolio data.'); }
    finally { setLoading(false); }
  }

  const manager = isManager(user, data.users);
  const enriched = useMemo(() => data.projects.map((project) => enrichProject(project, data.tasks)), [data.projects, data.tasks]);
  const openProject = enriched.find((project) => project.id === openProjectId);

  async function saveProject(project) {
    const source = data;
    const prepared = { ...project, ownerName: project.ownerName?.trim() || 'Unassigned' };
    const creating = !prepared.spId && !source.projects.some((row) => row.id === prepared.id);
    const previous = source.projects.find((row) => row.id === prepared.id);
    const template = creating ? createStarterTasks(prepared.projectKey, prepared) : source.tasks.filter((task) => task.projectKey === prepared.projectKey && previous && (previous.ownerKey !== prepared.ownerKey || previous.ownerEmail !== prepared.ownerEmail || previous.ownerName !== prepared.ownerName) && task.ownerKey === previous.ownerKey && task.ownerEmail === previous.ownerEmail).map((task) => ({ ...task, ownerName: prepared.ownerName, ownerKey: prepared.ownerKey, ownerEmail: prepared.ownerEmail }));
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
      if (template.length) {
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
      const normalized = { ...task, status, ownerName: task.ownerName?.trim() || 'Unassigned', finishDate: resolved ? (task.finishDate || todayIso()) : '' };
      const saved = await repo.store.saveTask(normalized);
      const updateState = (state) => ({ ...state, tasks: state.tasks.some((row) => row.id === saved.id) ? state.tasks.map((row) => row.id === saved.id ? saved : row) : [...state.tasks, saved] });
      setData(updateState);
      setToast({ message: `${saved.title} updated.` });
      return true;
    } catch (caught) { setToast({ tone: 'bad', message: caught.message }); return false; }
  }

  async function deleteTask(task) {
    if (!window.confirm(`Delete task “${task.title}”?`)) return false;
    try { await repo.store.recycle('tasks', task.spId, task.id); setData((state) => ({ ...state, tasks: state.tasks.filter((row) => row.id !== task.id) })); return true; }
    catch (caught) { setToast({ tone: 'bad', message: caught.message }); return false; }
  }

  async function saveUser(row) {
    try { const saved = await repo.store.saveUser(row); setData((state) => ({ ...state, users: (state.users || []).some((entry) => entry.id === saved.id) ? state.users.map((entry) => entry.id === saved.id ? saved : entry) : [...(state.users || []), saved] })); setToast({ message: 'User saved.' }); return true; }
    catch (caught) { setToast({ tone: 'bad', message: caught.message }); return false; }
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
    const justification = notRequired ? window.prompt('Why is this phase not required?')?.trim() : '';
    if (notRequired && !justification) return;
    const changed = affected.map((task) => ({
      ...task,
      status: notRequired ? 'Not Required' : 'Not Started',
      notRequiredJustification: notRequired ? justification : task.notRequiredJustification,
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
    <aside className="sidebar"><div className="brand"><img className="brand-logo" src={navairSeal} alt="NAVAIR" /><div><strong>MODERNIZATION</strong><span>Project Tracker</span></div></div><nav>{[...NAV.filter(([key]) => manager || ['my-work', 'glossary'].includes(key)), ['users', 'Users and managers', 'tasks']].map(([key, label, icon]) => <button className={view === key ? 'active' : ''} key={key} onClick={() => setView(key)}><Icon name={icon} /><span>{label}</span>{key === 'my-work' && <Badge>{enriched.filter((project) => isOwnedByUser(project, user)).length}</Badge>}</button>)}</nav><div className="sidebar-footer"><div className="sidebar-user"><span className="avatar">{userInitials(user)}</span><div><strong>{user?.title || 'SharePoint user'}</strong><span>{user?.email || user?.loginName || 'Full portfolio access'}</span></div></div></div></aside>
    <div className="main-shell"><header className="topbar"><div className="mobile-brand"><img className="brand-logo" src={navairSeal} alt="NAVAIR" /><strong>MODERNIZATION</strong></div><div className="top-actions"><button className="icon-button" onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')} aria-label={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}><Icon name={theme === 'dark' ? 'sun' : 'moon'} /></button><button className="button secondary export-button" disabled={exporting} onClick={exportWorkbook}><Icon name="download" /> {exporting ? 'Building Excel…' : 'Export Excel'}</button>{manager && <button className="button primary" onClick={() => setProjectEditor({})}><Icon name="plus" /> New project</button>}</div></header>
      <main>
        {manager && view === 'overview' && <Overview projects={enriched} tasks={data.tasks} risks={data.risks} onOpen={(project) => setOpenProjectId(project.id)} onDelete={manager ? deleteProject : undefined} phaseFilter={phaseFilter} setPhaseFilter={setPhaseFilter} onAttention={() => setView('attention')} />}
        {manager && view === 'board' && <Board projects={enriched} onOpen={(project) => setOpenProjectId(project.id)} onDelete={manager ? deleteProject : undefined} />}
        {manager && view === 'attention' && <Attention projects={enriched} onOpen={(project) => setOpenProjectId(project.id)} />}
        {view === 'users' && <UserDirectory users={data.users || []} manager={manager} testingEnabled={repo.config.testingManagerPassword !== false && repo.config.testingManagerPassword !== ''} onActivate={async (password) => { const saved = await repo.store.activateTestingManager(password); const loaded = await repo.store.load(); if (!isManager(user, loaded.users)) throw new Error('The directory still reports your account as a basic user. Manager access could not be confirmed; please reload and try again.'); setData(loaded); setToast({ message: `Manager access confirmed for ${saved.title}.` }); }} onSave={saveUser} />}
        {view === 'my-work' && <MyWork projects={enriched} tasks={data.tasks} user={user} onOpen={(project) => setOpenProjectId(project.id)} onTask={setTaskEditor} onDelete={manager ? deleteProject : undefined} />}
        {view === 'glossary' && <Glossary manager={manager} acronyms={data.acronyms} onAdd={addAcronym} onDelete={deleteAcronym} />}
      </main>
    </div>
    {openProject && <ProjectDrawer project={openProject} user={user} manager={manager} users={data.users || []} onDeleteTask={deleteTask} tasks={data.tasks} updates={data.updates} risks={data.risks} onClose={() => setOpenProjectId('')} onEdit={setProjectEditor} onSaveTask={saveTask} onToggleTask={toggleTaskComplete} onSetPhaseRequired={setPhaseRequired} onAddUpdate={addUpdate} onAddRisk={addRisk} onClaim={claimProject} />}
    {manager && projectEditor && <ProjectEditor project={projectEditor.id ? projectEditor : null} user={user} users={data.users || []} onClose={() => setProjectEditor(null)} onSave={saveProject} />}
    {taskEditor && <TaskEditor task={taskEditor} manager={manager} users={data.users || []} onClose={() => setTaskEditor(null)} onSave={saveTask} />}
    <Toast toast={toast} onClose={() => setToast(null)} />
  </div>;
}
