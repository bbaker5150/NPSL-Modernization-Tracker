import { ProjectUpdates } from './components/ProjectUpdates';
import { Feedback } from './components/Feedback';
import { PeopleCombobox, SavedUserPicker } from './components/SavedUserPicker';
import { AutoSizeTextarea } from './components/AutoSizeTextarea';
import { Attention, attentionRows } from './components/Attention';
import { validateAttachment } from './lib/taskAttachments';
import { ReferenceDocuments } from './components/ReferenceDocuments';
import { TrackerAccessSetup } from './components/TrackerAccessSetup';
import { PeopleInvite } from './components/PeopleInvite';
import { TaskDocuments } from './components/TaskDocuments';
import { FilePicker } from './components/FilePicker';
import { ProjectTeam, engineerInitials, projectTeamSearch, projectTeamText } from './components/ProjectTeam';
import { engineerFields, projectEngineers } from './lib/projectEngineers';
import { displayName } from './lib/displayName';
import { METCAL_ORGANIZATION, METCAL_TEAMS, ORGANIZATIONS, ROLES, normalizeOrganization, organizationAttentionLabel, projectProgress } from './data/workflow';
import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Icon } from './components/Icon';
import { createRepository, createStarterTasks, isOwnedByUser, workflowData, userIdentityKey } from './lib/repository';
import { downloadPortfolioWorkbook } from './lib/exportPortfolioWorkbook';
import { authorizedStore, isManager, canManageProject, canViewPortfolio, validateTask } from './lib/access';
import navairSeal from './assets/navair-seal-384.webp';

const NAV = [
  ['overview', 'Portfolio', 'overview'],
  ['board', 'Pipeline board', 'board'],
  ['my-work', 'My work', 'tasks'],
  ['references', 'Reference Documents', 'folder'],
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
function isOverdue(task) { const deadline = task.deferredDate || task.dueDate; return !isClosedTask(task) && deadline && deadline < todayIso(); }
function statusTone(value) {
  if (/complete|not required|not applicable|on track/i.test(value)) return 'good';
  if (/blocked|critical|program office/i.test(value)) return 'bad';
  if (/risk|high|review|progress/i.test(value)) return 'warn';
  return 'neutral';
}

function enrichProject(project, allTasks) {
  const tasks = allTasks.filter((task) => task.projectKey === project.projectKey);

  const orderedTasks = [...tasks].sort((a, b) => phaseIndex(a.phaseKey) - phaseIndex(b.phaseKey) || a.order - b.order);
  const nextTask = orderedTasks.find((task) => !isClosedTask(task));
  const finalPhase = [...workflowData.phases].reverse().find((phase) => tasks.some((task) => task.phaseKey === phase.key && isClosedTask(task)));
  return {
    ...project,
    ownerName: project.ownerName || 'Unassigned',
    organization: [...new Set(tasks.filter((task) => !isClosedTask(task)).map((task) => task.organization))].filter(Boolean).join(', ') || 'No outstanding tasking',
    tasks,
    ...projectProgress(tasks, project.progressMode || 'phases'),
    currentStageKey: project.currentStageKey || nextTask?.phaseKey || finalPhase?.key || 'requirement',
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
  return <div className="loading-shell"><img className="brand-logo loading-logo" src={navairSeal} alt="NAVAIR" /><div><strong>Modernization Tracker</strong><span>Loading portfolio data…</span></div></div>;
}

function KpiCard({ icon, label, value, detail, tone = 'blue', onClick }) {
  return <article className={`kpi-card kpi-${tone}`} role={onClick ? "button" : undefined} tabIndex={onClick ? 0 : undefined} onClick={onClick} onKeyDown={(event) => { if (onClick && ['Enter', ' '].includes(event.key)) { event.preventDefault(); onClick(); } }}><div className="kpi-icon"><Icon name={icon} /></div><div><span>{label}</span><strong>{value}</strong><small>{detail}</small></div></article>;
}

function PipelineRail({ projects, selected, onSelect }) {
  return <div className="pipeline-rail">
    {workflowData.phases.map((phase, index) => {
      const count = projects.filter((project) => project.currentStageKey === phase.key).length;
      return <button key={phase.key} className={`phase-node ${selected.includes(phase.key) ? 'selected' : ''}`} aria-pressed={selected.includes(phase.key)} onClick={(event) => onSelect?.(event.ctrlKey || event.metaKey ? (selected.includes(phase.key) ? selected.filter((key) => key !== phase.key) : [...selected, phase.key]) : (selected.length === 1 && selected[0] === phase.key ? [] : [phase.key]))}>
        <span className="phase-label">{phase.name}</span>
        <span className="phase-line" aria-hidden="true" />
        <span className="phase-count" aria-label={`${count} projects`}>{count}</span>

      </button>;
    })}
  </div>;
}

function useDismissibleMenu() {
  const [menuOpen, setMenuOpen] = useState(false);
  const menuRef = useRef(null);
  useEffect(() => {
    if (!menuOpen) return;
    const dismiss = (event) => { if (!menuRef.current?.contains(event.target)) setMenuOpen(false); };
    const escape = (event) => { if (event.key === 'Escape') { setMenuOpen(false); menuRef.current?.querySelector('button')?.focus(); } };
    document.addEventListener('pointerdown', dismiss);
    document.addEventListener('focusin', dismiss);
    document.addEventListener('keydown', escape);
    return () => { document.removeEventListener('pointerdown', dismiss); document.removeEventListener('focusin', dismiss); document.removeEventListener('keydown', escape); };
  }, [menuOpen]);
  return { menuOpen, setMenuOpen, menuRef };
}

function ProjectCard({ project, onOpen, onDelete }) {
  const { menuOpen, setMenuOpen, menuRef } = useDismissibleMenu();
  const phase = workflowData.phases.find((item) => item.key === project.currentStageKey);
  const openProject = () => onOpen(project);
  return <article className="project-card" role="button" tabIndex="0" onClick={openProject} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); openProject(); } }}>
    <div className="project-card-top"><div className="project-card-badges"><Badge tone={statusTone(project.health)} dot>{project.health}</Badge><Badge>{project.organization}</Badge></div><div ref={menuRef} className="project-menu-wrap" hidden={!onDelete}><button className="project-menu" aria-label={`Project actions for ${project.title}`} aria-expanded={menuOpen} onClick={(event) => { event.stopPropagation(); setMenuOpen((open) => !open); }}><Icon name="more" size={16} /></button>{menuOpen && onDelete && <div className="project-menu-popover"><button onClick={(event) => { event.stopPropagation(); setMenuOpen(false); onDelete(project); }}>Delete project</button></div>}</div></div>
    <div><span className="eyebrow">{project.measurementArea}</span><h3>{project.title}</h3></div>
    <div className="project-owner"><ProjectTeam project={project} /></div>
    <div className="project-phase"><span>Stage {phaseIndex(project.currentStageKey) + 1} of {workflowData.phases.length}</span><strong>{phase?.short}</strong></div>
    <Progress value={project.percentComplete} compact />
    <div className="project-card-bottom"><span>{project.percentComplete}% complete</span><span title={`Target completion: ${displayDate(project.targetFinish)}`}><Icon name="calendar" size={14} /> {project.targetFinish ? compactDate(project.targetFinish) : 'No target date'}</span></div>
  </article>;
}

function Overview({ projects, tasks, onOpen, phaseFilter, setPhaseFilter, onAttention, onBoard, register, setRegister }) {
  const active = projects.filter((project) => project.tasks.some((task) => !isClosedTask(task)) || !project.tasks.length);
  const filtered = phaseFilter.length ? projects.filter((project) => phaseFilter.includes(project.currentStageKey)) : projects;
  return <div className="page-stack">
    <section className="hero-row"><div><span className="section-kicker">Portfolio command center</span><h1>Modernization at a glance</h1><p>Measurement areas, milestones, and handoffs across the modernization pipeline.</p></div></section>
    <section className="kpi-grid">
      <KpiCard icon="projects" label="Active projects" value={active.length} detail={`${projects.length} total measurement areas`} onClick={onBoard} />
      {ORGANIZATIONS.map((organization) => { const count = tasks.filter((task) => task.organization === organization && !isClosedTask(task)).length; return <KpiCard key={organization} icon="alert" label={organizationAttentionLabel(organization)} value={count} detail="Outstanding tasks" tone="amber" onClick={() => onAttention(organization)} />; })}
    </section>
    <section className="panel pipeline-panel"><div className="panel-heading"><h2>Projects by pipeline stage</h2>{phaseFilter.length > 0 && <button className="text-button" onClick={() => setPhaseFilter([])}>Clear filter</button>}</div><PipelineRail projects={projects} selected={phaseFilter} onSelect={setPhaseFilter} /></section>
    {projects.length > 0 && <ProjectsTable projects={filtered} onOpen={onOpen} register={register} setRegister={setRegister} />}
  </div>;
}

function ProgressChoice({ value, onChange, label = 'Progress calculation' }) {
  return <label className="progress-choice" title={`Progress by ${value}`}><Icon name="chevron" size={15} /><select aria-label={label} value={value} onChange={(event) => onChange(event.target.value)}><option value="phases">By phase</option><option value="tasks">By task</option></select></label>;
}

function Board({ projects, onOpen, onDelete }) {
  const [sortBy, setSortBy] = useState('default');
  const sortedProjects = [...projects].sort((a, b) => sortBy === 'organization' ? a.organization.localeCompare(b.organization) || a.title.localeCompare(b.title) : 0);
  return <div className="page-stack"><section className="page-heading inline"><div><span className="section-kicker">End-to-end flow</span><h1>Pipeline board</h1><p>Scan where every project sits and open a card to update its work breakdown.</p></div><label className="select-wrap"><Icon name="filter" /><select aria-label="Sort pipeline projects" value={sortBy} onChange={(event) => setSortBy(event.target.value)}><option value="default">Default order</option><option value="organization">Sort by organization</option></select></label></section>
    <div className="board-scroll"><div className="kanban-board">{workflowData.phases.map((phase, index) => {
      const rows = sortedProjects.filter((project) => project.currentStageKey === phase.key);
      return <section className="kanban-column" key={phase.key}><header><div><span>{String(index + 1).padStart(2, '0')}</span><strong>{phase.name}</strong></div><Badge>{rows.length}</Badge></header><div className="kanban-list">{rows.map((project) => <ProjectCard key={project.id} project={project} onOpen={onOpen} onDelete={onDelete} />)}{!rows.length && <div className="kanban-empty">No projects in this stage</div>}</div></section>;
    })}</div></div>
  </div>;
}

function ProjectsTable({ projects, onOpen, register, setRegister }) {
  const { progressMode, health, organization, search, sortKey, sortDirection } = register;
  const setProgressMode = (progressMode) => setRegister((state) => ({ ...state, progressMode }));
  const setHealth = (health) => setRegister((state) => ({ ...state, health }));
  const setSearch = (search) => setRegister((state) => ({ ...state, search }));
  const setOrganization = (organization) => setRegister((state) => ({ ...state, organization }));
  const sort = (key) => setRegister((state) => ({ ...state, sortKey: key, sortDirection: state.sortKey === key && state.sortDirection === 'asc' ? 'desc' : 'asc' }));
  const query = search.trim().toLowerCase();
  const healthOrder = new Map(HEALTHS.map((value, index) => [value, index]));
  const values = (project, key) => ({ project: project.title, owner: projectTeamText(project), organization: project.organization, stage: phaseIndex(project.currentStageKey), health: healthOrder.get(project.health) ?? 99, progress: projectProgress(project.tasks, progressMode).percentComplete, milestone: project.nextMilestone })[key];
  const rows = projects.filter((project) => (!health || project.health === health) && (!organization || project.tasks.some((task) => !isClosedTask(task) && task.organization === organization)) && (!query || [project.title, project.measurementArea, projectTeamSearch(project), project.organization, project.nextMilestone, project.health].some((value) => String(value || '').toLowerCase().includes(query)))).sort((a, b) => { const left = values(a, sortKey); const right = values(b, sortKey); const result = typeof left === 'number' ? left - right : String(left || '').localeCompare(String(right || '')); return (sortDirection === 'desc' ? -1 : 1) * (result || a.title.localeCompare(b.title)); });
  const heading = (key, label, className = '') => <button type="button" className={`table-sort ${className}`} onClick={() => sort(key)} aria-sort={sortKey === key ? (sortDirection === 'asc' ? 'ascending' : 'descending') : 'none'}>{label}{sortKey === key && <span aria-hidden="true">{sortDirection === 'asc' ? '↑' : '↓'}</span>}</button>;
  return <div className="panel portfolio-register"><section className="portfolio-heading"><div><span className="section-kicker">Portfolio register</span><h2>Portfolio projects</h2><p>Detailed ownership, organization, health, stage, and milestone visibility.</p></div><div className="portfolio-filters"><label className="search-box"><Icon name="search" /><input aria-label="Search portfolio projects" placeholder="Search projects, engineers, milestones…" value={search} onChange={(event) => setSearch(event.target.value)} />{search && <button onClick={() => setSearch('')} aria-label="Clear search"><Icon name="close" size={14} /></button>}</label><label className="select-wrap"><Icon name="filter" /><select aria-label="Filter by organization" value={organization} onChange={(event) => setOrganization(event.target.value)}><option value="">All organizations</option>{ORGANIZATIONS.map((item) => <option key={item}>{item}</option>)}</select></label><label className="select-wrap"><Icon name="filter" /><select aria-label="Filter by health" value={health} onChange={(event) => setHealth(event.target.value)}><option value="">All health states</option>{HEALTHS.map((item) => <option key={item}>{item}</option>)}</select></label></div></section>
    <section className="table-panel"><div className="project-table" role="table"><div className="project-table-row table-header" role="row">{heading('project', 'Project')}{heading('owner', 'Project engineers')}{heading('organization', 'Organization')}{heading('stage', 'Stage')}{heading('health', 'Health')}<span className="progress-column-heading">{heading('progress', 'Progress')}<ProgressChoice label="Portfolio progress calculation" value={progressMode} onChange={setProgressMode} /></span>{heading('milestone', 'Next milestone')}<span /></div>{rows.map((project) => <button className="project-table-row" role="row" key={project.id} onClick={() => onOpen(project)}><span><strong>{project.title}</strong><small>{project.measurementArea}</small></span><span><ProjectTeam project={project} /></span><span>{project.organization}</span><span>{workflowData.phases.find((phase) => phase.key === project.currentStageKey)?.short}</span><span><Badge tone={statusTone(project.health)} dot>{project.health}</Badge></span><span title={projectProgress(project.tasks, progressMode).progressDetail}><strong>{projectProgress(project.tasks, progressMode).percentComplete}%</strong><Progress value={projectProgress(project.tasks, progressMode).percentComplete} compact /></span><span><strong>{project.nextMilestone}</strong><small>{displayDate(project.nextMilestoneDate)}</small></span><span><Icon name="chevron" /></span></button>)}</div>{!rows.length && <p className="portfolio-no-results">No projects match the selected filters.</p>}</section>
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
  const [adding, setAdding] = useState(false);
  const normalizedQuery = query.trim().toLowerCase();
  const rows = [...acronyms]
    .filter((entry) => [entry.acronym, entry.term, entry.definition].some((value) => String(value || '').toLowerCase().includes(normalizedQuery)))
    .sort((a, b) => a.acronym.localeCompare(b.acronym));
  const set = (key) => (event) => setDraft((entry) => ({ ...entry, [key]: event.target.value }));
  const submit = async (event) => {
    event.preventDefault();
    if (adding || !draft.acronym.trim() || !draft.term.trim()) return;
    setAdding(true);
    try {
      if (await onAdd(draft)) setDraft({ acronym: '', term: '', definition: '' });
    } finally { setAdding(false); }
  };
  return <div className="page-stack"><section className="page-heading inline"><div><span className="section-kicker">Modernization reference</span><h1>Acronym glossary</h1><p>Add terms and definitions to the shared glossary.</p></div><label className="select-wrap glossary-search"><Icon name="search" /><input aria-label="Search acronym glossary" placeholder="Search terms or definitions…" value={query} onChange={(event) => setQuery(event.target.value)} /></label></section>
    <form className="panel glossary-composer" onSubmit={submit} onKeyDown={event => {
      if (event.key === 'Enter' && !event.nativeEvent.isComposing && event.target.tagName === 'INPUT') submit(event);
    }}>
      <label><span>Acronym</span><input aria-label="Acronym" disabled={adding} placeholder="e.g., CSS" value={draft.acronym} onChange={set('acronym')} /></label>
      <label><span>Full term</span><input aria-label="Full term" disabled={adding} placeholder="Expanded term" value={draft.term} onChange={set('term')} /></label>
      <label className="glossary-definition"><span>Definition</span><input aria-label="Definition" disabled={adding} placeholder="Optional plain-language definition" value={draft.definition} onChange={set('definition')} /></label>
      <button className="button primary" type="button" onClick={submit} disabled={adding || !draft.acronym.trim() || !draft.term.trim()}><Icon name="plus" /> Add acronym</button>
    </form>
    <section className="panel glossary-panel"><div className="glossary-table" role="table"><div className="glossary-row glossary-header" role="row"><span>Acronym</span><span>Full term</span><span>Definition</span>{manager && <span>Actions</span>}</div>{rows.map((entry) => <div className="glossary-row" role="row" key={entry.id}><strong>{entry.acronym}</strong><span>{entry.term}</span><p>{entry.definition || '—'}</p>{manager && <button className="glossary-delete" type="button" aria-label={`Remove ${entry.acronym}`} onClick={() => onDelete(entry)}><Icon name="trash" size={16} /></button>}</div>)}</div>{!rows.length && <EmptyState title={acronyms.length ? 'No matching acronym' : 'No acronyms yet'} message={acronyms.length ? 'Try a shorter acronym, term, or keyword.' : 'Add the first acronym using the fields above.'} />}</section>
  </div>;
}

function Field({ label, children, wide = false }) { return <label className={`field ${wide ? 'field-wide' : ''}`}><span>{label}</span>{children}</label>; }

function Modal({ title, subtitle, onClose, children, actions, wide = false }) {
  useEffect(() => { const listener = (event) => event.key === 'Escape' && onClose(); window.addEventListener('keydown', listener); return () => window.removeEventListener('keydown', listener); }, [onClose]);
  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className={`modal ${wide ? 'modal-wide' : ''}`} role="dialog" aria-modal="true"><header><div><h2>{title}</h2>{subtitle && <p>{subtitle}</p>}</div><button className="icon-button" onClick={onClose} aria-label="Close"><Icon name="close" /></button></header><div className="modal-body">{children}</div>{actions && <footer>{actions}</footer>}</section></div>;
}

const PROJECT_STATUSES = ['Planned', 'In Progress', 'On Hold', 'Complete', 'Cancelled'];
const isEmail = value => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);

function engineerEmail(person) {
  const signIn = person?.loginName?.split('|').pop();
  return person?.email || (isEmail(signIn) ? signIn : '');
}

function sameEngineer(engineer, person) {
  return isOwnedByUser({ ownerKey: engineer.loginName, ownerEmail: engineer.email }, person);
}

async function verifyEngineer(store, candidate, typedEmail) {
  const person = await store.resolvePerson(candidate?.loginName || typedEmail);
  if (!person?.title || !person.loginName ||
      (candidate && userIdentityKey(person) !== userIdentityKey(candidate)) ||
      (typedEmail && !isOwnedByUser({ ownerEmail: typedEmail }, person))) {
    throw new Error('The resolved SharePoint account does not match the selection or email. Check the address and try again.');
  }
  const email = engineerEmail(person) || candidate?.email;
  if (!isEmail(email)) throw new Error('SharePoint did not return an email address for this account. Choose another account or check its directory entry.');
  return { ...person, email };
}

export function ProjectEditor({ project, user, store, manager, displayedStage, onClose, onSave }) {
  const [draft, setDraft] = useState(() => project ? { ...project } : {
    id: uid('project'), projectKey: uid('project-key'), title: '', measurementArea: '', description: '',
    ownerName: 'Unassigned', ownerEmail: '', ownerKey: '', priority: 'Medium', health: 'Needs Review',
    status: 'Planned', currentStageKey: '', targetFinish: '', percentComplete: 0,
    nextMilestone: 'Define project plan', nextMilestoneDate: '',
  });
  const [assignment, setAssignment] = useState(null);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState('');
  const set = key => event => setDraft(row => ({ ...row, [key]: event.target.value }));
  useEffect(() => {
    if (assignment?.state !== 'searching' && assignment?.state !== 'verifying') return;
    let active = true;
    const query = assignment.email.trim();
    const finish = result => {
      if (active) setAssignment(current => current === assignment ? result : current);
    };
    async function lookup() {
      let matches = [];
      try {
        let candidate = assignment.candidate;
        if (!candidate) {
          matches = await store.searchPeople(query);
          if (!active) return;
          const exact = matches.filter(person => isOwnedByUser({ ownerEmail: query }, person));
          // A partially typed domain can already have valid email syntax.
          if (!isEmail(query) || (!exact.length && matches.length)) {
            finish({ state: 'choose', email: assignment.email, matches });
            return;
          }
          if (exact.length > 1) throw new Error('More than one account matches this email. Select the intended account or enter its unique sign-in email.');
          candidate = exact[0];
        }
        const person = await verifyEngineer(store, candidate, assignment.candidate ? '' : query);
        if (!active) return;
        setDraft(current => {
          const team = projectEngineers(current);
          const existing = team.findIndex(engineer => sameEngineer(engineer, person));
          const verified = { name: person.title, email: person.email, loginName: person.loginName };
          return { ...current, ...engineerFields(existing < 0 ? [...team, verified] : team.map((engineer, index) => index === existing ? verified : engineer)) };
        });
        finish({ state: 'added', email: '', message: `${displayName(person.title)} is on the project team.` });
      } catch (error) {
        finish({ state: 'error', email: assignment.email, matches, error: error.message });
      }
    }
    let timer;
    if (assignment.candidate) lookup();
    else timer = setTimeout(lookup, 300);
    return () => { active = false; clearTimeout(timer); };
  }, [assignment, store]);

  function changeEmail(value) {
    setSaveError('');
    const email = value.trim();
    if (!email) setAssignment({ state: 'empty', email: '' });
    else if (email.length >= 2) setAssignment({ state: 'searching', email: value });
    else setAssignment({ state: 'typing', email: value });
  }

  const unresolved = assignment && assignment.state !== 'empty' && assignment.state !== 'added';
  const automaticStage = project && !project.currentStageKey && workflowData.phases.find(phase => phase.key === displayedStage);
  const engineers = projectEngineers(draft);
  const statusOptions = [...new Set([...PROJECT_STATUSES, ...(draft.status ? [draft.status] : [])])];

  async function save() {
    if (saving || unresolved || !draft.title?.trim()) return;
    setSaving(true);
    setSaveError('');
    const row = manager ? { ...draft, ...engineerFields(engineers) } : draft;
    const teamChanged = manager && JSON.stringify(engineers) !== JSON.stringify(projectEngineers(project || {}));
    try {
      if (teamChanged) await onSave(row, engineers.map(engineer => engineer.loginName || engineer.email));
      else await onSave(row);
    }
    catch (error) { setSaveError(error.message || 'Could not save the project. Try again.'); }
    finally { setSaving(false); }
  }

  const close = () => { if (!saving) onClose(); };
  return <Modal title={project?.id ? 'Edit project' : 'Add modernization project'} onClose={close} actions={<>
    <button className="button secondary" disabled={saving} onClick={close}>Cancel</button>
    <button className="button primary" disabled={saving || !draft.title?.trim() || unresolved} onClick={save}>{saving ? 'Saving…' : 'Save project'}</button>
  </>}>
    <div className="field-grid">
      <Field label="Project name"><input disabled={!manager || saving} value={draft.title || ''} onChange={set('title')} /></Field>
      <Field label="Measurement area"><input disabled={!manager || saving} value={draft.measurementArea || ''} onChange={set('measurementArea')} /></Field>
      <div className="field field-wide project-team-field">
        <div className="project-team-heading"><span>Project engineers</span><span className="badge">{engineers.length} assigned</span></div>
        {engineers.length ? <ul className="project-team-selection" aria-label="Assigned project engineers">{engineers.map((engineer, index) => <li className="project-engineer-chip" key={engineer.loginName || engineer.email || `${engineer.name}-${index}`}>
          <span className="avatar small" aria-hidden="true">{engineerInitials(engineer)}</span>
          <span className="project-engineer-identity"><strong>{displayName(engineer.name || engineer.email || engineer.loginName)}</strong><small>{engineer.email || engineer.loginName || 'Existing assignment'}</small></span>
          {manager && <button type="button" className="icon-button project-engineer-remove" disabled={saving} aria-label={`Remove ${displayName(engineer.name || engineer.email || engineer.loginName)} from project`} title="Remove project engineer" onClick={() => {
            setSaveError('');
            setDraft(current => ({ ...current, ...engineerFields(projectEngineers(current).filter((_, position) => position !== index)) }));
            setAssignment(current => current && !['empty', 'added'].includes(current.state) ? current : { state: 'added', email: '', message: `${displayName(engineer.name || engineer.email || engineer.loginName)} removed from the project team.` });
            document.getElementById('project-engineer-email')?.focus();
          }}><Icon name="close" size={14} /></button>}
        </li>)}</ul> : <p className="project-team-empty">No project engineers assigned.</p>}
        <div className="project-team-search">
          <label htmlFor="project-engineer-email">Add project engineer</label>
          <PeopleCombobox id="project-engineer-email" value={assignment?.email || ''} disabled={!manager || saving}
            placeholder={manager ? 'Search by name or email' : 'Managed by a Tracker Manager'} listLabel="Matching SharePoint users"
            matches={(assignment?.matches || []).filter(person => !engineers.some(engineer => sameEngineer(engineer, person)))} emptyMessage={assignment?.state === 'choose' ? 'No additional people found. Try another name or email.' : ''}
            onQuery={changeEmail} onSelect={candidate => {
              setSaveError('');
              setAssignment({ state: 'verifying', candidate, email: engineerEmail(candidate) || assignment.email });
            }}
            renderOption={person => <><strong>{person.title}</strong><small>{engineerEmail(person) || person.loginName}</small>{person.detail && <small>{person.detail}</small>}</>}
            action={manager && <button type="button" disabled={saving || engineers.some(engineer => sameEngineer(engineer, user))} onClick={() => {
              setSaveError('');
              setAssignment({ state: 'verifying', candidate: user, email: engineerEmail(user) });
            }}>Assign me</button>}
          />
        </div>
        <div className="project-team-feedback">
          {!assignment || assignment.state === 'empty' ? <small>{manager ? 'Add one or more engineers. Each person can access and update this project’s progress.' : 'A Tracker Manager can change this project’s team.'}</small> : null}
          {assignment?.state === 'added' && <small role="status">{assignment.message}</small>}
          {assignment?.state === 'typing' && <small>Enter at least two characters.</small>}
          {assignment?.state === 'searching' && <small role="status">Searching SharePoint…</small>}
          {assignment?.state === 'verifying' && <small role="status">Verifying account…</small>}
          {assignment?.state === 'error' && <small className="inline-error" role="alert">{assignment.error}</small>}
          {manager && unresolved && <button type="button" className="text-button" disabled={saving} onClick={() => setAssignment({ state: 'empty', email: '' })}>Clear search</button>}
        </div>
      </div>
      <Field label="Priority"><select disabled={!manager || saving} value={draft.priority || 'Medium'} onChange={set('priority')}>{PRIORITIES.map(item => <option key={item}>{item}</option>)}</select></Field>
      <Field label="Health"><select disabled={saving} value={draft.health || 'Needs Review'} onChange={set('health')}>{HEALTHS.map(item => <option key={item}>{item}</option>)}</select></Field>
      <Field label="Project status"><select disabled={saving} value={draft.status || 'Planned'} onChange={set('status')}>{statusOptions.map(item => <option key={item}>{item}</option>)}</select></Field>
      <Field label="Current stage"><select disabled={saving} value={draft.currentStageKey || ''} onChange={set('currentStageKey')}><option value="">Follow work breakdown{automaticStage ? ` (${automaticStage.name})` : ''}</option>{workflowData.phases.map(phase => <option value={phase.key} key={phase.key}>{phase.name}</option>)}</select></Field>
      <Field label="Target completion"><input disabled={!manager || saving} type="date" value={draft.targetFinish || ''} onChange={set('targetFinish')} /></Field>
      <Field label="Description" wide><textarea disabled={!manager || saving} rows="4" value={draft.description || ''} onChange={set('description')} /></Field>
    </div>
    {saveError && <p role="alert" className="inline-error">{saveError}</p>}
  </Modal>;
}

function TaskEditor({ task, manager, canDeleteDocuments = false, readOnly = false, store, saved = true, users = [], onClose, onSave, onDelete }) {
  const [draft, setDraft] = useState({ ...task, organization: normalizeOrganization(task.organization), assignedDate: task.assignedDate || (!saved ? todayIso() : '') });
  const [error, setError] = useState('');
  const [saving, setSaving] = useState(false);
  const [persisted, setPersisted] = useState(saved);
  const [pendingFiles, setPendingFiles] = useState([]);
  const [ownerUnresolved, setOwnerUnresolved] = useState(false);
  const set = (key) => (event) => setDraft((row) => ({ ...row, [key]: event.target.value }));
  const status = draft.status === 'Not Applicable' ? 'Not Required' : draft.status;
  function queueFiles(event) {
    const files = [...event.target.files]; event.target.value = '';
    const next = [...pendingFiles], failures = [];
    for (const file of files) {
      try { validateAttachment(file); if (next.some((row) => row.name.toLowerCase() === file.name.toLowerCase())) throw new Error('Already selected.'); next.push(file); }
      catch (caught) { failures.push(`${file.name}: ${caught.message}`); }
    }
    setPendingFiles(next); setError(failures.join('\n'));
  }
  async function save() {
    if (saving || ownerUnresolved) return;
    setError('');
    try {
      validateTask({ ...draft, status }); setSaving(true);
      const result = await onSave({ ...draft, status });
      if (!result) { setError('Task could not be saved. Please try again.'); return; }
      const record = typeof result === 'object' ? result : draft;
      setDraft(record); setPersisted(true);
      const failed = [], messages = [];
      for (const file of pendingFiles) {
        try { await store.addTaskAttachment(record.id, file); }
        catch (caught) { failed.push(file); messages.push(`${file.name}: ${caught.message}`); }
      }
      setPendingFiles(failed);
      if (failed.length) setError(`Task saved. ${failed.length} document(s) could not upload. Save again to retry.\n${messages.join('\n')}`);
      else onClose();
    } catch (caught) { setError(caught.message); }
    finally { setSaving(false); }
  }
  return <Modal title={readOnly ? 'Task details' : persisted ? 'Update task' : 'Add task'} subtitle={workflowData.phases.find((phase) => phase.key === task.phaseKey)?.name} onClose={saving ? () => {} : onClose} actions={<>{manager && persisted && onDelete && <button className="button secondary" disabled={saving} onClick={async () => { if (await onDelete(draft)) onClose(); }}>Delete task</button>}<button className="button secondary" disabled={saving} onClick={onClose}>{readOnly ? 'Close' : 'Cancel'}</button>{!readOnly && <button className="button primary" disabled={saving || ownerUnresolved} onClick={save}>{saving ? 'Saving task and documents…' : 'Save task'}</button>}</>}>
    {error && <p className="inline-error" role="alert">{error}</p>}
    <fieldset className="field-grid task-fields" disabled={readOnly || saving}>
      <Field label="Task name" wide><input aria-label="Task name" disabled={!manager} value={draft.title} onChange={set('title')} /></Field>
      <Field label="Status"><select value={status} onChange={set('status')}>{TASK_STATUSES.map((item) => <option key={item}>{item}</option>)}</select></Field>
      <Field label="Organization"><select value={draft.organization} onChange={event => setDraft(row => ({ ...row, organization: event.target.value, subAssignment: '' }))}>{ORGANIZATIONS.map((item) => <option key={item}>{item}</option>)}</select></Field>
      {draft.organization === METCAL_ORGANIZATION && <Field label="METCAL responsibility"><select aria-label="METCAL responsibility" value={draft.subAssignment || ''} onChange={set('subAssignment')}><option value="">Not yet assigned</option>{METCAL_TEAMS.map(team => <option key={team}>{team}</option>)}</select></Field>}
      <SavedUserPicker label="Task owner" users={users} disabled={readOnly || saving} value={draft.ownerName} onQuery={(name) => { setOwnerUnresolved(!!name.trim()); setDraft((row) => ({ ...row, ownerName: name, ownerEmail: '', ownerKey: '' })); }} onSelect={(entry) => { setOwnerUnresolved(false); setDraft((row) => ({ ...row, ownerName: entry.title, ownerEmail: entry.email || '', ownerKey: userIdentityKey(entry) })); }} />
      <Field label="Est. Hours"><input aria-label="Est. Hours" type="number" min="0" step="any" value={draft.estimatedHours ?? ''} onChange={set('estimatedHours')} /></Field>
      <Field label="Original due date"><input aria-label="Original due date" type="date" disabled={!manager} value={draft.dueDate || ''} onChange={set('dueDate')} />{!manager && <small className="field-hint">Set by a manager</small>}</Field>
      <Field label="Assigned / creation date"><input aria-label="Assigned / creation date" disabled={!manager} type="date" value={draft.assignedDate || ''} onChange={set('assignedDate')} />{!manager && <small className="field-hint">Set by a manager</small>}</Field>
      <Field label="Deferred date"><input aria-label="Deferred date" type="date" value={draft.deferredDate || ''} onChange={set('deferredDate')} /></Field>
      <Field label="Deferral justification" wide><textarea value={draft.deferredJustification || ''} onChange={set('deferredJustification')} /></Field>
      {status === 'Not Required' && <Field label="Not required justification" wide><textarea value={draft.notRequiredJustification || ''} onChange={set('notRequiredJustification')} /></Field>}
      {status === 'Blocked' && <Field label="Blocked reason" wide><textarea value={draft.blockedReason || ''} onChange={set('blockedReason')} /></Field>}
      <Field label="Notes" wide><AutoSizeTextarea aria-label="Notes" value={draft.notes || ''} onChange={set('notes')} /></Field>
    </fieldset>
    {persisted && store && <TaskDocuments taskId={draft.id} store={store} readOnly={readOnly || saving || pendingFiles.length > 0} canDelete={canDeleteDocuments && !saving} refreshKey={saving} />}
    {!readOnly && (!persisted || pendingFiles.length > 0) && <section className="task-documents queued-documents"><h3>Documents</h3><FilePicker label="Attach documents" hint="Up to 50 MB per file. Documents upload when you save the task." ariaLabel="Attach documents" multiple disabled={saving} onChange={queueFiles} /><ul className="document-list">{pendingFiles.map((file) => <li key={file.name}><div className="document-name"><strong>{file.name}</strong><small>{(file.size / 1024 / 1024).toFixed(2)} MB · Uploads when you save</small></div><button type="button" className="text-button" disabled={saving} onClick={() => setPendingFiles((rows) => rows.filter((row) => row !== file))}>Remove</button></li>)}</ul></section>}
  </Modal>;
}

function DirectoryUserActions({ entry, currentUser, busy, onUpdate, onDelete }) {
  const { menuOpen, setMenuOpen, menuRef } = useDismissibleMenu();
  const self = isOwnedByUser({ ownerKey: entry.loginName, ownerEmail: entry.email }, currentUser);
  return <div className="directory-actions project-menu-wrap" ref={menuRef}><button type="button" className="button secondary" disabled={busy} aria-expanded={menuOpen} onClick={() => setMenuOpen((value) => !value)}>Edit</button>{menuOpen && <div className="project-menu-popover"><button type="button" onClick={() => { setMenuOpen(false); onUpdate(entry); }}>Update User</button><button type="button" disabled={self || busy} title={self ? 'You cannot delete your own account' : 'Remove from tracker directory'} onClick={() => { setMenuOpen(false); onDelete(entry); }}>Delete User</button></div>}</div>;
}

function UserDirectory({ users, currentUser, onDelete, store, config, localPreview, manager, testingEnabled, onActivate, onSave }) {
  const [userSearch, setUserSearch] = useState('');
  const filteredUsers = users.filter((entry) => `${entry.title} ${entry.email || ''} ${entry.role}`.toLowerCase().includes(userSearch.trim().toLowerCase())).sort((a, b) => a.title.localeCompare(b.title));
  const [password, setPassword] = useState('');
  const [testingManagerOpen, setTestingManagerOpen] = useState(false);
  const [activationError, setActivationError] = useState('');
  const [saveError, setSaveError] = useState('');
  const [saveMessage, setSaveMessage] = useState('');
  useEffect(() => {
    if (manager || !testingEnabled) return undefined;
    const openTestingManager = (event) => {
      if (!event.ctrlKey || event.altKey || event.shiftKey || event.key.toLowerCase() !== 'm') return;
      event.preventDefault();
      setTestingManagerOpen(true);
    };
    window.addEventListener('keydown', openTestingManager);
    return () => window.removeEventListener('keydown', openTestingManager);
  }, [manager, testingEnabled]);
  async function activate(event) {
    event.preventDefault();
    if (saving || !password) return;
    setSaving(true); setActivationError('');
    try { await onActivate(password); setPassword(''); }
    catch (error) { setActivationError(error.message || 'Manager activation failed. Please try again.'); }
    finally { setSaving(false); }
  }
  const empty = { title: '', email: '', loginName: '', role: 'Viewer' };
  const [draft, setDraft] = useState(empty);
  const [saving, setSaving] = useState(false);
  async function saveDirectoryUser(event) {
    event?.preventDefault();
    if (saving) return;
    setSaveError(''); setSaveMessage('');
    if (!draft.title.trim() || !(draft.loginName?.trim() || draft.email.trim())) {
      setSaveError('Name and an email or existing SharePoint identity are required.'); return;
    }
    setSaving(true);
    try {
      await onSave({ ...draft, id: draft.id || uid('user'), email: draft.email.trim(), loginName: draft.loginName?.trim() || draft.email.trim() });
      setSaveMessage(`${displayName(draft.title)} saved as ${draft.role}.`); setDraft(empty);
    } catch (error) { setSaveError(error.message || 'User could not be saved. Please try again.'); }
    finally { setSaving(false); }
  }
  async function saveInvitee(person) {
    setSaveError('');
    setSaveMessage('');
    const existing = users.find(entry => isOwnedByUser({ ownerKey: entry.loginName, ownerEmail: entry.email }, person));
    if (existing) {
      if (existing.role !== person.role) throw new Error('This person already has a different tracker role. Use Edit to change it first.');
      return existing;
    }
    return onSave({ ...person, id: uid('user') });
  }
  async function deleteDirectoryUser(entry) {
    if (saving || isOwnedByUser({ ownerKey: entry.loginName, ownerEmail: entry.email }, currentUser)) return;
    setSaving(true); setSaveError(''); setSaveMessage('');
    try { await onDelete(entry); setDraft(empty); setSaveMessage(`${displayName(entry.title)} removed from the tracker directory.`); }
    catch (error) { setSaveError(error.message || 'User could not be deleted. Please try again.'); }
    finally { setSaving(false); }
  }
  return <section className="page-stack"><h1>Users and managers</h1><p>Signed-in users and their application roles.</p>{manager && <TrackerAccessSetup store={store} />}{manager && <PeopleInvite store={store} config={config} localPreview={localPreview} onSave={saveInvitee} />}{!manager && testingEnabled && testingManagerOpen && <form className="panel directory-form" onSubmit={activate}><h2>Testing manager access</h2><p>Testing only: enter the shared testing password to grant your signed-in account manager access.</p><Field label="Testing password"><input type="password" autoComplete="off" value={password} onChange={(event) => setPassword(event.target.value)} /></Field>{activationError && <p role="alert" className="inline-error">{activationError}</p>}<button type="button" className="button primary" onClick={activate} disabled={saving || !password}>{saving ? 'Saving and verifying role…' : 'Enable manager access'}</button>{saving && <p role="status">Please wait while your manager role is saved and verified.</p>}</form>}{manager && draft.id && <form className="panel directory-form" noValidate onSubmit={saveDirectoryUser}>
    {['title', 'email'].map((key) => <Field key={key} label={{ title: 'Name', email: 'Email', loginName: 'SharePoint login' }[key]}><input type={key === 'email' ? 'email' : 'text'} required={key === 'title' || !draft.loginName?.trim()} disabled={saving} value={draft[key] || ''} onChange={(event) => setDraft((row) => ({ ...row, [key]: event.target.value }))} /></Field>)}
    <Field label="Role"><select aria-label="Role" disabled={saving} value={draft.role} onChange={(event) => setDraft((row) => ({ ...row, role: event.target.value }))}>{ROLES.map((role) => <option key={role}>{role}</option>)}</select></Field><button type="button" className="button primary" onClick={saveDirectoryUser} disabled={saving}>{saving ? 'Saving and verifying role…' : 'Save user'}</button>{draft.id && <button type="button" className="button secondary" disabled={saving} onClick={() => { setDraft(empty); setSaveError(''); setSaveMessage(''); }}>Cancel edit</button>}
    </form>}{saveError && <p role="alert" className="inline-error">{saveError}</p>}{saveMessage && <p role="status">{saveMessage}</p>}<section className="panel directory-list-panel"><div className="directory-list-heading"><h2>Users <small>({users.length})</small></h2><input type="search" aria-label="Search users" placeholder="Search name, email or role" value={userSearch} onChange={(event) => setUserSearch(event.target.value)} /></div><div className="directory-list" role="region" aria-label="User directory" tabIndex={0}>{!filteredUsers.length && <p className="empty-state">No users match your search.</p>}{filteredUsers.map((entry) => <article className="directory-row" key={entry.id}><div className="directory-identity"><strong>{displayName(entry.title)}</strong><span>{entry.email || 'Email not available'}</span></div><Badge>{entry.role}</Badge>{manager && <DirectoryUserActions entry={entry} currentUser={currentUser} busy={saving} onUpdate={(entry) => { setDraft({ ...empty, ...entry, email: entry.email || '' }); setSaveError(''); setSaveMessage(''); }} onDelete={deleteDirectoryUser} />}</article>)}</div></section></section>;
}

function ProjectDrawer({ project, user, manager, readOnly, store, users, onDeleteTask, onProgressMode, tasks, updates, risks, onClose, onEdit, onSaveTask, onToggleTask, onSetPhaseRequired, onAddUpdate, onEditUpdate, onDeleteUpdate, onRefreshUpdates, onDelete }) {
  const [tab, setTab] = useState('overview');
  const [phaseReason, setPhaseReason] = useState(null);
  const [reason, setReason] = useState('');
  const managesProject = canManageProject(project, user, users);
  const [projectAttachments, setProjectAttachments] = useState([]);
  const [documentRevision, setDocumentRevision] = useState(0);
  const { menuOpen, setMenuOpen, menuRef } = useDismissibleMenu();
  const [progressMode, setProgressMode] = useState(project.progressMode || 'phases');
  const [taskEditor, setTaskEditor] = useState(null);
  const projectTasks = tasks.filter((task) => task.projectKey === project.projectKey).sort((a, b) => phaseIndex(a.phaseKey) - phaseIndex(b.phaseKey) || (a.order || 0) - (b.order || 0));
  const progress = projectProgress(projectTasks, progressMode);
  const currentTask = projectTasks.find((task) => !isClosedTask(task));
  const markers = progressMode === 'tasks' ? projectTasks.map((task) => ({ key: task.id, label: `${task.title} — ${task.status}`, done: task.status === 'Complete', current: task.id === currentTask?.id })) : workflowData.phases.map((item) => { const rows = projectTasks.filter((task) => task.phaseKey === item.key); return { key: item.key, label: item.name, done: rows.length > 0 && rows.every(isClosedTask), current: item.key === project.currentStageKey }; });
  const projectUpdates = updates.filter((item) => item.projectKey === project.projectKey).sort((a, b) => String(b.entryDate).localeCompare(String(a.entryDate)));
  const projectRisks = risks.filter((item) => item.projectKey === project.projectKey);
  const phase = workflowData.phases.find((item) => item.key === project.currentStageKey);
  useEffect(() => {
    let active = true;
    store.listProjectAttachments(project.projectKey).then((rows) => { if (active) setProjectAttachments(rows); }).catch(() => { if (active) setProjectAttachments([]); });
    return () => { active = false; };
  }, [project.projectKey, store, taskEditor, documentRevision]);
  return <div className="drawer-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><aside className="project-drawer" role="dialog" aria-modal="true"><header className="drawer-header"><div className="drawer-title-row"><div><span className="section-kicker">{project.measurementArea}</span><h2>{project.title}</h2><div className="drawer-badges"><Badge tone={statusTone(project.health)} dot>{project.health}</Badge><Badge>{project.organization}</Badge><Badge>{project.priority} priority</Badge></div></div><div className="drawer-header-actions">{managesProject && <div ref={menuRef} className="project-menu-wrap"><button className="icon-button" aria-label="Project settings" aria-expanded={menuOpen} onClick={() => setMenuOpen(!menuOpen)}><Icon name="more" className="vertical-more" /></button>{menuOpen && <div className="project-menu-popover"><button onClick={() => { setMenuOpen(false); onEdit(project); }}>Edit project</button>{manager && <button onClick={() => { setMenuOpen(false); onDelete(project); }}>Delete project</button>}</div>}</div>}<button className="icon-button" onClick={onClose} aria-label="Close"><Icon name="close" /></button></div></div>
    <div className="drawer-progress"><div className="progress-caption"><span>{progress.percentComplete}% complete</span><strong>{progressMode === 'tasks' ? (currentTask?.title || (projectTasks.length ? 'No outstanding tasks' : 'No tasks yet')) : phase?.name}</strong><ProgressChoice value={progressMode} onChange={(mode) => { setProgressMode(mode); if (managesProject) onProgressMode(project.id, mode); }} /></div><Progress value={progress.percentComplete} /><p className="progress-detail">{progress.progressDetail}</p></div>
    <div className="marker-scroll"><div className={`mini-pipeline ${progressMode === 'tasks' ? 'task-markers' : ''}`} style={{ gridTemplateColumns: `repeat(${Math.max(1, markers.length)}, minmax(32px, 1fr))`, minWidth: progressMode === 'tasks' ? `${markers.length * 37}px` : undefined }} aria-label={progressMode === 'tasks' ? 'Task progress markers' : 'Phase progress markers'}>{markers.map((item, index) => <span key={item.key} className={item.done ? 'done' : item.current ? 'current' : ''} title={item.label} aria-label={item.label} aria-current={item.current && !item.done ? 'step' : undefined}>{index + 1}</span>)}</div></div>
    <nav className="drawer-tabs">{[['overview', 'Overview'], ['tasks', `Work breakdown (${projectTasks.length})`], ['updates', `Updates (${projectUpdates.length + projectRisks.length})`]].map(([key, label]) => <button className={tab === key ? 'active' : ''} key={key} onClick={() => setTab(key)}>{label}</button>)}</nav></header>
    <div className="drawer-content">
      {tab === 'overview' && <div className="drawer-section-stack"><section className="detail-grid"><div className="project-team-detail"><span>Project engineers</span><ProjectTeam project={project} detailed /></div><div><span>Project status</span><strong>{project.status || 'Planned'}</strong></div><div><span>Target completion</span><strong>{displayDate(project.targetFinish)}</strong></div><div><span>Next milestone</span><strong>{project.nextMilestone}</strong><small>{displayDate(project.nextMilestoneDate)}</small></div><div><span>Open actions</span><strong>{projectTasks.filter((task) => !isClosedTask(task)).length}</strong><small>{project.blockedCount} blocked</small></div></section><section className="detail-block"><h3>Upcoming work</h3><div className="upcoming-list">{projectTasks.filter((task) => !isClosedTask(task)).slice(0, 5).map((task) => <button key={task.id} onClick={() => setTaskEditor(task)}><div><strong>{task.title}</strong><small>{workflowData.phases.find((item) => item.key === task.phaseKey)?.short}</small></div><Badge tone={statusTone(task.status)}>{task.status}</Badge></button>)}</div></section><section className="detail-block project-documents"><TaskDocuments projectKey={project.projectKey} store={store} readOnly canDelete={managesProject} refreshKey={taskEditor?.id} onChange={() => setDocumentRevision((value) => value + 1)} /></section></div>}
      {tab === 'tasks' && <div className="phase-task-list">{workflowData.phases.map((stage) => { const stageTasks = projectTasks.filter((task) => task.phaseKey === stage.key); const done = stageTasks.filter(isClosedTask).length; const stageDocuments = projectAttachments.filter((file) => stageTasks.some((task) => task.id === file.taskId)); const documentSummary = stageDocuments.map((file) => `${file.taskTitle}: ${file.name}`).join("\n"); const allNotRequired = stageTasks.length > 0 && stageTasks.every((task) => ['Not Required', 'Not Applicable'].includes(task.status)); return <section key={stage.key} className={`task-stage ${allNotRequired ? 'phase-not-required' : ''}`}><header><div><span>{String(phaseIndex(stage.key) + 1).padStart(2, '0')}</span><strong>{stage.name}</strong></div><div className="phase-header-actions">{!!stageDocuments.length && <span className="phase-document-indicator" tabIndex={0} title={documentSummary} aria-label={`Documents uploaded to ${stage.name}: ${documentSummary}`}><Icon name="paperclip" size={15} /></span>}<Badge>{done}/{stageTasks.length}</Badge>{manager && <button className="icon-button add-task-button" aria-label={`Add task to ${stage.name}`} title={`Add task to ${stage.name}`} onClick={() => setTaskEditor({ id: uid('task'), projectKey: project.projectKey, title: '', phaseKey: stage.key, order: Math.max(0, ...projectTasks.map((task) => task.order || 0)) + 1, status: 'Not Started', ownerName: project.ownerName, ownerKey: project.ownerKey, ownerEmail: project.ownerEmail, dueDate: '' })}><Icon name="plus" size={16} /></button>}{managesProject && <button type="button" className={`phase-required-button ${allNotRequired ? 'active' : ''}`} disabled={!stageTasks.length} aria-label={`${allNotRequired ? 'Restore' : 'Mark'} ${stage.name} ${allNotRequired ? '' : 'not required'}`.trim()} onClick={() => { if (allNotRequired) onSetPhaseRequired(project, stage.key, false); else { setPhaseReason(stage.key); setReason(''); } }}><Icon name={allNotRequired ? 'refresh' : 'ban'} size={13} />{allNotRequired ? 'Restore phase' : 'Not required'}</button>}</div></header><div>{stageTasks.map((task) => { const resolved = isClosedTask(task); return <div key={task.id} className="phase-task" role="button" tabIndex="0" onClick={() => setTaskEditor(task)} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); setTaskEditor(task); } }}><button disabled={readOnly || !managesProject} className={`task-check ${resolved ? 'checked' : ''}`} aria-label={`${resolved ? 'Reopen' : 'Complete'} ${task.title}`} onClick={(event) => { event.stopPropagation(); onToggleTask(task); }}>{resolved && <Icon name="check" size={13} />}</button><div><strong>{task.title}</strong>{task.dataIssue && <small className="issue-text"><Icon name="alert" size={12} />{task.dataIssue}</small>}</div><Badge tone={isOverdue(task) ? 'bad' : statusTone(task.status)}>{task.status === 'Not Applicable' ? 'Not Required' : task.status}</Badge><span className="task-date">{isClosedTask(task) ? 'Resolved' : compactDate(task.dueDate)}</span></div>; })}</div></section>; })}</div>}
      {tab === 'updates' && <ProjectUpdates key={project.projectKey} project={project} tasks={projectTasks} updates={projectUpdates} risks={projectRisks} store={store} canPost={managesProject && !readOnly} onPost={onAddUpdate} onEdit={onEditUpdate} onDelete={onDeleteUpdate} onRefresh={onRefreshUpdates} onTask={setTaskEditor} />}
    </div>
    {phaseReason && <Modal title="Why is this phase not required?" onClose={() => setPhaseReason(null)} actions={<><button className="button secondary" onClick={() => setPhaseReason(null)}>Cancel</button><button className="button primary" disabled={!reason.trim()} onClick={() => { onSetPhaseRequired(project, phaseReason, true, reason.trim()); setPhaseReason(null); }}>Apply to phase</button></>}><Field label="Phase justification"><textarea value={reason} onChange={(event) => setReason(event.target.value)} /></Field></Modal>}
    {taskEditor && <TaskEditor task={taskEditor} manager={manager} canDeleteDocuments={managesProject} readOnly={readOnly} store={store} saved={projectTasks.some((row) => row.id === taskEditor.id)} users={users} onClose={() => setTaskEditor(null)} onSave={onSaveTask} onDelete={manager ? onDeleteTask : undefined} />}
  </aside></div>;
}

function Toast({ toast, onClose }) { if (!toast) return null; return <div className={`toast toast-${toast.tone || 'good'}`} role={toast.tone === 'bad' ? 'alert' : 'status'} aria-live={toast.tone === 'bad' ? 'assertive' : 'polite'}><Icon name={toast.tone === 'bad' ? 'alert' : toast.tone === 'info' ? 'clock' : 'check'} /><span>{toast.message}</span><button type="button" aria-label="Dismiss notification" onClick={onClose}><Icon name="close" size={13} /></button></div>; }

export function App() {
  const [register, setRegister] = useState({ progressMode: 'phases', health: '', organization: '', search: '', sortKey: 'project', sortDirection: 'asc' });
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
  const [attentionOrganization, setAttentionOrganization] = useState('');
  const [attentionControls, setAttentionControls] = useState({ groupBy: 'status', search: '', sort: 'due' });
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
      if (!readiness.ready) await repo.store.provision({ allowCompatibleRead: true });
      await repo.store.registerCurrentUser();
      const loaded = await repo.store.load();
      setData(loaded);
      setView(canViewPortfolio(currentUser, loaded.users) ? 'overview' : 'my-work');
      // Metadata maintenance is owner-only and does not block reading the app.
      repo.store.maintainListNames?.().catch(caught => setToast(`List names could not be updated: ${caught.message}`));
    } catch (caught) { setError(caught.message || 'Could not load portfolio data.'); }
    finally { setLoading(false); }
  }

  const manager = isManager(user, data.users);
  const portfolioAccess = canViewPortfolio(user, data.users);
  const enriched = useMemo(() => data.projects.map((project) => enrichProject(project, data.tasks)), [data.projects, data.tasks]);
  const openProject = enriched.find((project) => project.id === openProjectId);

  async function saveProject(project, assignedLogin) {
    const creating = !projectEditor?.id;
    try {
      const persisted = creating && data.projects.find(row => row.id === project.id);
      const saved = await repo.store.saveProject(persisted?.spId ? { ...project, spId: persisted.spId } : project, assignedLogin);
      if (creating) {
        const template = createStarterTasks(saved.projectKey, saved);
        const existingIds = new Set(data.tasks.map(task => task.id));
        const missing = template.filter(task => !existingIds.has(task.id));
        for (let index = 0; index < missing.length; index += 10) {
          const writes = await Promise.allSettled(missing.slice(index, index + 10).map(task => repo.store.saveTask(task)));
          const failed = writes.find(write => write.status === 'rejected');
          if (failed) throw failed.reason;
        }
      }
      setData(await repo.store.load());
      setOpenProjectId(saved.id);
      setProjectEditor(null);
      setToast({ message: `${saved.title} is ready.` });
    } catch (error) {
      // A group, project or starter task may already be saved. Refresh the list
      // for a retry while keeping the same editor, record IDs and entered values.
      try { setData(await repo.store.load()); } catch { /* keep the current view */ }
      throw error;
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
      return saved;
    } catch (caught) { setToast({ tone: 'bad', message: caught.message }); return false; }
  }

  async function deleteTask(task) {
    try { await repo.store.recycle('tasks', task.spId, task.id); setData((state) => ({ ...state, tasks: state.tasks.filter((row) => row.id !== task.id) })); return true; }
    catch (caught) { setToast({ tone: 'bad', message: caught.message }); return false; }
  }

  async function saveUser(row) {
    try { const saved = await repo.store.saveUser(row); setData((state) => ({ ...state, users: (state.users || []).some((entry) => entry.id === saved.id) ? state.users.map((entry) => entry.id === saved.id ? saved : entry) : [...(state.users || []), saved] })); setToast({ message: 'User saved.' }); return saved; }
    catch (caught) { setToast({ tone: 'bad', message: caught.message }); throw caught; }
  }

  async function deleteUser(row) {
    await repo.store.recycle('users', row.spId, row.id);
    setData((state) => ({ ...state, users: state.users.filter((entry) => entry.id !== row.id) }));
  }

  async function addUpdate(project, draft) {
    const saved = await repo.store.saveUpdate({ id: uid('update'), projectKey: project.projectKey, ...draft });
    setData(state => ({ ...state, updates: [saved, ...state.updates] }));
    return saved;
  }

  async function editUpdate(item, changes) {
    const saved = item.legacyRisk ? await repo.store.editRisk({ ...item, ...changes }) : await repo.store.editUpdate({ ...item, ...changes });
    const collection = item.legacyRisk ? 'risks' : 'updates';
    setData(state => ({ ...state, [collection]: state[collection].map(row => row.id === saved.id ? saved : row) }));
    await refreshUpdates();
    return saved;
  }
  async function refreshUpdates() {
    try {
      const fresh = await repo.store.load();
      setData(state => ({ ...state, updates: fresh.updates, risks: fresh.risks }));
    } catch { setToast({ tone: 'info', message: 'History could not refresh. Reload the tracker before editing an update.' }); }
  }
  async function deleteUpdate(item) {
    const collection = item.legacyRisk ? 'risks' : 'updates';
    if (item.legacyRisk) await repo.store.deleteRisk(item);
    else await repo.store.deleteUpdate(item);
    setData(state => ({ ...state, [collection]: state[collection].filter(row => row.id !== item.id) }));
  }

  async function toggleTaskComplete(task) {
    await saveTask({ ...task, status: isClosedTask(task) ? 'Not Started' : 'Complete' });
  }

  async function setPhaseRequired(project, phaseKey, notRequired, justification = '') {
    const source = data;
    const affected = source.tasks.filter((task) => task.projectKey === project.projectKey && task.phaseKey === phaseKey);
    if (!affected.length) return;
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
    if (repo.store.scopedAccess && !manager) { setToast({ tone: 'bad', message: 'Only managers can delete a project and revoke its engineer access.' }); return; }
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
      const fresh = await repo.store.load();
      const exportProjects = fresh.projects.map((project) => enrichProject(project, fresh.tasks));
      const query = register.search.trim().toLowerCase();
      let scope = exportProjects;
      if (view === 'overview') scope = scope.filter((project) => (!phaseFilter.length || phaseFilter.includes(project.currentStageKey)) && (!register.health || project.health === register.health) && (!register.organization || project.tasks.some((task) => !isClosedTask(task) && task.organization === register.organization)) && (!query || [project.title, project.measurementArea, projectTeamSearch(project), project.organization, project.nextMilestone, project.health].some((value) => String(value || '').toLowerCase().includes(query)))).map((project) => ({ ...project, progressMode: register.progressMode, ...projectProgress(project.tasks, register.progressMode) }));
      if (view === 'my-work') scope = scope.filter((project) => isOwnedByUser(project, user) || project.tasks.some((task) => isOwnedByUser(task, user) && !isClosedTask(task)));
      const attentionTaskIds = view === 'attention' ? new Set(attentionRows(exportProjects, attentionOrganization, attentionControls.search, attentionControls.sort, attentionControls.responsibility).map(({ task }) => task.id)) : null;
      if (attentionTaskIds) scope = scope.filter((project) => project.tasks.some((task) => attentionTaskIds.has(task.id)));
      await downloadPortfolioWorkbook({
        projects: scope,
        tasks: fresh.tasks,
        attentionTaskIds,
        updates: fresh.updates,
        risks: fresh.risks,
        phases: workflowData.phases,
        glossary: fresh.acronyms,
        user,
        sourceLabel: `${view === 'overview' ? 'Filtered portfolio' : view} · ${repo.mode === 'sharepoint' ? 'SharePoint' : 'Local'} · ${scope.length} projects`,
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
    <aside className="sidebar"><button className="brand" aria-label="Portfolio home" onClick={() => { setOpenProjectId(''); setView('overview'); }}><img className="brand-logo" src={navairSeal} alt="NAVAIR" /><div><strong>MODERNIZATION</strong><span>Project Tracker</span></div></button><nav>{[...NAV.filter(([key]) => portfolioAccess || ['my-work', 'references', 'glossary'].includes(key)), ['users', 'Users and managers', 'tasks']].map(([key, label, icon]) => <button className={view === key ? 'active' : ''} key={key} onClick={() => setView(key)}><Icon name={icon} /><span>{label}</span>{key === 'my-work' && <Badge>{enriched.filter((project) => isOwnedByUser(project, user)).length}</Badge>}</button>)}</nav><div className="sidebar-footer"><div className="sidebar-user"><span className="avatar">{userInitials(user)}</span><div><strong>{displayName(user?.title || 'SharePoint user')}</strong><span>{user?.email || 'Email not available'}</span></div></div></div></aside>
    <div className="main-shell"><header className="topbar"><button className="mobile-brand" aria-label="Portfolio home" onClick={() => { setOpenProjectId(''); setView('overview'); }}><img className="brand-logo" src={navairSeal} alt="NAVAIR" /><strong>MODERNIZATION</strong></button><div className="top-actions"><button className="button secondary export-button" disabled={exporting} onClick={exportWorkbook}><Icon name="download" /> {exporting ? 'Building Excel…' : 'Export Excel'}</button>{manager && <button className="button primary" onClick={() => setProjectEditor({})}><Icon name="plus" /> New project</button>}<Feedback store={repo.store} /><button className="icon-button" onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')} aria-label={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}><Icon name={theme === 'dark' ? 'sun' : 'moon'} /></button></div></header>
      <main>
        {view === 'overview' && <Overview projects={enriched} tasks={data.tasks} risks={data.risks} onOpen={(project) => setOpenProjectId(project.id)} onDelete={manager ? deleteProject : undefined} phaseFilter={phaseFilter} setPhaseFilter={setPhaseFilter} onAttention={(organization) => { setAttentionOrganization(organization); setView('attention'); }} onBoard={() => setView('board')} register={register} setRegister={setRegister} />}
        {view === 'board' && <Board projects={enriched} onOpen={(project) => setOpenProjectId(project.id)} onDelete={manager ? deleteProject : undefined} />}
        {view === 'attention' && <Attention store={repo.store} refreshKey={`${taskEditor?.id || ''}:${openProjectId || ''}`} projects={enriched} organization={attentionOrganization} controls={attentionControls} onControls={setAttentionControls} onTask={setTaskEditor} onOpen={(project) => setOpenProjectId(project.id)} />}
        {view === 'references' && <ReferenceDocuments store={repo.store} readOnly={!manager} />}
        {view === 'users' && <UserDirectory currentUser={user} onDelete={deleteUser} store={repo.store} config={repo.config} localPreview={repo.mode === 'local'} users={data.users || []} manager={manager} testingEnabled={repo.config.testingManagerPassword !== false && repo.config.testingManagerPassword !== ''} onActivate={async (password) => { const saved = await repo.store.activateTestingManager(password); const loaded = await repo.store.load(); if (!isManager(user, loaded.users)) throw new Error('The directory still reports your account as a basic user. Manager access could not be confirmed; please reload and try again.'); setData(loaded); setToast({ message: `Manager access confirmed for ${saved.title}.` }); }} onSave={saveUser} />}
        {view === 'my-work' && <MyWork projects={enriched} tasks={data.tasks} user={user} onOpen={(project) => setOpenProjectId(project.id)} onTask={setTaskEditor} onDelete={manager ? deleteProject : undefined} />}
        {view === 'glossary' && <Glossary manager={manager} acronyms={data.acronyms} onAdd={addAcronym} onDelete={deleteAcronym} />}
      </main>
    </div>
    {openProject && <ProjectDrawer readOnly={!canManageProject(openProject, user, data.users)} store={repo.store} key={openProject.id} project={openProject} user={user} manager={manager} users={data.users || []} onDeleteTask={deleteTask} onProgressMode={async (id, mode) => { try { const saved = await repo.store.saveProgressMode(id, mode); setData((state) => ({ ...state, projects: state.projects.map((row) => row.id === saved.id ? saved : row) })); } catch (error) { setToast({ tone: 'bad', message: error.message }); } }} tasks={data.tasks} updates={data.updates} risks={data.risks} onClose={() => setOpenProjectId('')} onEdit={(project) => setProjectEditor(data.projects.find((row) => row.id === project.id))} onSaveTask={saveTask} onToggleTask={toggleTaskComplete} onSetPhaseRequired={setPhaseRequired} onAddUpdate={addUpdate} onEditUpdate={editUpdate} onDeleteUpdate={deleteUpdate} onRefreshUpdates={refreshUpdates} onDelete={deleteProject} />}
    {projectEditor && (manager || canManageProject(projectEditor, user, data.users)) && <ProjectEditor project={projectEditor.id ? projectEditor : null} user={user} store={repo.store} displayedStage={openProject?.currentStageKey} manager={manager} onClose={() => setProjectEditor(null)} onSave={saveProject} />}
    {taskEditor && <TaskEditor task={taskEditor} manager={manager} onDelete={manager ? deleteTask : undefined} canDeleteDocuments={canManageProject(data.projects.find((project) => project.projectKey === taskEditor.projectKey), user, data.users)} readOnly={!canManageProject(data.projects.find((project) => project.projectKey === taskEditor.projectKey), user, data.users)} store={repo.store} users={data.users || []} onClose={() => setTaskEditor(null)} onSave={saveTask} />}
    <Toast toast={toast} onClose={() => setToast(null)} />
  </div>;
}
