import React, { useMemo, useState } from 'react';
import { TaskDocuments } from './TaskDocuments';
import { displayName } from '../lib/displayName';
import { organizationAttentionLabel, workflowData } from '../data/workflow';
import { Icon } from './Icon';

const statusOrder = ['Blocked', 'In Progress', 'Not Started'];
const closed = new Set(['Complete', 'Not Required', 'Not Applicable']);
const effectiveDue = (task) => task.deferredDate || task.dueDate || '';
const formatHours = (hours) => hours.toLocaleString(undefined, { maximumFractionDigits: 2 });
const date = (value) => value ? new Date(`${value}T12:00:00`).toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' }) : 'Not set';
const overdue = (task) => !!(task.deferredDate || task.dueDate) && (task.deferredDate || task.dueDate) < new Date().toLocaleDateString('en-CA');
const compareMeasurementArea = (a, b) => {
  const left = String(a.measurementArea || '').trim(), right = String(b.measurementArea || '').trim();
  return (left ? 0 : 1) - (right ? 0 : 1) || left.localeCompare(right, undefined, { sensitivity: 'base', numeric: true });
};

export function attentionRows(projects, organization = '', search = '', sort = 'due') {
  const query = search.toLowerCase().trim();
  return projects.flatMap((project) => project.tasks.filter((task) => !closed.has(task.status) && (!organization || task.organization === organization)).map((task) => ({ project, task })))
    .filter(({ project, task }) => !query || [project.title, task.title, task.ownerName, displayName(task.ownerName), task.organization].join(' ').toLowerCase().includes(query))
    .sort((a, b) => {
      if (sort === 'measurement-area') return compareMeasurementArea(a.project, b.project) || a.project.title.localeCompare(b.project.title) || a.task.title.localeCompare(b.task.title);
      const value = ({ task }) => sort === 'title' ? task.title : sort === 'hours' ? -(Number(task.estimatedHours) || 0) : sort === 'owner' ? displayName(task.ownerName || '') || '\uffff' : (effectiveDue(task) || '9999');
      const x = value(a), y = value(b);
      return (typeof x === 'number' ? x - y : x.localeCompare(y)) || a.project.title.localeCompare(b.project.title) || a.task.title.localeCompare(b.task.title);
    });
}

export function attentionGroups(rows, groupBy, sort) {
  const groups = new Map();
  for (const row of rows) {
    const key = groupBy === 'project' ? row.project.projectKey : row.task.status;
    if (!groups.has(key)) groups.set(key, { key, project: row.project, items: [], hours: 0, overdueCount: 0, nextDue: '', firstOwner: '\uffff' });
    const group = groups.get(key);
    group.items.push(row);
    group.hours += Number(row.task.estimatedHours) || 0;
    group.overdueCount += overdue(row.task) ? 1 : 0;
    const due = effectiveDue(row.task);
    if (due && (!group.nextDue || due < group.nextDue)) group.nextDue = due;
    const owner = displayName(row.task.ownerName || '') || '\uffff';
    if (owner.localeCompare(group.firstOwner) < 0) group.firstOwner = owner;
  }
  return [...groups.values()].sort((a, b) => {
    if (groupBy !== 'project') return statusOrder.indexOf(a.key) - statusOrder.indexOf(b.key);
    const order = sort === 'hours' ? b.hours - a.hours
      : sort === 'measurement-area' ? compareMeasurementArea(a.project, b.project)
      : sort === 'owner' ? a.firstOwner.localeCompare(b.firstOwner)
      : sort === 'title' ? 0 : (a.nextDue || '9999').localeCompare(b.nextDue || '9999');
    return order || a.project.title.localeCompare(b.project.title) || String(a.key).localeCompare(String(b.key));
  });
}

const filled = (value) => value != null && String(value).trim() !== '';

function TaskFacts({ task }) {
  const owner = task.ownerName?.trim();
  const facts = [
    ['Assigned to', owner && owner !== 'Unassigned' ? displayName(owner) : task.ownerEmail || ''],
    ['Original due', filled(task.dueDate) ? date(task.dueDate) : '', !task.deferredDate && overdue(task)],
    ['Deferred to', filled(task.deferredDate) ? date(task.deferredDate) : '', task.deferredDate && overdue(task)],
    ['Est. hours', filled(task.estimatedHours) ? Number(task.estimatedHours).toLocaleString() : ''],
  ].filter(([, value]) => filled(value));
  return facts.length ? <dl className="attention-task-facts">{facts.map(([label, value, late]) => <div key={label}><dt>{label}</dt><dd className={late ? 'attention-overdue' : ''}>{value}</dd></div>)}</dl> : null;
}

function TaskDetails({ task, store, refreshKey }) {
  const [expanded, setExpanded] = useState(false);
  const hasDetails = [task.notes, task.deferredJustification, task.blockedReason, task.assignedDate].some(filled);
  if (!hasDetails && !store) return null;
  return <details className="attention-task-details" onToggle={(event) => setExpanded(event.currentTarget.open)}><summary>Details & justification</summary><div className="attention-details-content">
    {filled(task.assignedDate) && <p><strong>Assigned / created</strong>{date(task.assignedDate)}</p>}
    {filled(task.notes) && <p><strong>Notes</strong>{task.notes}</p>}
    {filled(task.deferredJustification) && <p><strong>Deferral justification</strong>{task.deferredJustification}</p>}
    {filled(task.blockedReason) && <p><strong>Blocked reason</strong>{task.blockedReason}</p>}
    {expanded && store && <TaskDocuments taskId={task.id} store={store} readOnly compact refreshKey={refreshKey} />}
  </div></details>;
}

export function Attention({ projects, organization, onOpen, onTask, controls, onControls, store, refreshKey }) {
  const { groupBy, search, sort } = controls;
  const byProject = groupBy === 'project';
  const rows = useMemo(() => attentionRows(projects, organization, search, sort), [projects, organization, search, sort]);
  const ordered = useMemo(() => attentionGroups(rows, groupBy, sort), [rows, groupBy, sort]);
  const hours = rows.reduce((sum, { task }) => sum + (Number(task.estimatedHours) || 0), 0);
  const projectSortHint = { due: 'Projects ordered by next task due date.', title: 'Projects ordered by name.', owner: 'Projects ordered by first assignee A–Z.', 'measurement-area': 'Projects ordered by measurement area.', hours: 'Projects ordered by total estimated hours.' }[sort];
  return <div className="page-stack attention-page">
    <section className="page-heading"><span className="section-kicker">Outstanding tasking</span><h1>{organization ? organizationAttentionLabel(organization) : 'Needs Attention'}</h1><p>Assignments, deadlines and workload across your projects.</p></section>
    <section className="attention-summary" aria-label="Outstanding task summary"><div><strong>{rows.length}</strong><span>Open tasks</span></div><div><strong>{new Set(rows.map(({ project }) => project.id)).size}</strong><span>Projects</span></div><div className={rows.some(({ task }) => overdue(task)) ? 'attention-overdue' : ''}><strong>{rows.filter(({ task }) => overdue(task)).length}</strong><span>Overdue</span></div><div><strong>{hours.toLocaleString(undefined, { maximumFractionDigits: 2 })}</strong><span>Estimated hours</span></div></section>
    <section className="attention-view-controls" aria-label="Attention view controls">
      <div className="attention-toolbar">
        <div className="attention-group-toggle" role="group" aria-label="View by">
          {[['status', 'Tasks', 'tasks'], ['project', 'Projects', 'folder']].map(([group, label, icon]) => <button key={group} type="button" aria-pressed={groupBy === group} aria-controls="attention-results" onClick={() => onControls({ ...controls, groupBy: group })}><Icon name={icon} size={16} />{label}</button>)}
        </div>
        <label className="search-box"><Icon name="search" size={16} /><input type="search" aria-label="Search attention tasks" placeholder="Search tasks, projects or people…" value={search} onChange={(event) => onControls({ ...controls, search: event.target.value })} /></label>
        <label className="select-wrap attention-sort"><span>Sort</span><select aria-label={byProject ? 'Sort attention projects' : 'Sort attention tasks'} value={sort} onChange={(event) => onControls({ ...controls, sort: event.target.value })}><option value="due">{byProject ? 'Next due soonest' : 'Due soonest'}</option><option value="title">{byProject ? 'Project name A–Z' : 'Task name A–Z'}</option><option value="owner">{byProject ? 'First assignee A–Z' : 'Assignee A–Z'}</option><option value="measurement-area">Measurement Area A–Z</option><option value="hours">{byProject ? 'Total hours high–low' : 'Hours high–low'}</option></select></label>
      </div>
      <p className="attention-view-hint">{byProject ? `${projectSortHint} Totals reflect the tasks shown.` : 'Tasks grouped by status, with your chosen sort applied within each group.'}</p>
    </section>
    <div id="attention-results" className={`attention-groups attention-groups-${groupBy}`} role="region" aria-label={byProject ? 'Tasks by project' : 'Tasks by status'} tabIndex={byProject ? undefined : 0}>{ordered.map(({ key, items, project, hours, overdueCount, nextDue }) => <details className="panel attention-section" key={`${groupBy}-${key}`} open><summary>
      {byProject ? <><span className="attention-project-icon"><Icon name="folder" size={19} /></span><span className="attention-project-heading"><strong>{project.title}</strong><span className="attention-project-stats"><span className="attention-count">{items.length} open {items.length === 1 ? 'task' : 'tasks'}</span>{overdueCount > 0 && <span className="attention-overdue">{overdueCount} overdue</span>}<span>{formatHours(hours)} est. hours</span>{nextDue && <span>Next due {date(nextDue)}</span>}</span></span></> : <><span className={`attention-dot status-${key.toLowerCase().replaceAll(' ', '-')}`} aria-hidden="true" /><strong>{key}</strong><span className="attention-count">{items.length} {items.length === 1 ? 'task' : 'tasks'}</span></>}
    </summary><div className="attention-task-list">{items.map(({ project, task }) => <article className="attention-task-card" key={task.id}>
      <div className="attention-task-heading"><div><button className="attention-project-link" onClick={() => onOpen(project)}>{project.title}</button><button className="attention-task-title" onClick={() => onTask(task)}>{task.title}</button></div><span className={`attention-status status-${task.status.toLowerCase().replaceAll(' ', '-')}`}>{task.status}</span></div>
      <div className="attention-task-meta">{filled(task.organization) && <span>{task.organization}</span>}{filled(project.measurementArea) && <span>{project.measurementArea}</span>}{workflowData.phases.some((phase) => phase.key === task.phaseKey) && <span>{workflowData.phases.find((phase) => phase.key === task.phaseKey).name}</span>}</div>
      <TaskFacts task={task} />
      <TaskDetails task={task} store={store} refreshKey={refreshKey} />
    </article>)}</div></details>)}</div>
    {!rows.length && <section className="panel attention-empty"><h2>{search ? 'No matching tasks' : 'Nothing needs attention'}</h2><p>{search ? 'Try a different task, project or assignee.' : 'There are no outstanding tasks for this organization.'}</p></section>}
  </div>;
}
