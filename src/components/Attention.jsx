import React, { useMemo } from 'react';
import { displayName } from '../lib/displayName';
import { workflowData } from '../data/workflow';

const statusOrder = ['Blocked', 'In Progress', 'Not Started'];
const closed = new Set(['Complete', 'Not Required', 'Not Applicable']);
const date = (value) => value ? new Date(`${value}T12:00:00`).toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' }) : 'Not set';
const overdue = (task) => !!(task.deferredDate || task.dueDate) && (task.deferredDate || task.dueDate) < new Date().toLocaleDateString('en-CA');

export function attentionRows(projects, organization = '', search = '', sort = 'due') {
  const query = search.toLowerCase().trim();
  return projects.flatMap((project) => project.tasks.filter((task) => !closed.has(task.status) && (!organization || task.organization === organization)).map((task) => ({ project, task })))
    .filter(({ project, task }) => !query || [project.title, task.title, task.ownerName, displayName(task.ownerName), task.organization].join(' ').toLowerCase().includes(query))
    .sort((a, b) => {
      const value = ({ task }) => sort === 'hours' ? -(Number(task.estimatedHours) || 0) : sort === 'owner' ? displayName(task.ownerName || '') : (task.deferredDate || task.dueDate || '9999');
      const x = value(a), y = value(b);
      return (typeof x === 'number' ? x - y : x.localeCompare(y)) || a.project.title.localeCompare(b.project.title) || a.task.title.localeCompare(b.task.title);
    });
}

export function Attention({ projects, organization, onOpen, onTask, controls, onControls }) {
  const { groupBy, search, sort } = controls;
  const rows = useMemo(() => attentionRows(projects, organization, search, sort), [projects, organization, search, sort]);
  const groups = new Map();
  for (const row of rows) {
    const key = groupBy === 'project' ? row.project.projectKey : row.task.status;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  const ordered = [...groups].sort(([a, x], [b, y]) => groupBy === 'project' ? x[0].project.title.localeCompare(y[0].project.title) : statusOrder.indexOf(a) - statusOrder.indexOf(b));
  const hours = rows.reduce((sum, { task }) => sum + (Number(task.estimatedHours) || 0), 0);
  return <div className="page-stack attention-page">
    <section className="page-heading"><span className="section-kicker">Outstanding tasking</span><h1>{organization ? `${organization} Needs Attention` : 'Needs Attention'}</h1><p>Assignments, deadlines and workload across your projects.</p></section>
    <section className="attention-summary" aria-label="Outstanding task summary"><div><strong>{rows.length}</strong><span>Open tasks</span></div><div><strong>{new Set(rows.map(({ project }) => project.id)).size}</strong><span>Projects</span></div><div className={rows.some(({ task }) => overdue(task)) ? 'attention-overdue' : ''}><strong>{rows.filter(({ task }) => overdue(task)).length}</strong><span>Overdue</span></div><div><strong>{hours.toLocaleString(undefined, { maximumFractionDigits: 2 })}</strong><span>Estimated hours</span></div></section>
    <section className="attention-toolbar"><label className="search-box"><input type="search" aria-label="Search attention tasks" placeholder="Search tasks, projects or people…" value={search} onChange={(event) => onControls({ ...controls, search: event.target.value })} /></label><div className="attention-group-toggle" aria-label="Group tasks">{['status', 'project'].map((group) => <button key={group} type="button" aria-pressed={groupBy === group} onClick={() => onControls({ ...controls, groupBy: group })}>By {group}</button>)}</div><label className="select-wrap"><select aria-label="Sort attention tasks" value={sort} onChange={(event) => onControls({ ...controls, sort: event.target.value })}><option value="due">Due soonest</option><option value="owner">Assignee A–Z</option><option value="hours">Hours high–low</option></select></label></section>
    <div className="attention-groups">{ordered.map(([key, items]) => <details className="panel attention-section" key={`${groupBy}-${key}`} open><summary><span className={`attention-dot status-${items[0].task.status.toLowerCase().replaceAll(' ', '-')}`} aria-hidden="true" /><strong>{groupBy === 'project' ? items[0].project.title : key}</strong><span className="attention-count">{items.length} {items.length === 1 ? 'task' : 'tasks'}</span></summary><div className="attention-task-list">{items.map(({ project, task }) => <article className="attention-task-card" key={task.id}>
      <div className="attention-task-heading"><div><button className="attention-project-link" onClick={() => onOpen(project)}>{project.title}</button><button className="attention-task-title" onClick={() => onTask(task)}>{task.title}</button></div><span className={`attention-status status-${task.status.toLowerCase().replaceAll(' ', '-')}`}>{task.status}</span></div>
      <div className="attention-task-meta"><span>{task.organization}</span><span>{workflowData.phases.find((phase) => phase.key === task.phaseKey)?.name}</span></div>
      <dl className="attention-task-facts"><div><dt>Assigned to</dt><dd>{displayName(task.ownerName || 'Unassigned')}</dd></div><div><dt>Original due</dt><dd className={!task.deferredDate && overdue(task) ? 'attention-overdue' : ''}>{date(task.dueDate)}</dd></div><div><dt>Deferred to</dt><dd className={task.deferredDate && overdue(task) ? 'attention-overdue' : ''}>{task.deferredDate ? date(task.deferredDate) : '—'}</dd></div><div><dt>Est. hours</dt><dd>{task.estimatedHours == null || task.estimatedHours === '' ? '—' : Number(task.estimatedHours).toLocaleString()}</dd></div></dl>
      {(task.notes || task.deferredJustification || task.blockedReason || task.assignedDate) && <details className="attention-task-details"><summary>Details & justification</summary><div>{task.assignedDate && <p><strong>Assigned / created</strong>{date(task.assignedDate)}</p>}{task.notes && <p><strong>Notes</strong>{task.notes}</p>}{task.deferredJustification && <p><strong>Deferral justification</strong>{task.deferredJustification}</p>}{task.blockedReason && <p><strong>Blocked reason</strong>{task.blockedReason}</p>}</div></details>}
    </article>)}</div></details>)}</div>
    {!rows.length && <section className="panel attention-empty"><h2>{search ? 'No matching tasks' : 'Nothing needs attention'}</h2><p>{search ? 'Try a different task, project or assignee.' : 'There are no outstanding tasks for this organization.'}</p></section>}
  </div>;
}
