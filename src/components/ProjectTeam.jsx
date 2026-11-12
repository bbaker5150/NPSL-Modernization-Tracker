import React from 'react';
import { displayName } from '../lib/displayName';
import { projectEngineers } from '../lib/projectEngineers';

export function engineerInitials(engineer) {
  const name = displayName(engineer.name || engineer.email || engineer.loginName || 'Unassigned');
  return name.split(/[\s,]+/).filter(Boolean).slice(0, 2).map(part => part[0]).join('').toUpperCase() || '?';
}

export function projectTeamText(project) {
  const engineers = projectEngineers(project);
  return engineers.length ? engineers.map(engineer => displayName(engineer.name || engineer.email || engineer.loginName)).join(', ') : 'Unassigned';
}

export function projectTeamSearch(project) {
  return projectEngineers(project).map(engineer => `${engineer.name} ${displayName(engineer.name)} ${engineer.email} ${engineer.loginName}`).join(' ');
}

export function ProjectTeam({ project, detailed = false }) {
  const engineers = projectEngineers(project);
  if (!engineers.length) return <span className="project-team-empty">Unassigned</span>;
  return <ul className={`project-team-display ${detailed ? 'detailed' : 'compact'}`} aria-label="Project engineers">
    {engineers.map((engineer, index) => <li key={engineer.loginName || engineer.email || `${engineer.name}-${index}`}>
      <span className={`avatar ${detailed ? 'small' : 'tiny'}`} aria-hidden="true">{engineerInitials(engineer)}</span>
      <span className="project-engineer-identity"><strong>{displayName(engineer.name || engineer.email || engineer.loginName)}</strong>{detailed && <small>{engineer.email || engineer.loginName || 'No SharePoint identity assigned'}</small>}</span>
    </li>)}
  </ul>;
}
