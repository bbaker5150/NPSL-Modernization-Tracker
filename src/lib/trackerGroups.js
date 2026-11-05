// Basic access is shared by all Tracker roles. Keep this separate from elevated
// role groups so demotion to Viewer does not remove access to the application.
export const TRACKER_USERS_GROUP = 'Tracker Users';
export const TRACKER_GROUPS = {
  'Project Engineer': 'Tracker Project Engineers',
  Manager: 'Tracker Managers',
};
