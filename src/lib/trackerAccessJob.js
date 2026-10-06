// Keep the running operation outside the mounted Users screen. SharePoint is
// the source of truth after a host reload; no cached permission decisions.
export class TrackerAccessJob {
  constructor(store) {
    this.store = store;
    this.listeners = new Set();
    this.state = { busy: false, status: '', error: '' };
  }
  getSnapshot = () => this.state;
  subscribe = listener => { this.listeners.add(listener); return () => this.listeners.delete(listener); };
  publish(change) { this.state = { ...this.state, ...change }; this.listeners.forEach(listener => listener()); }
  run = () => {
    if (this.running) return this.running;
    this.publish({ busy: true, error: '', status: 'Checking tracker access…' });
    this.running = Promise.resolve().then(() => this.store.prepareTrackerAccess(
      status => this.publish({ status }),
    )).then(result => {
      this.publish({ status: `Verified tracker groups for ${result.users} users and engineer permissions for ${result.projects} projects.` });
      return result;
    }).catch(error => {
      this.publish({ error: `Setup stopped. Retry to check existing permissions and update only unfinished or changed items. ${error.message}` });
    }).finally(() => { this.running = null; this.publish({ busy: false }); });
    return this.running;
  };
}
