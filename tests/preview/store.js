// Test-only Chrome mock. Never included in the extension ZIP.
window.fixtureStore = {
  data: {},
  listeners: [],
  async get() { return { ...this.data }; },
  async set(patch) {
    const changes = {};
    for (const [key, newValue] of Object.entries(patch)) {
      changes[key] = { oldValue: this.data[key], newValue }; this.data[key] = newValue;
    }
    localStorage.setItem('extender-fixture-settings', JSON.stringify(this.data));
    for (const { owner, listener } of this.listeners) if (!owner.closed) listener(changes, 'local');
  }
};
