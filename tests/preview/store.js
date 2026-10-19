// Test-only Chrome mock. Never included in the extension ZIP.
window.fixtureStore = {
  data: JSON.parse(localStorage.getItem('extender-fixture-settings') || '{"selectedTone":"Gpt_61_Sol_Reasoning","inPageEnabled":true}'),
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
