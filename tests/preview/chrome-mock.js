const store = window.parent.fixtureStore;
window.chrome = {
  storage: { local: { get: () => store.get(), set: patch => store.set(patch) }, onChanged: { addListener(listener) { store.listeners.push({ owner: window, listener }); } } },
  runtime: { id: 'fixture-extension', onMessage: { addListener(listener) { window.fixtureRespond = listener; } } },
  tabs: { async query() { return [{ id: 1 }]; }, async sendMessage(id, message) {
    return new Promise((resolve, reject) => {
      const tab = window.parent.document.getElementById('copilot').contentWindow;
      if (!tab.fixtureRespond) return reject(new Error('No tab'));
      tab.fixtureRespond(message, { id: 'fixture-extension' }, resolve);
    });
  } }
};
