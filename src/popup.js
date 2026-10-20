(() => {
  "use strict";
  const core = globalThis.M365ExtenderCore;
  const state = { tone: core.DEFAULT_TONE, loaded: false, saving: false,
    connected: false, hookReady: false, error: "", observed: {} };
  const revisions = { selectedTone: 0 };
  const panel = globalThis.M365ExtenderPanel.mount(document.getElementById("panel"), {
    onSelect: tone => save({ selectedTone: tone })
  });
  const render = () => panel.update(state);
  async function save(patch) {
    state.saving = true; state.error = ""; render();
    try {
      await chrome.storage.local.set(patch);
      applySettings(patch);
    } catch { state.error = "Could not save changes. Please try again."; }
    finally { state.saving = false; render(); }
  }
  function applySettings(settings) {
    if ("selectedTone" in settings) state.tone = core.modelFor(settings.selectedTone)?.tone || core.DEFAULT_TONE;
  }
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local") return;
    for (const key of Object.keys(revisions)) if (changes[key]) revisions[key]++;
    applySettings(Object.fromEntries(Object.entries(changes).map(([key, change]) => [key, change.newValue])));
    render();
  });
  const initialRevisions = { ...revisions };
  chrome.storage.local.get(["selectedTone"]).then(settings => {
    for (const key of Object.keys(revisions)) {
      if (revisions[key] === initialRevisions[key] && key in settings) applySettings({ [key]: settings[key] });
    }
    state.loaded = true; render();
  }).catch(() => { state.error = "Could not load saved preferences. Reopen this popup to retry."; render(); });

  async function readTab() {
    try {
      // Reading tab IDs and messaging our own content script needs no tabs permission.
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      const snapshot = await chrome.tabs.sendMessage(tab.id, { type: "M365_EXTENDER_SNAPSHOT" });
      state.connected = true; state.hookReady = snapshot.hookReady;
      state.observed = snapshot.observed || {};
    } catch { state.connected = false; state.hookReady = false; state.observed = {}; }
    render();
  }
  readTab();
  // Popup only: stay current when a reply finishes while it remains open.
  setInterval(readTab, 1500);
  render();
})();
