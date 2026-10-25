// @ts-check
(() => {
  "use strict";
  const core = globalThis.M365ExtenderCore;
  // Oldest per-conversation choices are dropped beyond this many.
  const MAX_CONVERSATION_TONES = 200;
  const state = { tone: core.DEFAULT_TONE, loaded: false, saving: false,
    connected: false, hookReady: false, error: "", observed: {},
    /** @type {{ key: string } | null} */ conversation: null,
    /** @type {string | null} */ conversationTone: null };
  /** @type {ConversationTones} */
  let conversationTones = {};
  const revisions = { selectedTone: 0, conversationTones: 0 };
  const panel = globalThis.M365ExtenderPanel.mount(/** @type {HTMLElement} */ (document.getElementById("panel")), {
    onSelect: tone => save({ selectedTone: tone }),
    onConversationSelect: tone => saveConversation(tone)
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
  /** @param {string | null} tone */
  function saveConversation(tone) {
    const key = state.conversation?.key;
    if (!key) return;
    const model = core.modelFor(tone);
    const next = { ...conversationTones };
    if (model) next[key] = { tone: model.tone, at: Date.now() };
    else delete next[key];
    const recent = Object.entries(next).sort(([, a], [, b]) => (b?.at || 0) - (a?.at || 0)).slice(0, MAX_CONVERSATION_TONES);
    return save({ conversationTones: Object.fromEntries(recent) });
  }
  function refreshConversation() {
    const key = state.conversation?.key;
    state.conversationTone = (key && core.modelFor(conversationTones[key]?.tone)?.tone) || null;
  }
  function applySettings(settings) {
    if ("selectedTone" in settings) state.tone = core.modelFor(settings.selectedTone)?.tone || core.DEFAULT_TONE;
    if ("conversationTones" in settings) conversationTones = settings.conversationTones && typeof settings.conversationTones === "object" ? settings.conversationTones : {};
    refreshConversation();
  }
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local") return;
    for (const key of Object.keys(revisions)) if (changes[key]) revisions[key]++;
    applySettings(Object.fromEntries(Object.entries(changes).map(([key, change]) => [key, change.newValue])));
    render();
  });
  const initialRevisions = { ...revisions };
  chrome.storage.local.get(Object.keys(revisions)).then(settings => {
    for (const key of Object.keys(revisions)) {
      if (revisions[key] === initialRevisions[key] && key in settings) applySettings({ [key]: settings[key] });
    }
    state.loaded = true; render();
  }).catch(() => { state.error = "Could not load saved preferences. Reopen this popup to retry."; render(); });

  async function readTab() {
    try {
      // Reading tab IDs and messaging our own content script needs no tabs permission.
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      const snapshot = await chrome.tabs.sendMessage(/** @type {number} */ (tab.id), { type: "M365_EXTENDER_SNAPSHOT" });
      state.connected = true; state.hookReady = snapshot.hookReady;
      state.observed = snapshot.observed || {};
      state.conversation = typeof snapshot.conversation?.key === "string" ? { key: snapshot.conversation.key } : null;
    } catch { state.connected = false; state.hookReady = false; state.observed = {}; state.conversation = null; }
    refreshConversation();
    render();
  }
  readTab();
  // Popup only: stay current when a reply finishes while it remains open.
  setInterval(readTab, 1500);
  render();
})();
