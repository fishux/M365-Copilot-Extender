// @ts-check
(() => {
  "use strict";
  const core = globalThis.M365ExtenderCore;
  // The default model is fixed in models.json; only per-chat choices are saved.
  const state = { tone: core.DEFAULT_TONE, loaded: false, saving: false,
    connected: false, hookReady: false, error: "", observed: {},
    /** @type {{ key: string } | null} */ conversation: null,
    /** An unsent new chat; its model is chosen before it has a conversation. @type {{ tone: string | null } | null} */ newChat: null,
    /** @type {string | null} */ conversationTone: null };
  /** @type {ConversationTones} */
  let conversationTones = {};
  /** Last allowances reported in any tab. @type {Usage | null} */
  let usage = null;
  /** @type {any} */
  let tabObserved = {};
  const revisions = { conversationTones: 0, usage: 0 };
  const panel = globalThis.M365ExtenderPanel.mount(/** @type {HTMLElement} */ (document.getElementById("panel")), {
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
  /**
   * Choosing a model in a chat keeps that chat on it from then on. Choosing the
   * default model removes the entry, so the chat follows the default again.
   * @param {string | null} tone
   */
  function saveConversation(tone) {
    const key = state.conversation?.key;
    if (!key) return state.newChat ? chooseForNewChat(tone) : undefined;
    const model = core.modelFor(tone);
    const next = { ...conversationTones };
    if (model && model.tone !== core.DEFAULT_TONE) next[key] = { tone: model.tone, at: Date.now() };
    else delete next[key];
    const recent = Object.entries(next).sort(([, a], [, b]) => (b?.at || 0) - (a?.at || 0)).slice(0, core.MAX_SAVED_CONVERSATIONS);
    return save({ conversationTones: Object.fromEntries(recent) });
  }
  /** @param {string | null} tone */
  async function chooseForNewChat(tone) {
    state.saving = true; state.error = ""; render();
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      const reply = await chrome.tabs.sendMessage(/** @type {number} */ (tab.id), { type: "M365_EXTENDER_NEW_CHAT_TONE", tone });
      if (!reply?.ok) throw new Error("rejected");
      await readTab();
    } catch { state.error = "Could not change the model. Reload Copilot and try again."; }
    finally { state.saving = false; render(); }
  }
  function refreshConversation() {
    const key = state.conversation?.key;
    state.conversationTone = (key && core.modelFor(conversationTones[key]?.tone)?.tone)
      || (!key && core.modelFor(state.newChat?.tone)?.tone) || null;
  }
  /** Tab readings, with allowances from storage when they are newer (or the tab has none). */
  function refreshObserved() {
    const observed = { ...tabObserved };
    if (usage && usage.quotaAt > (Number(observed.quotaAt) || 0)) {
      observed.metering = usage.metering; observed.quotaAt = usage.quotaAt;
    }
    state.observed = observed;
  }
  function applySettings(settings) {
    if ("conversationTones" in settings) conversationTones = settings.conversationTones && typeof settings.conversationTones === "object" ? settings.conversationTones : {};
    if ("usage" in settings) {
      const value = settings.usage;
      usage = value && typeof value === "object" && Number(value.quotaAt) > 0 && value.metering && typeof value.metering === "object"
        ? { metering: value.metering, quotaAt: Number(value.quotaAt) } : null;
    }
    refreshConversation();
    refreshObserved();
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
  }).catch(() => { state.error = "Could not load saved choices. Reopen this popup to retry."; render(); });

  async function readTab() {
    try {
      // Reading tab IDs and messaging our own content script needs no tabs permission.
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      const snapshot = await chrome.tabs.sendMessage(/** @type {number} */ (tab.id), { type: "M365_EXTENDER_SNAPSHOT" });
      state.connected = true; state.hookReady = snapshot.hookReady;
      tabObserved = snapshot.observed || {};
      state.conversation = typeof snapshot.conversation?.key === "string" ? { key: snapshot.conversation.key } : null;
      state.newChat = snapshot.newChat && typeof snapshot.newChat === "object" ? { tone: typeof snapshot.newChat.tone === "string" ? snapshot.newChat.tone : null } : null;
    } catch { state.connected = false; state.hookReady = false; tabObserved = {}; state.conversation = null; state.newChat = null; }
    refreshConversation();
    refreshObserved();
    render();
  }
  readTab();
  // Popup only: stay current when a reply finishes while it remains open.
  setInterval(readTab, 1500);
  render();
})();
