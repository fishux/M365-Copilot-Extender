/* Isolated world: storage, UI, and a strictly filtered metadata bridge. */
(() => {
  "use strict";
  const core = globalThis.M365ExtenderCore;
  const CHANNEL = "m365-copilot-extender-v1";
  const state = { tone: core.DEFAULT_TONE, inPageEnabled: true, loaded: false, saving: false,
    hookReady: false, connected: true, error: "", observed: {} };
  const warnings = {
    "settings-unavailable": "Preference was not ready; this request used Copilot’s setting.",
    "binary-unsupported": "Copilot sent an unsupported frame; override was not applied.",
    "request-rejected": "Copilot rejected the request. Try a new conversation or another model.",
    "connection-closed": "Connection closed before the preference could be applied.",
    "send-failed": "Copilot could not send the pending message. Please retry."
  };
  let host, panel;
  let preferenceRevision = 0;
  let visibilityRevision = 0;
  let activeSocketId = null;
  const post = data => window.postMessage({ channel: CHANNEL, direction: "to-page", ...data }, location.origin);
  const configure = () => post({ kind: "configure", tone: state.loaded ? state.tone : null });
  function render() {
    if (!host) return;
    host.style.display = state.loaded && state.inPageEnabled ? "block" : "none";
    panel.update(state);
  }

  async function selectTone(tone) {
    if (!core.modelFor(tone)) return;
    state.saving = true; state.error = ""; render();
    try {
      await chrome.storage.local.set({ selectedTone: tone });
      state.tone = tone; configure();
    } catch { state.error = "Could not save your model. Your previous preference is still active."; }
    finally { state.saving = false; render(); }
  }

  function mount() {
    if (!document.body || host) return;
    host = document.createElement("m365-copilot-extender");
    Object.assign(host.style, { position: "fixed", top: "84px", right: "16px", zIndex: "2147483000",
      maxWidth: "calc(100vw - 24px)", maxHeight: "calc(100vh - 100px)", overflow: "auto",
      borderRadius: "18px", colorScheme: "light dark" });
    panel = globalThis.M365ExtenderPanel.mount(host, { onSelect: selectTone });
    document.body.append(host); render();
  }
  if (document.body) mount();
  else document.addEventListener("DOMContentLoaded", mount, { once: true });

  window.addEventListener("message", event => {
    const message = event.data;
    if (event.source !== window || event.origin !== location.origin || message?.channel !== CHANNEL || message.direction !== "from-page") return;
    const observed = state.observed;
    if (message.kind === "ready") {
      state.hookReady = true;
      if (state.loaded) configure();
    } else if (message.kind === "configured") {
      state.hookReady = true;
    } else if (message.kind === "connection") {
      // A new socket may belong to a new signed-in account. Do not reuse quotas.
      activeSocketId = message.socketId;
      state.observed = { scenario: safeScenario(message.scenario) };
    } else if (message.kind === "applied" && core.modelFor(message.tone)) {
      if (activeSocketId !== message.socketId) state.observed = {};
      activeSocketId = message.socketId;
      state.observed.appliedTone = message.tone;
      state.observed.scenario = safeScenario(message.scenario);
      delete state.observed.warning; delete state.observed.result; delete state.observed.origin;
    } else if (message.kind === "warning" && warnings[message.code]) {
      if (message.socketId && message.socketId !== activeSocketId) return;
      observed.warning = warnings[message.code];
    } else if (message.kind === "metadata") {
      if (message.socketId !== activeSocketId) return;
      observed.scenario = safeScenario(message.scenario);
      if (message.metering && typeof message.metering === "object") {
        // Revalidate data from the page, which is not an authenticated boundary.
        observed.metering = core.parseMetering(Object.fromEntries(Object.entries(message.metering).map(([key, value]) => [key, { remainingAllowance: value }])));
        observed.quotaAt = Date.now();
      }
      if (core.allowedResults.includes(message.result)) observed.result = message.result;
      if (core.allowedOrigins.includes(message.origin)) observed.origin = message.origin;
      if (core.count(message.conversation?.used) !== null && core.count(message.conversation?.limit) !== null) observed.conversation = { used: message.conversation.used, limit: message.conversation.limit };
      else if (message.final) delete observed.conversation;
      if ("expiresAt" in message) observed.expiresAt = core.parseExpiry(message.expiresAt);
      else if (message.final) delete observed.expiresAt;
    } else { return; }
    render();
  });

  function safeScenario(value) {
    return ["OfficeWebPaidCopilot", "OfficeWebIncludedCopilot"].includes(value) ? value : "Unknown";
  }
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local") return;
    if (changes.selectedTone) {
      preferenceRevision++;
      state.tone = core.modelFor(changes.selectedTone.newValue)?.tone || core.DEFAULT_TONE;
      if (state.loaded) configure();
    }
    if (changes.inPageEnabled) { visibilityRevision++; state.inPageEnabled = changes.inPageEnabled.newValue !== false; }
    render();
  });
  chrome.runtime.onMessage.addListener((message, sender, respond) => {
    if (sender.id !== chrome.runtime.id || message?.type !== "M365_EXTENDER_SNAPSHOT") return;
    respond({ hookReady: state.hookReady, observed: state.observed });
  });
  const initialRevision = preferenceRevision;
  const initialVisibilityRevision = visibilityRevision;
  chrome.storage.local.get(["selectedTone", "inPageEnabled"]).then(settings => {
    if (preferenceRevision === initialRevision) state.tone = core.modelFor(settings.selectedTone)?.tone || core.DEFAULT_TONE;
    if (visibilityRevision === initialVisibilityRevision) state.inPageEnabled = settings.inPageEnabled !== false;
    state.loaded = true; configure(); render();
  }).catch(() => {
    state.error = "Saved preferences could not be loaded. Reload this tab to retry.";
    state.observed.warning = "Preferences unavailable; this tab is using Copilot’s setting.";
    post({ kind: "configure", tone: null }); render();
  });
  post({ kind: "hello" });
  setTimeout(() => {
    if (!state.hookReady) { state.observed.warning = "Extension hook unavailable. Reload Copilot after installing or updating."; render(); }
  }, 2500);
  // Refresh only the visible timestamp label, without network requests.
  setInterval(() => { if (state.inPageEnabled) render(); }, 60_000);
})();
