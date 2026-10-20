/* Isolated world: storage and a strictly filtered metadata bridge. */
(() => {
  "use strict";
  const core = globalThis.M365ExtenderCore;
  const CHANNEL = "m365-copilot-extender-v1";
  const state = { tone: core.DEFAULT_TONE, loaded: false, hookReady: false, observed: {} };
  const warnings = {
    "settings-unavailable": "Preference was not ready; this request used Copilot’s setting.",
    "binary-unsupported": "Copilot sent an unsupported frame; override was not applied.",
    "request-rejected": "Copilot rejected the request. Try a new conversation or another model.",
    "connection-closed": "Connection closed before the preference could be applied.",
    "send-failed": "Copilot could not send the pending message. Please retry."
  };
  let preferenceRevision = 0;
  let activeSocketId = null;
  const post = data => window.postMessage({ channel: CHANNEL, direction: "to-page", ...data }, location.origin);
  const configure = () => post({ kind: "configure", tone: state.loaded ? state.tone : null });

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
      // Copilot may pre-open another socket after a reply. Keep the last
      // readings until a chat is actually sent on a different connection.
      if (activeSocketId === null) {
        activeSocketId = message.socketId;
        state.observed.scenario = safeScenario(message.scenario);
      }
    } else if (message.kind === "request" || (message.kind === "applied" && core.modelFor(message.tone))) {
      if (activeSocketId !== message.socketId) state.observed = {};
      activeSocketId = message.socketId;
      if (message.kind === "applied") state.observed.appliedTone = message.tone;
      state.observed.scenario = safeScenario(message.scenario);
      if (state.loaded) delete state.observed.warning;
    } else if (message.kind === "warning" && warnings[message.code]) {
      if (message.socketId && message.socketId !== activeSocketId) return;
      observed.warning = warnings[message.code];
    } else if (message.kind === "metadata") {
      if (message.socketId !== activeSocketId) return;
      observed.scenario = safeScenario(message.scenario);
      if (message.metering && typeof message.metering === "object") {
        // Revalidate data from the page, which is not an authenticated boundary.
        const metering = core.parseMetering(Object.fromEntries(Object.entries(message.metering).map(([key, value]) => [key, { remainingAllowance: value }])));
        if (Object.keys(metering).length) {
          observed.metering = { ...observed.metering, ...metering };
          observed.quotaAt = Date.now();
        }
      }
      if (core.allowedResults.includes(message.result)) observed.result = message.result;
      if (core.allowedOrigins.includes(message.origin)) observed.origin = message.origin;
      if (core.count(message.conversation?.used) !== null && core.count(message.conversation?.limit) !== null) observed.conversation = { used: message.conversation.used, limit: message.conversation.limit };
      const expiresAt = core.parseExpiry(message.expiresAt);
      if (expiresAt) observed.expiresAt = expiresAt;
    } else { return; }
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
  });
  chrome.runtime.onMessage.addListener((message, sender, respond) => {
    if (sender.id !== chrome.runtime.id || message?.type !== "M365_EXTENDER_SNAPSHOT") return;
    respond({ hookReady: state.hookReady, observed: state.observed });
  });
  const initialRevision = preferenceRevision;
  chrome.storage.local.get(["selectedTone"]).then(settings => {
    if (preferenceRevision === initialRevision) state.tone = core.modelFor(settings.selectedTone)?.tone || core.DEFAULT_TONE;
    state.loaded = true; configure();
  }).catch(() => {
    state.observed.warning = "Preferences unavailable; this tab is using Copilot’s setting.";
    post({ kind: "configure", tone: null });
  });
  post({ kind: "hello" });
  setTimeout(() => {
    if (!state.hookReady) { state.observed.warning = "Extension hook unavailable. Reload Copilot after installing or updating."; }
  }, 2500);
})();
