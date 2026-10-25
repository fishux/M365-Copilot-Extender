// @ts-check
/* Isolated world: storage and a strictly filtered metadata bridge. */
(() => {
  "use strict";
  const core = globalThis.M365ExtenderCore;
  const CHANNEL = "m365-copilot-extender-v1";
  const MAX_SLOTS = 50;
  /** @type {{ tone: string, loaded: boolean, failed: boolean, hookReady: boolean, conversationTones: ConversationTones, tabWarning?: string }} */
  const state = { tone: core.DEFAULT_TONE, loaded: false, failed: false, hookReady: false, conversationTones: {} };
  const warnings = {
    "settings-unavailable": "Preference was not ready; this request used Copilot’s setting.",
    "conversation-unavailable": "This conversation’s model was not ready; the default model was used.",
    "binary-unsupported": "Copilot sent an unsupported frame; override was not applied.",
    "request-rejected": "Copilot rejected the request. Try a new conversation or another model.",
    "connection-closed": "Connection closed before the preference could be applied.",
    "send-failed": "Copilot could not send the pending message. Please retry."
  };
  // Warnings about a chat the user just sent make its conversation current.
  const pendingSendWarnings = new Set(["settings-unavailable", "conversation-unavailable", "connection-closed", "send-failed"]);

  // Account-wide readings: allowances belong to the account, not a conversation.
  /** @type {{ metering?: Record<string, number>, quotaAt?: number, scenario?: string }} */
  const account = {};
  // Per-conversation readings, keyed by ConversationId, or "socket:<n>" when a
  // connection carries no conversation ID. Tab-local and never persisted.
  /** @type {Map<string, Observed>} */
  const slots = new Map();
  const sentSlots = new Set();
  const seenInUrl = new Set();
  /** @type {Map<string, string>} conversationId -> hashed storage key */
  const keys = new Map();
  const hashing = new Set();
  /** @type {string | null} */
  let activeSlot = null;
  let preferenceRevision = 0;
  let conversationRevision = 0;
  let lastPublished = "";

  const post = data => window.postMessage({ channel: CHANNEL, direction: "to-page", ...data }, location.origin);
  const configure = () => post({ kind: "configure", tone: state.loaded ? state.tone : null });
  const slotName = message => core.safeConversationId(message.conversationId) || `socket:${Number(message.socketId) || 0}`;
  const isConversation = name => typeof name === "string" && !name.startsWith("socket:");

  function slot(name) {
    let observed = slots.get(name);
    if (!observed) {
      observed = {};
      slots.set(name, observed);
      for (const old of slots.keys()) {
        if (slots.size <= MAX_SLOTS) break;
        if (old !== activeSlot && old !== name) { slots.delete(old); sentSlots.delete(old); }
      }
    }
    return observed;
  }

  // Conversation IDs are hashed before they touch storage or the popup, so
  // saved per-conversation choices cannot be matched to server records.
  async function keyFor(id) {
    const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(`m365-copilot-extender:${id}`));
    return [...new Uint8Array(digest)].slice(0, 16).map(byte => byte.toString(16).padStart(2, "0")).join("");
  }

  function ensureKey(id) {
    if (!id || keys.has(id) || hashing.has(id)) return;
    hashing.add(id);
    keyFor(id).then(key => {
      keys.set(id, key);
      answerHook(id);
      publishTab();
    }).catch(() => { /* Without a key this conversation uses the default model. */ })
      .finally(() => hashing.delete(id));
  }

  /** Tell the hook which tone this conversation uses (null = default). */
  function answerHook(id) {
    if (!state.loaded && !state.failed) return;
    const key = keys.get(id);
    if (!key && !state.failed) return;
    post({ kind: "conversation", conversationId: id, tone: (key && state.conversationTones[key]?.tone) || null });
  }

  function noteLocation() {
    const id = core.conversationFromPath(location.pathname);
    if (id) { seenInUrl.add(id); ensureKey(id); }
    return id;
  }

  /** The conversation the page is showing, or null for an unsent new chat. */
  function currentSlot() {
    const urlId = noteLocation();
    if (urlId) {
      if (slots.has(urlId)) return urlId;
      // The connection carried no conversation ID: fall back to the last send.
      return activeSlot && !isConversation(activeSlot) ? activeSlot : urlId;
    }
    // Navigation can lag the first reply of a new chat by a turn.
    return activeSlot && sentSlots.has(activeSlot) && !seenInUrl.has(activeSlot) ? activeSlot : null;
  }

  function effectiveTone(name) {
    const key = isConversation(name) ? keys.get(name) : undefined;
    return (key && state.conversationTones[key]?.tone) || state.tone;
  }

  function snapshot() {
    const name = currentSlot();
    const observed = { ...account, ...(name ? slots.get(name) : {}) };
    observed.warning = observed.warning || state.tabWarning;
    const key = isConversation(name) ? keys.get(/** @type {string} */ (name)) : undefined;
    return { hookReady: state.hookReady, observed, conversation: key ? { key } : null };
  }

  // Keep the toolbar badge and icon in step with this tab.
  function publishTab(force = false) {
    const tone = effectiveTone(currentSlot());
    const message = { type: "M365_EXTENDER_TAB", tone, remaining: core.lowestAllowance(tone, account.metering) };
    const serialized = JSON.stringify(message);
    if (serialized === lastPublished && !force) return;
    lastPublished = serialized;
    try { chrome.runtime.sendMessage(message)?.catch?.(() => {}); } catch { /* Extension reloaded. */ }
  }

  window.addEventListener("message", event => {
    const message = event.data;
    if (event.source !== window || event.origin !== location.origin || message?.channel !== CHANNEL || message.direction !== "from-page") return;
    if (message.kind === "ready") {
      state.hookReady = true;
      if (state.loaded) configure();
      for (const id of keys.keys()) answerHook(id);
    } else if (message.kind === "configured") {
      state.hookReady = true;
    } else if (message.kind === "connection") {
      // Copilot may pre-open another socket after a reply. Readings stay with
      // their conversation until a chat is actually sent on it.
      account.scenario = core.safeScenario(message.scenario);
      const id = core.safeConversationId(message.conversationId);
      if (id) { if (keys.has(id)) answerHook(id); else ensureKey(id); }
    } else if (message.kind === "request" || (message.kind === "applied" && core.modelFor(message.tone))) {
      const name = slotName(message);
      activeSlot = name;
      sentSlots.add(name);
      const observed = slot(name);
      if (message.kind === "applied") observed.appliedTone = core.modelFor(message.tone)?.tone;
      account.scenario = core.safeScenario(message.scenario);
      if (state.loaded) delete observed.warning;
    } else if (message.kind === "warning" && warnings[message.code]) {
      const name = slotName(message);
      if (pendingSendWarnings.has(message.code)) { activeSlot = name; sentSlots.add(name); }
      slot(name).warning = warnings[message.code];
    } else if (message.kind === "metadata") {
      const name = slotName(message);
      // Replies only count once a chat was sent in this conversation (or on this
      // connection), so an idle pre-opened socket cannot inject readings.
      if (!sentSlots.has(name)) return;
      const observed = slot(name);
      // Allowances are account-wide; only the latest send's replies update them,
      // so a late reply on an older connection cannot overwrite newer readings.
      if (name === activeSlot) {
        account.scenario = core.safeScenario(message.scenario);
        if (message.metering && typeof message.metering === "object") {
          // Revalidate data from the page, which is not an authenticated boundary.
          const metering = core.parseMetering(Object.fromEntries(Object.entries(message.metering).map(([key, value]) => [key, { remainingAllowance: value }])));
          if (Object.keys(metering).length) {
            account.metering = { ...account.metering, ...metering };
            account.quotaAt = Date.now();
          }
        }
      }
      if (core.allowedResults.includes(message.result)) observed.result = message.result;
      if (core.allowedOrigins.includes(message.origin)) observed.origin = message.origin;
      if (core.count(message.conversation?.used) !== null && core.count(message.conversation?.limit) !== null) observed.conversation = { used: message.conversation.used, limit: message.conversation.limit };
      const expiresAt = core.parseExpiry(message.expiresAt);
      if (expiresAt) observed.expiresAt = expiresAt;
    } else return;
    publishTab();
  });

  /** @returns {ConversationTones} */
  function validTones(raw) {
    const valid = {};
    if (!raw || typeof raw !== "object") return valid;
    for (const [key, entry] of Object.entries(raw)) {
      const model = core.modelFor(entry?.tone);
      if (/^[0-9a-f]{32}$/.test(key) && model) valid[key] = { tone: model.tone, at: Number(entry.at) || 0 };
    }
    return valid;
  }

  function applyConversationTones(raw) {
    state.conversationTones = validTones(raw);
    for (const id of keys.keys()) answerHook(id);
  }

  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local") return;
    if (changes.selectedTone) {
      preferenceRevision++;
      state.tone = core.modelFor(changes.selectedTone.newValue)?.tone || core.DEFAULT_TONE;
      if (state.loaded) configure();
    }
    if (changes.conversationTones) { conversationRevision++; applyConversationTones(changes.conversationTones.newValue); }
    publishTab();
  });
  chrome.runtime.onMessage.addListener((message, sender, respond) => {
    if (sender.id !== chrome.runtime.id || message?.type !== "M365_EXTENDER_SNAPSHOT") return;
    respond(snapshot());
  });
  // Single-page navigation (sidebar conversations) changes the current conversation.
  // The worker resets a tab's badge when it starts loading, so re-send after
  // navigation settles rather than relying on the de-duplicated state.
  const republish = () => setTimeout(() => publishTab(true), 300);
  try { /** @type {any} */ (window).navigation?.addEventListener("currententrychange", republish); } catch { /* Older Chrome. */ }
  window.addEventListener("pageshow", republish);

  const initialRevision = preferenceRevision, initialConversationRevision = conversationRevision;
  chrome.storage.local.get(["selectedTone", "conversationTones"]).then(settings => {
    if (preferenceRevision === initialRevision) state.tone = core.modelFor(settings.selectedTone)?.tone || core.DEFAULT_TONE;
    state.loaded = true;
    if (conversationRevision === initialConversationRevision) state.conversationTones = validTones(settings.conversationTones);
    configure();
    for (const id of keys.keys()) answerHook(id);
    publishTab();
  }).catch(() => {
    state.failed = true;
    state.tabWarning = "Preferences unavailable; this tab is using Copilot’s setting.";
    configure();
    for (const id of keys.keys()) answerHook(id);
  });
  post({ kind: "hello" });
  setTimeout(() => {
    if (!state.hookReady) state.tabWarning = "Extension hook unavailable. Reload Copilot after installing or updating.";
  }, 2500);
})();
