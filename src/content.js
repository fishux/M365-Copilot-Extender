// @ts-check
/* Isolated world: storage and a strictly filtered metadata bridge. */
(() => {
  "use strict";
  const core = globalThis.M365ExtenderCore;
  const CHANNEL = "m365-copilot-extender-v1";
  const MAX_SLOTS = 50;
  // Streaming replies update readings many times; save them at most this often.
  const PERSIST_DELAY = 250;
  /** @type {{ loaded: boolean, failed: boolean, hookReady: boolean, conversationTones: ConversationTones, readings: ConversationReadings, tabWarning?: string }} */
  const state = { loaded: false, failed: false, hookReady: false, conversationTones: {}, readings: {} };
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
  // Allowances are mirrored to storage so every tab shows the last reported value.
  /** @type {{ metering?: Record<string, number>, quotaAt?: number, scenario?: string }} */
  const account = {};
  // Per-conversation readings, keyed by ConversationId, or "socket:<n>" when a
  // connection carries no conversation ID. Conversations are also saved under
  // their hashed key so another tab, or a restart, shows the last reply.
  /** @type {Map<string, Observed>} */
  const slots = new Map();
  const sentSlots = new Set();
  const seenInUrl = new Set();
  /** @type {Map<string, string>} conversationId -> hashed storage key */
  const keys = new Map();
  const hashing = new Set();
  /** Conversations whose readings changed and still need saving. */
  const dirty = new Set();
  /** @type {ReturnType<typeof setTimeout> | null} */
  let persistTimer = null;
  /** @type {string | null} */
  let activeSlot = null;
  // A model chosen for a new chat before its first message. A new chat has no
  // conversation yet, so the choice waits here and is saved once the chat is sent.
  /** @type {{ tone: string | null } | null} */
  let pendingChoice = null;
  /** Conversation IDs seen on connections while the page showed an unsent new chat. */
  const newChatIds = new Set();
  /** Whether each connection's chat was opened with Work IQ on, by slot name. */
  /** @type {Map<string, boolean>} */
  const connectionWorkIQ = new Map();
  const revisions = { conversationTones: 0, usage: 0, conversationReadings: 0 };
  let lastPublished = "";

  const post = data => window.postMessage({ channel: CHANNEL, direction: "to-page", ...data }, location.origin);
  const configure = () => post({ kind: "configure", tone: state.loaded ? core.DEFAULT_TONE : null });
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

  /** Keep only the newest entries of a saved per-conversation map. */
  const newest = entries => Object.fromEntries(Object.entries(entries)
    .sort(([, a], [, b]) => (b?.at || 0) - (a?.at || 0)).slice(0, core.MAX_SAVED_CONVERSATIONS));

  function touch(name) {
    const observed = slot(name);
    observed.at = Date.now();
    if (!isConversation(name)) return;
    dirty.add(name);
    persistTimer ??= setTimeout(persist, PERSIST_DELAY);
  }

  function persist() {
    persistTimer = null;
    if (state.failed) return;
    // Wait for the initial read so saved readings from other chats are kept.
    if (!state.loaded) { persistTimer = setTimeout(persist, PERSIST_DELAY); return; }
    const next = { ...state.readings };
    let changed = false;
    for (const id of [...dirty]) {
      const key = keys.get(id), observed = slots.get(id);
      if (!key) continue; // Saved once hashing finishes.
      dirty.delete(id);
      if (!observed) continue;
      // Warnings describe this tab's connection; they are not saved.
      const { warning, ...saved } = observed;
      next[key] = { ...saved, at: observed.at || Date.now() };
      changed = true;
    }
    if (!changed) return;
    state.readings = newest(next);
    chrome.storage.local.set({ conversationReadings: state.readings }).catch(() => { /* Readings stay in this tab. */ });
  }

  function saveUsage() {
    if (state.failed || !account.metering || !account.quotaAt) return;
    chrome.storage.local.set({ usage: { metering: account.metering, quotaAt: account.quotaAt } }).catch(() => { /* Tab keeps its reading. */ });
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
      if (dirty.has(id)) persistTimer ??= setTimeout(persist, PERSIST_DELAY);
      publishTab();
    }).catch(() => { /* Without a key this conversation uses the default model. */ })
      .finally(() => hashing.delete(id));
  }

  /** Tell the hook which tone this conversation uses (null = default). */
  function answerHook(id) {
    if (!state.loaded && !state.failed) return;
    const key = keys.get(id);
    if (!key && !state.failed) return;
    const saved = key && state.conversationTones[key]?.tone;
    const chosen = pendingChoice && newChatIds.has(id) ? pendingChoice.tone : null;
    post({ kind: "conversation", conversationId: id, tone: saved || chosen || null });
  }

  /** Save a new chat's pre-send choice under its conversation once the chat is sent. */
  function adoptPendingChoice(id) {
    const choice = pendingChoice, key = keys.get(id);
    if (!choice || !newChatIds.has(id)) return;
    pendingChoice = null;
    if (!key || !choice.tone || choice.tone === core.DEFAULT_TONE) return;
    const tones = newest({ ...state.conversationTones, [key]: { tone: choice.tone, at: Date.now() } });
    chrome.storage.local.set({ conversationTones: tones }).catch(() => { /* This tab keeps the choice for the session. */ });
    state.conversationTones = tones;
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
    if (!name && pendingChoice) return pendingChoice.tone || core.DEFAULT_TONE;
    const key = isConversation(name) ? keys.get(name) : undefined;
    return (key && state.conversationTones[key]?.tone) || core.DEFAULT_TONE;
  }

  /** Whether a conversation (or, for null, the unsent new chat) uses Work IQ; undefined if unknown. */
  function workIQFor(name) {
    if (!name) return [...newChatIds].reverse().map(id => connectionWorkIQ.get(id)).find(value => value !== undefined);
    const key = isConversation(name) ? keys.get(name) : undefined;
    return slots.get(name)?.workIQ ?? (key ? state.readings[key]?.workIQ : undefined) ?? connectionWorkIQ.get(name);
  }

  function snapshot() {
    const name = currentSlot();
    const key = isConversation(name) ? keys.get(/** @type {string} */ (name)) : undefined;
    const live = name ? slots.get(name) : undefined;
    const saved = key ? state.readings[key] : undefined;
    // The newer of this tab's readings and the saved ones (possibly from another tab).
    const readings = saved && (saved.at || 0) >= (live?.at || 0) ? saved : live;
    const observed = { ...account, ...readings };
    delete observed.at;
    const workIQ = workIQFor(name);
    if (workIQ === undefined) delete observed.workIQ; else observed.workIQ = workIQ;
    observed.warning = live?.warning || state.tabWarning;
    if (!observed.warning) delete observed.warning;
    return { hookReady: state.hookReady, observed, conversation: key ? { key } : null,
      newChat: name ? null : { tone: pendingChoice?.tone || null } };
  }

  // Keep the toolbar badge and icon in step with this tab.
  function publishTab(force = false) {
    const name = currentSlot();
    // Show the model a request really uses: Work IQ may fall back to the default.
    const chosen = effectiveTone(name);
    const workIQ = workIQFor(name) === true;
    const tone = core.requestFor(chosen, workIQ)?.model.tone ?? chosen;
    const message = { type: "M365_EXTENDER_TAB", tone, remaining: core.lowestAllowance(tone, account.metering), workIQ };
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
      const connection = slotName(message);
      connectionWorkIQ.delete(connection);
      connectionWorkIQ.set(connection, message.workIQ === true);
      for (const old of connectionWorkIQ.keys()) { if (connectionWorkIQ.size <= MAX_SLOTS) break; connectionWorkIQ.delete(old); }
      if (id && !sentSlots.has(id) && !core.conversationFromPath(location.pathname)) newChatIds.add(id);
      if (id) { if (keys.has(id)) answerHook(id); else ensureKey(id); }
    } else if (message.kind === "request" || (message.kind === "applied" && core.modelFor(message.tone))) {
      const name = slotName(message);
      activeSlot = name;
      sentSlots.add(name);
      adoptPendingChoice(name);
      newChatIds.delete(name);
      const observed = slot(name);
      const workIQ = message.workIQ === true;
      if (message.kind === "applied") { observed.appliedTone = core.modelFor(message.tone)?.tone; observed.workIQ = workIQ; touch(name); }
      else if (observed.workIQ !== workIQ) { observed.workIQ = workIQ; touch(name); }
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
          const metering = validMetering(message.metering);
          if (Object.keys(metering).length) {
            account.metering = { ...account.metering, ...metering };
            account.quotaAt = Date.now();
            saveUsage();
          }
        }
      }
      if (core.allowedResults.includes(message.result)) observed.result = message.result;
      if (core.allowedOrigins.includes(message.origin)) observed.origin = message.origin;
      if (core.count(message.conversation?.used) !== null && core.count(message.conversation?.limit) !== null) observed.conversation = { used: message.conversation.used, limit: message.conversation.limit };
      const expiresAt = core.parseExpiry(message.expiresAt);
      if (expiresAt) observed.expiresAt = expiresAt;
      touch(name);
    } else return;
    publishTab();
  });

  // Revalidate data from the page or storage, neither of which is an authenticated boundary.
  const validMetering = raw => core.parseMetering(Object.fromEntries(Object.entries(raw && typeof raw === "object" ? raw : {})
    .map(([key, value]) => [key, { remainingAllowance: value }])));

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

  /** @returns {ConversationReadings} */
  function validReadings(raw) {
    /** @type {ConversationReadings} */
    const valid = {};
    if (!raw || typeof raw !== "object") return valid;
    for (const [key, entry] of Object.entries(raw)) {
      if (!/^[0-9a-f]{32}$/.test(key) || !entry || typeof entry !== "object") continue;
      /** @type {Observed & { at: number }} */
      const reading = { at: Number(entry.at) || 0 };
      const applied = core.modelFor(entry.appliedTone);
      if (applied) reading.appliedTone = applied.tone;
      if (typeof entry.workIQ === "boolean") reading.workIQ = entry.workIQ;
      if (core.allowedResults.includes(entry.result)) reading.result = entry.result;
      if (core.allowedOrigins.includes(entry.origin)) reading.origin = entry.origin;
      const used = core.count(entry.conversation?.used), limit = core.count(entry.conversation?.limit);
      if (used !== null && limit !== null) reading.conversation = { used, limit };
      const expiresAt = core.parseExpiry(entry.expiresAt);
      if (expiresAt) reading.expiresAt = expiresAt;
      valid[key] = reading;
    }
    return valid;
  }

  function applyUsage(raw) {
    const quotaAt = Number(raw?.quotaAt) || 0;
    // Keep this tab's reading if it is newer than what was saved.
    if (!quotaAt || quotaAt < (account.quotaAt || 0)) return;
    account.metering = validMetering(raw.metering);
    account.quotaAt = quotaAt;
  }

  function applyConversationTones(raw) {
    state.conversationTones = validTones(raw);
    for (const id of keys.keys()) answerHook(id);
  }

  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local") return;
    for (const key of Object.keys(revisions)) if (changes[key]) revisions[key]++;
    if (changes.conversationTones) applyConversationTones(changes.conversationTones.newValue);
    if (changes.usage) applyUsage(changes.usage.newValue);
    if (changes.conversationReadings) state.readings = validReadings(changes.conversationReadings.newValue);
    publishTab();
  });
  chrome.runtime.onMessage.addListener((message, sender, respond) => {
    if (sender.id !== chrome.runtime.id) return;
    if (message?.type === "M365_EXTENDER_SNAPSHOT") respond(snapshot());
    else if (message?.type === "M365_EXTENDER_NEW_CHAT_TONE") {
      // Only an unsent new chat can take a choice before it has a conversation.
      const ok = currentSlot() === null && (message.tone === null || !!core.modelFor(message.tone));
      if (ok) {
        pendingChoice = { tone: core.modelFor(message.tone)?.tone ?? null };
        for (const id of newChatIds) answerHook(id);
        publishTab();
      }
      respond({ ok });
    }
  });
  // Single-page navigation (sidebar conversations) changes the current conversation.
  // The worker resets a tab's badge when it starts loading, so re-send after
  // navigation settles rather than relying on the de-duplicated state.
  const republish = () => setTimeout(() => publishTab(true), 300);
  // Moving to another conversation abandons a choice made for the unsent new chat.
  const abandonChoice = () => {
    const id = core.conversationFromPath(location.pathname);
    if (id && !newChatIds.has(id)) { pendingChoice = null; newChatIds.clear(); }
    republish();
  };
  try { /** @type {any} */ (window).navigation?.addEventListener("currententrychange", abandonChoice); } catch { /* Older Chrome. */ }
  window.addEventListener("pageshow", republish);

  const initialRevisions = { ...revisions };
  chrome.storage.local.get(Object.keys(revisions)).then(settings => {
    // A cross-tab change that arrived during this read is newer; keep it.
    const fresh = key => revisions[key] === initialRevisions[key];
    if (fresh("conversationTones")) state.conversationTones = validTones(settings.conversationTones);
    if (fresh("usage")) applyUsage(settings.usage);
    if (fresh("conversationReadings")) state.readings = validReadings(settings.conversationReadings);
    state.loaded = true;
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
