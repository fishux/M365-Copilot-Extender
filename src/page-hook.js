// @ts-check
/* Runs at document_start in the page's MAIN world. */
(() => {
  "use strict";
  const core = globalThis.M365ExtenderCore;
  const CHANNEL = "m365-copilot-extender-v1";
  const NativeWebSocket = window.WebSocket;
  const nativeSend = NativeWebSocket.prototype.send;
  const nativeAddEvent = EventTarget.prototype.addEventListener;
  const post = window.postMessage.bind(window);
  const origin = location.origin;
  /** @type {WeakMap<WebSocket, SocketState>} */
  const sockets = new WeakMap();
  /** @type {Set<WebSocket>} */
  const waiting = new Set();
  // Per-conversation overrides from the isolated script: conversationId -> tone,
  // or null when the conversation uses the default. Absent means "not yet known".
  /** @type {Map<string, string | null>} */
  const conversationTones = new Map();
  /** @type {string | null} */
  let selectedTone = null;
  let initialized = false;
  let nextSocketId = 0;

  function emit(kind, data = {}) {
    post({ channel: CHANNEL, direction: "from-page", kind, ...data }, origin);
  }

  /** @param {WebSocket} socket @returns {SocketState | null} */
  function watch(socket) {
    if (sockets.has(socket)) return sockets.get(socket) || null;
    const info = core.connectionInfo(socket.url);
    if (!info) return null;
    /** @type {SocketState} */
    const state = { ...info, socketId: ++nextSocketId, tone: null, queue: [], timer: null };
    sockets.set(socket, state);
    emit("connection", { ...info, socketId: state.socketId });
    nativeAddEvent.call(socket, "message", event => {
      for (const frame of core.parseRecords(event.data)) {
        for (const metadata of core.readMetadata(frame)) {
          emit("metadata", { ...metadata, ...tag(state), tone: state.tone });
        }
        if (frame.type === 3 && typeof frame.error === "string") emit("warning", { code: "request-rejected", ...tag(state) });
      }
    });
    nativeAddEvent.call(socket, "close", () => {
      clearTimeout(state.timer ?? undefined);
      state.queue = [];
      waiting.delete(socket);
    });
    return state;
  }

  /** @param {SocketState} state */
  const tag = state => ({ scenario: state.scenario, socketId: state.socketId, conversationId: state.conversationId });

  /** A chat must wait until both the default and this conversation's override are known. */
  function ready(state) {
    return initialized && (!state.conversationId || conversationTones.has(state.conversationId));
  }

  /** @param {SocketState} state */
  function toneFor(state) {
    return (state.conversationId && conversationTones.get(state.conversationId)) || selectedTone;
  }

  function sendNow(socket, state, data) {
    const rewritten = core.rewrite(data, toneFor(state));
    const result = nativeSend.call(socket, rewritten.data);
    if (rewritten.chats) emit("request", tag(state));
    if (rewritten.tones.length) {
      state.tone = rewritten.tones.at(-1) ?? null;
      emit("applied", { ...tag(state), tone: state.tone });
    }
    return result;
  }

  function flush(socket, state) {
    clearTimeout(state.timer ?? undefined);
    state.timer = null;
    waiting.delete(socket);
    const queue = state.queue.splice(0);
    if (socket.readyState !== NativeWebSocket.OPEN) {
      emit("warning", { code: "connection-closed", ...tag(state) });
      return;
    }
    for (const data of queue) {
      try { sendNow(socket, state, data); }
      catch { emit("warning", { code: "send-failed", ...tag(state) }); }
    }
  }

  function flushReady() {
    for (const socket of waiting) {
      const state = sockets.get(socket);
      if (state && ready(state)) flush(socket, state);
    }
  }

  window.addEventListener("message", event => {
    const message = event.data;
    if (event.source !== window || event.origin !== origin || message?.channel !== CHANNEL || message.direction !== "to-page") return;
    if (message.kind === "hello") emit("ready");
    if (message.kind === "configure" && (core.modelFor(message.tone) || message.tone === null)) {
      selectedTone = core.modelFor(message.tone)?.tone ?? null;
      initialized = true;
      emit("configured", { tone: selectedTone });
      flushReady();
    }
    if (message.kind === "conversation") {
      const id = core.safeConversationId(message.conversationId);
      if (!id || !(core.modelFor(message.tone) || message.tone === null)) return;
      conversationTones.set(id, core.modelFor(message.tone)?.tone ?? null);
      flushReady();
    }
  });

  // Proxy preserves static constants, prototype, instanceof, and subclassing.
  window.WebSocket = new Proxy(NativeWebSocket, {
    construct(target, args, newTarget) {
      const socket = Reflect.construct(target, args, newTarget);
      try { watch(socket); } catch { /* Never prevent socket construction. */ }
      return socket;
    }
  });

  NativeWebSocket.prototype.send = function (data) {
    let state;
    try { state = watch(this); } catch { return nativeSend.call(this, data); }
    // Preserve native errors (including CONNECTING / invalid receiver).
    if (!state || this.readyState !== NativeWebSocket.OPEN) return nativeSend.call(this, data);
    // Once ready, skip this extra parse; sendNow parses each payload once.
    if (state.queue.length || (!ready(state) && core.parseRecords(data).some(core.isChat))) {
      const socketState = state;
      socketState.queue.push(data);
      waiting.add(this);
      if (!socketState.timer) {
        socketState.timer = setTimeout(() => {
          // Fail open after 1.5 s rather than breaking Copilot when storage fails.
          if (!initialized) {
            emit("warning", { code: "settings-unavailable", ...tag(socketState) });
            flush(this, socketState);
          } else {
            // Default known, conversation override not: send with the default,
            // and stop holding later turns until the override arrives.
            if (socketState.conversationId) conversationTones.set(socketState.conversationId, null);
            flush(this, socketState);
            emit("warning", { code: "conversation-unavailable", ...tag(socketState) });
          }
        }, 1500);
      }
      return undefined;
    }
    if (typeof data !== "string") emit("warning", { code: "binary-unsupported", ...tag(state) });
    return sendNow(this, state, data);
  };
  emit("ready");
})();
