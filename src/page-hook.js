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
  const sockets = new WeakMap();
  const waiting = new Set();
  let selectedTone = null;
  let initialized = false;
  let nextSocketId = 0;

  function emit(kind, data = {}) {
    post({ channel: CHANNEL, direction: "from-page", kind, ...data }, origin);
  }

  function watch(socket) {
    if (sockets.has(socket)) return sockets.get(socket);
    const info = core.connectionInfo(socket.url);
    if (!info) return null;
    const state = { ...info, socketId: ++nextSocketId, tone: null, queue: [], timer: null };
    sockets.set(socket, state);
    emit("connection", { ...info, socketId: state.socketId });
    nativeAddEvent.call(socket, "message", event => {
      for (const frame of core.parseRecords(event.data)) {
        for (const metadata of core.readMetadata(frame)) {
          emit("metadata", { ...metadata, scenario: state.scenario, tone: state.tone, socketId: state.socketId });
        }
        if (frame.type === 3 && typeof frame.error === "string") emit("warning", { code: "request-rejected", socketId: state.socketId });
      }
    });
    nativeAddEvent.call(socket, "close", () => {
      clearTimeout(state.timer);
      state.queue = [];
      waiting.delete(socket);
    });
    return state;
  }

  function sendNow(socket, state, data) {
    const rewritten = core.rewrite(data, selectedTone);
    const result = nativeSend.call(socket, rewritten.data);
    if (rewritten.chats) {
      emit("request", { scenario: state.scenario, socketId: state.socketId });
    }
    if (rewritten.tones.length) {
      state.tone = rewritten.tones.at(-1);
      emit("applied", { tone: state.tone, scenario: state.scenario, socketId: state.socketId });
    }
    return result;
  }

  function flush(socket, state) {
    clearTimeout(state.timer);
    state.timer = null;
    waiting.delete(socket);
    const queue = state.queue.splice(0);
    if (socket.readyState !== NativeWebSocket.OPEN) {
      emit("warning", { code: "connection-closed" });
      return;
    }
    for (const data of queue) {
      try { sendNow(socket, state, data); }
      catch { emit("warning", { code: "send-failed" }); }
    }
  }

  window.addEventListener("message", event => {
    const message = event.data;
    if (event.source !== window || event.origin !== origin || message?.channel !== CHANNEL || message.direction !== "to-page") return;
    if (message.kind === "hello") emit("ready");
    if (message.kind !== "configure" || !(core.modelFor(message.tone) || message.tone === null)) return;
    selectedTone = message.tone;
    initialized = true;
    emit("configured", { tone: selectedTone });
    for (const socket of waiting) flush(socket, sockets.get(socket));
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
    // Once configured, skip this extra parse; sendNow parses each payload once.
    if (state.queue.length || (!initialized && core.parseRecords(data).some(core.isChat))) {
      state.queue.push(data);
      waiting.add(this);
      if (!state.timer) {
        state.timer = setTimeout(() => {
          // Fail open after 1.5 s rather than breaking Copilot when storage fails.
          emit("warning", { code: "settings-unavailable" });
          flush(this, state);
        }, 1500);
      }
      return undefined;
    }
    if (typeof data !== "string") emit("warning", { code: "binary-unsupported" });
    return sendNow(this, state, data);
  };
  emit("ready");
})();
