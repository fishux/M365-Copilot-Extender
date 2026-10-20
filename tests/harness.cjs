const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const load = (ctx, file) => vm.runInContext(fs.readFileSync(path.join(__dirname, '../src', file), 'utf8'), ctx, { filename: file });
const loadBundle = (ctx, file) => vm.runInContext(fs.readFileSync(path.join(__dirname, '../extension', file), 'utf8'), ctx, { filename: file });
const settle = async () => { for (let i = 0; i < 12; i++) await Promise.resolve(); };
const clone = value => JSON.parse(JSON.stringify(value));
const chat = (tone = 'Magic', type = 4) => JSON.stringify({ type, target: 'chat', arguments: [{ tone, message: { text: 'PRIVATE PROMPT' }, optionsSets: ['keep-me'] }] }) + '\x1e';
const endpoint = 'wss://substrate.office.com/m365Copilot/Chathub/PRIVATE_ACCOUNT?access_token=PRIVATE_TOKEN&scenario=OfficeWebPaidCopilot';

function createStore(initial = {}) {
  const data = { ...initial }, listeners = [];
  const chrome = {
    storage: {
      local: {
        async get() { return { ...data }; },
        async set(patch) {
          const changes = {};
          for (const [key, value] of Object.entries(patch)) { changes[key] = { oldValue: data[key], newValue: value }; data[key] = value; }
          for (const listener of listeners) listener(changes, 'local');
        }
      },
      onChanged: { addListener(listener) { listeners.push(listener); } }
    },
    runtime: { id: 'test-extension', onMessage: { addListener() {} } }
  };
  return { chrome, data };
}

function environment(store = createStore(), { bundled = false } = {}) {
  const messages = [], timers = new Map(), subscribers = [], contents = [];
  const origin = 'https://m365.cloud.microsoft';
  let timerId = 0;
  class FakeSocket extends EventTarget {
    static OPEN = 1; static CONNECTING = 0; static CLOSED = 3;
    constructor(url) { super(); this.url = url; this.sent = []; this.readyState = 1; }
    send(data) { if (this.readyState === 0) throw new Error('InvalidStateError'); this.sent.push(data); }
    receive(data) { this.dispatchEvent(new MessageEvent('message', { data })); }
    close() { this.readyState = 3; this.dispatchEvent(new Event('close')); }
  }
  function world() {
    const sandbox = { console, URL, EventTarget, ArrayBuffer, Blob, location: { origin, href: origin + '/chat' },
      setTimeout(fn, delay) { timers.set(++timerId, { fn, delay }); return timerId; },
      clearTimeout(id) { timers.delete(id); }, setInterval() {},
      addEventListener(type, listener) { if (type === 'message') subscribers.push({ sandbox, listener }); },
      postMessage(message) {
        messages.push(clone(message));
        for (const { sandbox: recipient, listener } of subscribers) queueMicrotask(() => listener({ source: vm.runInContext('window', recipient), origin, data: clone(message) }));
      }
    };
    sandbox.window = sandbox;
    return vm.createContext(sandbox);
  }
  const main = world(); main.WebSocket = FakeSocket;
  if (bundled) loadBundle(main, 'page-hook.js');
  else { load(main, 'core.js'); load(main, 'page-hook.js'); }
  function addContent() {
    const isolated = world();
    let snapshotListener;
    isolated.chrome = { ...store.chrome, runtime: { ...store.chrome.runtime, onMessage: { addListener(listener) { snapshotListener = listener; } } } };
    // No document is provided: the bridge must run without touching page DOM.
    if (bundled) loadBundle(isolated, 'content.js');
    else { load(isolated, 'core.js'); load(isolated, 'content.js'); }
    isolated.snapshot = () => { let result; snapshotListener({ type: 'M365_EXTENDER_SNAPSHOT' }, { id: store.chrome.runtime.id }, value => { result = clone(value); }); return result; };
    contents.push(isolated);
    return isolated;
  }

  return { main, addContent, store, messages, contents, FakeSocket, timers,
    timeout(delay) { for (const [id, timer] of [...timers]) if (timer.delay === delay) { timers.delete(id); timer.fn(); } } };
}

module.exports = { load, loadBundle, settle, clone, chat, endpoint, createStore, environment };
