const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const load = (ctx, file) => vm.runInContext(fs.readFileSync(path.join(__dirname, '../extension', file), 'utf8'), ctx, { filename: file });
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

function environment(store = createStore()) {
  const messages = [], timers = new Map(), subscribers = [], panels = [], hosts = [];
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
    const sandbox = { console, URL, EventTarget, ArrayBuffer, Blob, location: { origin },
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
  load(main, 'core.js'); load(main, 'page-hook.js');
  function addContent() {
    const isolated = world(); isolated.chrome = store.chrome;
    isolated.document = { body: { append(host) { hosts.push(host); } }, createElement() { return { style: {} }; } };
    load(isolated, 'core.js');
    isolated.M365ExtenderPanel = { mount(host, callbacks) {
      const panel = { host, callbacks, state: null, update(state) { this.state = clone(state); } };
      panels.push(panel); return panel;
    } };
    load(isolated, 'content.js');
    return isolated;
  }
  return { main, addContent, store, messages, panels, hosts, FakeSocket, timers,
    timeout(delay) { for (const [id, timer] of [...timers]) if (timer.delay === delay) { timers.delete(id); timer.fn(); } } };
}

module.exports = { load, settle, clone, chat, endpoint, createStore, environment };
