const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { load, settle, createStore, clone } = require('./harness.cjs');

function popup(store, snapshot = { hookReady: true, observed: { scenario: 'OfficeWebPaidCopilot', metering: { GPT61SolQueryDaily: 6 } } }) {
  let callbacks, state;
  const chrome = { ...store.chrome, tabs: { async query() { return [{ id: 7 }]; }, async sendMessage(id, message) {
    assert.equal(id, 7); assert.equal(message.type, 'M365_EXTENDER_SNAPSHOT'); return snapshot;
  } } };
  const context = vm.createContext({ chrome, document: { getElementById() { return {}; } }, setInterval() {},
    M365ExtenderPanel: { mount(host, handlers) { callbacks = handlers; return { update(next) { state = clone(next); } }; } } });
  load(context, 'core.js'); load(context, 'popup.js');
  return { callbacks, chrome, state: () => state };
}
test('popup shows active-tab metadata and saves model plus independent panel visibility', async () => {
  const store = createStore({ selectedTone: 'Gpt_5_5_Reasoning', inPageEnabled: true });
  const view = popup(store); await settle();
  assert.equal(view.state().tone, 'Gpt_5_5_Reasoning'); assert.equal(view.state().observed.metering.GPT61SolQueryDaily, 6);
  await view.callbacks.onSelect('Gpt_6_Sol_Reasoning'); await view.callbacks.onVisibility(false); await settle();
  assert.equal(store.data.selectedTone, 'Gpt_6_Sol_Reasoning'); assert.equal(store.data.inPageEnabled, false);
  assert.equal(view.state().tone, 'Gpt_6_Sol_Reasoning'); assert.equal(view.state().inPageEnabled, false);
});
test('popup startup resolves changed and unchanged preferences independently', async () => {
  const store = createStore(); let resolve;
  store.chrome.storage.local.get = () => new Promise(r => { resolve = r; });
  const view = popup(store);
  await store.chrome.storage.local.set({ selectedTone: 'Gpt_5_6_Reasoning' });
  resolve({ selectedTone: 'Magic', inPageEnabled: false }); await settle();
  assert.equal(view.state().tone, 'Gpt_5_6_Reasoning'); assert.equal(view.state().inPageEnabled, false);
});
test('popup retains saved preference and reports a failed settings write', async () => {
  const store = createStore({ selectedTone: 'Magic' }); const view = popup(store); await settle();
  store.chrome.storage.local.set = async () => { throw new Error('No write'); };
  await view.callbacks.onSelect('Gpt_6_Reasoning'); await settle();
  assert.equal(view.state().tone, 'Magic'); assert.match(view.state().error, /Could not save/);
});
