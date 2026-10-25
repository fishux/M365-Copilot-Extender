const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { load, settle, createStore, clone } = require('./harness.cjs');

function popup(store, snapshot) {
  let callbacks, state;
  const chrome = { ...store.chrome, tabs: { async query() { return [{ id: 7 }]; }, async sendMessage(id, message) {
    assert.equal(id, 7); assert.equal(message.type, 'M365_EXTENDER_SNAPSHOT'); return await snapshot;
  } } };
  const context = vm.createContext({ chrome, document: { getElementById() { return {}; } }, setInterval() {},
  });
  load(context, 'core.js');
  context.M365ExtenderPanel = { mount(host, handlers) { callbacks = handlers; return { update(next) { state = clone(next); } }; } };
  load(context, 'popup.js');
  return { callbacks, chrome, state: () => state };
}
const KEY = 'a'.repeat(32);
const inChat = { hookReady: true, observed: { scenario: 'OfficeWebPaidCopilot', metering: { GPT61SolQueryDaily: 6 }, quotaAt: 100 }, conversation: { key: KEY } };

test('popup shows active-tab metadata and saves a model for the chat, not a default', async () => {
  const store = createStore({ selectedTone: 'Gpt_5_5_Reasoning', inPageEnabled: true });
  const view = popup(store, inChat); await settle();
  assert.equal(view.state().tone, 'Gpt_6_Sol_Reasoning'); assert.equal(view.state().observed.metering.GPT61SolQueryDaily, 6);
  assert.equal(view.callbacks.onSelect, undefined); assert.equal(view.callbacks.onVisibility, undefined);
  await view.callbacks.onConversationSelect('Gpt_5_6_Reasoning'); await settle();
  assert.equal(store.data.conversationTones[KEY].tone, 'Gpt_5_6_Reasoning');
  assert.equal(store.data.selectedTone, 'Gpt_5_5_Reasoning'); assert.equal(store.data.inPageEnabled, true);
  assert.equal(view.state().conversationTone, 'Gpt_5_6_Reasoning'); assert.equal(view.state().inPageEnabled, undefined);
});
test('popup startup resolves changed and unchanged settings independently', async () => {
  const store = createStore(); let resolve;
  store.chrome.storage.local.get = () => new Promise(r => { resolve = r; });
  const view = popup(store, inChat);
  await store.chrome.storage.local.set({ conversationTones: { [KEY]: { tone: 'Gpt_5_6_Reasoning', at: 2 } } });
  resolve({ conversationTones: { [KEY]: { tone: 'Magic', at: 1 } }, usage: { metering: { GPT61SolQueryDaily: 40 }, quotaAt: 200 } }); await settle();
  assert.equal(view.state().conversationTone, 'Gpt_5_6_Reasoning');
  assert.equal(view.state().observed.metering.GPT61SolQueryDaily, 40);
});
test('popup retains the chat model and reports a failed settings write', async () => {
  const store = createStore({ conversationTones: { [KEY]: { tone: 'Magic', at: 1 } } }); const view = popup(store, inChat); await settle();
  store.chrome.storage.local.set = async () => { throw new Error('No write'); };
  await view.callbacks.onConversationSelect('Gpt_6_Reasoning'); await settle();
  assert.equal(view.state().conversationTone, 'Magic'); assert.match(view.state().error, /Could not save/);
});
test('popup shows saved allowances outside Copilot and prefers the newer reading', async () => {
  const store = createStore({ usage: { metering: { GPT61SolQueryDaily: 12 }, quotaAt: 50 } });
  const offTab = popup(store, Promise.reject(new Error('No tab'))); await settle();
  assert.equal(offTab.state().connected, false);
  assert.equal(offTab.state().observed.metering.GPT61SolQueryDaily, 12);
  // The tab's reading (quotaAt 100) is newer than the saved one.
  const onTab = popup(store, inChat); await settle();
  assert.equal(onTab.state().observed.metering.GPT61SolQueryDaily, 6);
});
