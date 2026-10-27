const { test } = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const vm = require('node:vm');
const { environment, createStore, settle, endpoint, chat, load, clone } = require('./harness.cjs');

const A = '11111111-aaaa-4aaa-8aaa-111111111111';
const B = '22222222-bbbb-4bbb-8bbb-222222222222';
const at = id => `${endpoint}&ConversationId=${id}`;
const keyOf = id => crypto.createHash('sha256').update(`m365-copilot-extender:${id}`).digest('hex').slice(0, 32);
const toneSent = (socket, index = 0) => JSON.parse(socket.sent[index].split('\x1e')[0]).arguments[0].tone;
const reply = (socket, item) => socket.receive(JSON.stringify({ type: 2, item }) + '\x1e');
// Hashing is async (WebCrypto), so give it real turns of the event loop.
const flush = async () => { for (let i = 0; i < 3; i++) { await settle(); await new Promise(resolve => setTimeout(resolve, 5)); } await settle(); };

test('a saved per-conversation model applies only to that conversation', async () => {
  const store = createStore({ conversationTones: { [keyOf(A)]: { tone: 'Gpt_6_Reasoning', at: 1 } } });
  const env = environment(store); env.addContent(); await flush();
  const a = new env.main.WebSocket(at(A)), b = new env.main.WebSocket(at(B));
  a.send(chat()); b.send(chat());
  // Held until the conversation lookup answers, then sent in order.
  assert.equal(a.sent.length, 0);
  await flush();
  assert.equal(toneSent(a), 'Gpt_6_Reasoning');
  assert.equal(toneSent(b), 'Gpt_6_Sol_Reasoning');
});

test('changing a conversation model in another tab applies to the next turn', async () => {
  const store = createStore();
  const env = environment(store); env.addContent(); await flush();
  const socket = new env.main.WebSocket(at(A)); await flush();
  socket.send(chat()); assert.equal(toneSent(socket), 'Gpt_6_Sol_Reasoning');
  await store.chrome.storage.local.set({ conversationTones: { [keyOf(A)]: { tone: 'Gpt_5_6_Reasoning', at: 2 } } }); await flush();
  socket.send(chat()); assert.equal(toneSent(socket, 1), 'Gpt_5_6_Reasoning');
  await store.chrome.storage.local.set({ conversationTones: {} }); await flush();
  socket.send(chat()); assert.equal(toneSent(socket, 2), 'Gpt_6_Sol_Reasoning');
});

test('raw conversation IDs never reach storage or the popup snapshot', async () => {
  const store = createStore({ conversationTones: { [keyOf(A)]: { tone: 'Gpt_6_Reasoning', at: 1 } } });
  const env = environment(store, { pathname: `/chat/conversation/${A}` });
  const content = env.addContent(); await flush();
  const socket = new env.main.WebSocket(at(A)); socket.send(chat()); await flush();
  const snapshot = content.snapshot();
  assert.deepEqual(snapshot.conversation, { key: keyOf(A) });
  assert.ok(!JSON.stringify(snapshot).includes(A));
  assert.ok(!JSON.stringify(store.data).includes(A));
});

test('readings follow the conversation shown in the page, while allowances stay account-wide', async () => {
  const env = environment(createStore(), { pathname: '/chat' });
  const content = env.addContent(); await flush();
  const a = new env.main.WebSocket(at(A)); a.send(chat()); await flush();
  reply(a, { result: { value: 'Success' }, throttling: { numUserMessagesInConversation: 7, maxNumUserMessagesInConversation: 600, metering: { GPT61SolQueryDaily: { remainingAllowance: 30 } } } });
  await flush();
  // Before Copilot navigates to the new thread, the just-sent chat is current.
  assert.equal(content.snapshot().observed.conversation.used, 7);
  env.location.pathname = `/chat/conversation/${A}`;
  const b = new env.main.WebSocket(at(B)); b.send(chat()); await flush();
  reply(b, { throttling: { numUserMessagesInConversation: 2, maxNumUserMessagesInConversation: 600, metering: { GPT61SolQueryDaily: { remainingAllowance: 29 } } } });
  await flush();
  // The page still shows A: its own capacity, the latest account allowance.
  assert.equal(content.snapshot().observed.conversation.used, 7);
  assert.equal(content.snapshot().observed.metering.GPT61SolQueryDaily, 29);
  env.location.pathname = `/chat/conversation/${B}`;
  assert.equal(content.snapshot().observed.conversation.used, 2);
  // A new, unsent chat shows no conversation readings but keeps allowances.
  env.location.pathname = '/chat';
  assert.equal(content.snapshot().observed.conversation, undefined);
  assert.equal(content.snapshot().observed.metering.GPT61SolQueryDaily, 29);
  // Reopening A from the sidebar restores its readings without a new reply.
  env.location.pathname = `/chat/conversation/${A}`;
  assert.equal(content.snapshot().observed.conversation.used, 7);
});

test('a new socket in the same conversation keeps that conversation’s readings', async () => {
  const env = environment(createStore(), { pathname: `/chat/conversation/${A}` });
  const content = env.addContent(); await flush();
  const first = new env.main.WebSocket(at(A)); first.send(chat()); await flush();
  reply(first, { throttling: { numUserMessagesInConversation: 3, maxNumUserMessagesInConversation: 600 } }); await flush();
  const second = new env.main.WebSocket(at(A)); second.send(chat()); await flush();
  assert.equal(content.snapshot().observed.conversation.used, 3);
});

test('a lost conversation lookup fails open once with a tagged warning', async () => {
  const env = environment(createStore());
  env.addContent(); await flush();
  // Simulate the isolated script never answering for this conversation.
  env.main.crypto = undefined;
  const socket = new env.main.WebSocket(at('33333333-cccc-4ccc-8ccc-333333333333'));
  const listener = env.contents[0];
  listener.crypto = { subtle: { digest: () => new Promise(() => {}) } };
  socket.send(chat()); assert.equal(socket.sent.length, 0);
  env.timeout(1500);
  assert.equal(toneSent(socket), 'Gpt_6_Sol_Reasoning');
  const warning = env.messages.find(message => message.code === 'conversation-unavailable');
  assert.equal(warning.socketId, 1);
  // Later turns are no longer held.
  socket.send(chat()); assert.equal(socket.sent.length, 2);
});

test('every hook warning carries its socket and conversation', async () => {
  const env = environment(); const socket = new env.main.WebSocket(at(A));
  socket.send(chat()); env.timeout(1500); // settings-unavailable
  const closing = new env.main.WebSocket(endpoint); closing.send(chat());
  closing.readyState = 3; env.timeout(1500); // settings-unavailable, connection-closed
  const other = new env.main.WebSocket(endpoint); other.send(new Uint8Array([1])); // binary-unsupported
  assert.deepEqual(new Set(env.messages.filter(message => message.kind === 'warning').map(message => message.code)),
    new Set(['settings-unavailable', 'connection-closed', 'binary-unsupported']));
  for (const message of env.messages.filter(message => message.kind === 'warning')) assert.equal(typeof message.socketId, 'number');
  assert.equal(env.messages.find(message => message.code === 'settings-unavailable').conversationId, A);
});

test('Auto matches either casing and leaves an equivalent frame untouched', async () => {
  const context = vm.createContext({ URL }); load(context, 'core.js');
  const core = context.M365ExtenderCore;
  assert.equal(core.modelFor('MAGIC').tone, 'Magic');
  assert.equal(core.modelFor('gpt_6_reasoning').tone, 'Gpt_6_Reasoning');
  const lower = chat('magic');
  assert.equal(core.rewrite(lower, 'Magic').data, lower);
  assert.deepEqual(clone(core.rewrite(lower, 'Magic').tones), ['Magic']);
  // A saved lowercase choice is sent in Copilot's own casing.
  const store = createStore({ conversationTones: { [keyOf(A)]: { tone: 'magic', at: 1 } } });
  const env = environment(store); env.addContent(); await flush();
  const socket = new env.main.WebSocket(at(A)); socket.send(chat('Gpt_6_Reasoning')); await flush();
  assert.equal(toneSent(socket), 'Magic');
});

test('the tab reports its effective model and lowest allowance for the toolbar', async () => {
  const store = createStore({ conversationTones: { [keyOf(A)]: { tone: 'Gpt_61_Sol_Reasoning', at: 1 } } });
  const env = environment(store, { pathname: `/chat/conversation/${A}` }); env.addContent(); await flush();
  const socket = new env.main.WebSocket(at(A)); socket.send(chat()); await flush();
  reply(socket, { throttling: { metering: { GPT61SolQueryDaily: { remainingAllowance: 18 }, GPT61SolQueryWeekly: { remainingAllowance: 9 } } } });
  await flush();
  assert.deepEqual(store.tabMessages.at(-1), { type: 'M365_EXTENDER_TAB', tone: 'Gpt_61_Sol_Reasoning', remaining: 9, workIQ: false });
  await store.chrome.storage.local.set({ conversationTones: { [keyOf(A)]: { tone: 'Magic', at: 3 } } }); await flush();
  assert.deepEqual(store.tabMessages.at(-1), { type: 'M365_EXTENDER_TAB', tone: 'Magic', remaining: null, workIQ: false });
});

function popup(store, snapshot) {
  let callbacks, state;
  const chrome = { ...store.chrome, tabs: { async query() { return [{ id: 7 }]; }, async sendMessage() { return snapshot; } } };
  const context = vm.createContext({ chrome, document: { getElementById() { return {}; } }, setInterval() {} });
  load(context, 'core.js');
  context.M365ExtenderPanel = { mount(host, handlers) { callbacks = handlers; return { update(next) { state = clone(next); } }; } };
  load(context, 'popup.js');
  return { callbacks, state: () => state };
}

test('popup keeps a chosen model for the chat under its hashed key; choosing the default returns it to the default', async () => {
  const key = keyOf(A);
  const store = createStore();
  const view = popup(store, { hookReady: true, observed: {}, conversation: { key } }); await settle();
  assert.equal(view.state().conversationTone, null);
  assert.equal(view.state().tone, 'Gpt_6_Sol_Reasoning');
  await view.callbacks.onConversationSelect('Gpt_6_Reasoning'); await settle();
  assert.equal(store.data.conversationTones[key].tone, 'Gpt_6_Reasoning');
  assert.equal(view.state().conversationTone, 'Gpt_6_Reasoning');
  assert.equal(store.data.selectedTone, undefined);
  await view.callbacks.onConversationSelect('Gpt_6_Sol_Reasoning'); await settle();
  assert.deepEqual(clone(store.data.conversationTones), {});
  assert.equal(view.state().conversationTone, null);
});

test('popup keeps only the 2000 most recent conversation choices', async () => {
  const old = Object.fromEntries(Array.from({ length: 2005 }, (_, i) => [i.toString(16).padStart(32, '0'), { tone: 'Magic', at: i }]));
  const store = createStore({ conversationTones: old });
  const view = popup(store, { hookReady: true, observed: {}, conversation: { key: keyOf(B) } }); await settle();
  await view.callbacks.onConversationSelect('Gpt_5_5_Reasoning'); await settle();
  const saved = store.data.conversationTones;
  assert.equal(Object.keys(saved).length, 2000);
  assert.equal(saved[keyOf(B)].tone, 'Gpt_5_5_Reasoning');
  assert.equal(saved[(0).toString(16).padStart(32, '0')], undefined);
});

test('popup cannot set a conversation model outside a known conversation', async () => {
  const store = createStore();
  const view = popup(store, { hookReady: true, observed: {}, conversation: null }); await settle();
  await view.callbacks.onConversationSelect('Gpt_6_Reasoning'); await settle();
  assert.equal(store.data.conversationTones, undefined);
});
