const { test } = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const { settle, chat, endpoint, environment, createStore, clone } = require('./harness.cjs');

const A = '11111111-aaaa-4aaa-8aaa-111111111111';
const at = id => `${endpoint}&ConversationId=${id}`;
const keyOf = id => crypto.createHash('sha256').update(`m365-copilot-extender:${id}`).digest('hex').slice(0, 32);
const toneSent = (socket, index = 0) => JSON.parse(socket.sent[index].split('\x1e')[0]).arguments[0].tone;
// Hashing is async (WebCrypto), so give it real turns of the event loop.
const flush = async () => { for (let i = 0; i < 3; i++) { await settle(); await new Promise(resolve => setTimeout(resolve, 5)); } await settle(); };

test('the fixed default applies before first chat; native identity, constants and subclassing survive', async () => {
  const env = environment(createStore({ selectedTone: 'Gpt_5_6_Reasoning' })); // Legacy key is ignored.
  env.addContent(); await settle();
  const socket = new env.main.WebSocket(endpoint); socket.send(chat());
  await settle();
  assert.equal(JSON.parse(socket.sent[0].slice(0, -1)).arguments[0].tone, 'Gpt_6_Sol_Reasoning');
  assert.ok(socket instanceof env.main.WebSocket); assert.ok(socket instanceof env.FakeSocket);
  assert.equal(env.main.WebSocket.OPEN, 1);
  class SubSocket extends env.main.WebSocket {} assert.ok(new SubSocket(endpoint) instanceof SubSocket);
  assert.equal(env.contents[0].snapshot().observed.appliedTone, 'Gpt_6_Sol_Reasoning');
});
test('cold-start queues the chat and following records in order until storage resolves', async () => {
  const store = createStore({ selectedTone: 'Gpt_6_Sol_Reasoning' });
  let resolve; store.chrome.storage.local.get = () => new Promise(r => { resolve = r; });
  const env = environment(store); env.addContent();
  const socket = new env.main.WebSocket(endpoint);
  socket.send('{}\x1e'); socket.send(chat()); socket.send('{"type":6}\x1e');
  assert.equal(socket.sent.length, 1);
  resolve({ selectedTone: 'Gpt_6_Sol_Reasoning' }); await settle();
  assert.equal(socket.sent.length, 3);
  assert.equal(JSON.parse(socket.sent[1].slice(0, -1)).arguments[0].tone, 'Gpt_6_Sol_Reasoning');
  assert.equal(socket.sent[2], '{"type":6}\x1e');
});
test('timeout fails open without losing chat; later preferences still work', async () => {
  const env = environment(); const socket = new env.main.WebSocket(endpoint);
  socket.send(chat()); env.timeout(1500);
  assert.equal(socket.sent[0], chat());
  env.addContent(); await settle(); socket.send(chat());
  assert.equal(JSON.parse(socket.sent[1].slice(0, -1)).arguments[0].tone, 'Gpt_6_Sol_Reasoning');
});
test('closing during initialization drops pending sends and preserves native errors', () => {
  const env = environment(); const socket = new env.main.WebSocket(endpoint);
  socket.readyState = 0; assert.throws(() => socket.send(chat()), /InvalidStateError/);
  socket.readyState = 1; socket.send(chat()); socket.close(); env.timeout(1500);
  assert.equal(socket.sent.length, 0);
  assert.throws(() => env.main.WebSocket.prototype.send.call({}, chat()));
});
test('unrelated sockets pass through byte-for-byte; diagnostics contain no private payloads', async () => {
  const env = environment(); env.addContent(); await settle();
  const other = new env.main.WebSocket('wss://other.example/socket'); other.send(chat());
  assert.equal(other.sent[0], chat());
  const socket = new env.main.WebSocket(endpoint); socket.send(chat());
  socket.receive(JSON.stringify({ type: 2, item: { messages: [{ author: 'bot', contentOrigin: 'DeepLeo', text: 'PRIVATE REPLY' }], result: { value: 'Success', message: 'PRIVATE RESULT' } } }) + '\x1e');
  await settle(); assert.ok(!JSON.stringify(env.messages).includes('PRIVATE'));
});
test('a chat selection propagates to two tabs and survives a new context without a page UI', async () => {
  const store = createStore({ inPageEnabled: true }); // Legacy preference is ignored.
  const a = environment(store), b = environment(store);
  a.addContent(); b.addContent(); await flush();
  await store.chrome.storage.local.set({ conversationTones: { [keyOf(A)]: { tone: 'Gpt_6_Reasoning', at: 1 } } }); await flush();
  const restarted = environment(store); restarted.addContent(); await flush();
  for (const env of [a, b, restarted]) {
    const socket = new env.main.WebSocket(at(A)); socket.send(chat()); await flush();
    assert.equal(toneSent(socket), 'Gpt_6_Reasoning');
    // Other chats stay on the default.
    const other = new env.main.WebSocket(endpoint); other.send(chat());
    assert.equal(toneSent(other), 'Gpt_6_Sol_Reasoning');
  }
});
test('failed initial storage read passes requests through without falsely reporting an override', async () => {
  const store = createStore(); store.chrome.storage.local.get = async () => { throw new Error('storage unavailable'); };
  const env = environment(store); env.addContent(); await settle();
  const socket = new env.main.WebSocket(endpoint); socket.send(chat()); await settle();
  assert.equal(socket.sent[0], chat());
  assert.ok(!env.messages.some(message => message.kind === 'applied'));
  assert.match(env.contents[0].snapshot().observed.warning, /Preferences unavailable/);
});
test('final response updates usage, honors zero, retains missing data and ignores older sockets', async () => {
  const env = environment(); env.addContent(); await settle();
  const first = new env.main.WebSocket(endpoint); first.send(chat()); await settle();
  const frame = allowance => JSON.stringify({ type: 2, item: { result: { value: 'Success' }, throttling: { metering: { GPT61SolQueryDaily: { remainingAllowance: allowance } } } } }) + '\x1e';
  first.receive(frame(0)); await settle(); assert.equal(env.contents[0].snapshot().observed.metering.GPT61SolQueryDaily, 0);
  const second = new env.main.WebSocket(endpoint); second.send(chat()); await settle();
  second.receive(frame(5)); first.receive(frame(30)); await settle();
  assert.equal(env.contents[0].snapshot().observed.metering.GPT61SolQueryDaily, 5);
  second.receive('{"type":2,"item":{"result":{"value":"Success"}}}\x1e'); await settle();
  assert.deepEqual(env.contents[0].snapshot().observed.metering, { GPT61SolQueryDaily: 5 });
});
test('cross-tab changes arriving during startup are not overwritten by a stale initial read', async () => {
  const store = createStore(); let resolve;
  store.chrome.storage.local.get = () => new Promise(r => { resolve = r; });
  const env = environment(store); env.addContent();
  await store.chrome.storage.local.set({ conversationTones: { [keyOf(A)]: { tone: 'Gpt_5_5_Reasoning', at: 2 } } });
  resolve({ conversationTones: { [keyOf(A)]: { tone: 'Magic', at: 1 } } }); await flush();
  const socket = new env.main.WebSocket(at(A)); socket.send(chat()); await flush();
  assert.equal(toneSent(socket), 'Gpt_5_5_Reasoning');
});
test('conversation capacity and expiry flow through the hook without persisting per-chat data', async () => {
  const env = environment(); env.addContent(); await settle();
  const socket = new env.main.WebSocket(endpoint); socket.send(chat()); await settle();
  socket.receive(JSON.stringify({ type: 2, item: {
    conversationExpiryTime: '2026-11-18T22:30:00+10:00',
    throttling: { numUserMessagesInConversation: 540, maxNumUserMessagesInConversation: 600 }
  } }) + '\x1e'); await settle();
  assert.deepEqual(env.contents[0].snapshot().observed.conversation, { used: 540, limit: 600 });
  assert.equal(env.contents[0].snapshot().observed.expiresAt, '2026-11-18T12:30:00.000Z');
  assert.deepEqual(env.store.data, {});
  socket.receive('{"type":1,"target":"update","arguments":[{"messages":[]}]}\x1e'); await settle();
  assert.equal(env.contents[0].snapshot().observed.conversation.used, 540);
  socket.receive('{"type":2,"item":{"result":{"value":"Success"}}}\x1e'); await settle();
  assert.equal(env.contents[0].snapshot().observed.conversation.used, 540);
  assert.equal(env.contents[0].snapshot().observed.expiresAt, '2026-11-18T12:30:00.000Z');
});
test('a chat on a new socket clears prior conversation data and ignores older expiry updates', async () => {
  const env = environment(); env.addContent(); await settle();
  const first = new env.main.WebSocket(endpoint); first.send(chat()); await settle();
  const frame = JSON.stringify({ type: 2, item: { conversationExpiryTime: '2026-11-18T12:00:00Z', throttling: { numUserMessagesInConversation: 600, maxNumUserMessagesInConversation: 600 } } }) + '\x1e';
  first.receive(frame); await settle();
  const second = new env.main.WebSocket(endpoint); second.send(chat()); await settle();
  first.receive(frame); await settle();
  assert.equal(env.contents[0].snapshot().observed.conversation, undefined);
  assert.equal(env.contents[0].snapshot().observed.expiresAt, undefined);
});

test('allowances are saved for every tab and updated by each new reading', async () => {
  const store = createStore();
  const first = environment(store); first.addContent(); await settle();
  const socket = new first.main.WebSocket(endpoint); socket.send(chat()); await settle();
  const metering = (daily, weekly) => JSON.stringify({ type: 2, item: { throttling: { metering: {
    GPT61SolQueryDaily: { remainingAllowance: daily }, ...(weekly === undefined ? {} : { GPT61SolQueryWeekly: { remainingAllowance: weekly } }) } } } }) + '\x1e';
  socket.receive(metering(30, 70)); await settle();
  assert.deepEqual(clone(store.data.usage.metering), { GPT61SolQueryDaily: 30, GPT61SolQueryWeekly: 70 });
  // A tab opened later (or after a restart) shows the last reported allowance at once.
  const second = environment(store); const content = second.addContent(); await settle();
  assert.deepEqual(content.snapshot().observed.metering, { GPT61SolQueryDaily: 30, GPT61SolQueryWeekly: 70 });
  // A partial reading updates only what it reports, in every tab.
  socket.receive(metering(29)); await settle();
  assert.deepEqual(clone(store.data.usage.metering), { GPT61SolQueryDaily: 29, GPT61SolQueryWeekly: 70 });
  assert.equal(content.snapshot().observed.metering.GPT61SolQueryDaily, 29);
});

test('a chat keeps its last readings across tabs and restarts, keyed by a hash', async () => {
  const store = createStore();
  const env = environment(store, { pathname: `/chat/conversation/${A}` }); env.addContent(); await flush();
  const socket = new env.main.WebSocket(at(A)); socket.send(chat()); await flush();
  socket.receive(JSON.stringify({ type: 2, item: { result: { value: 'Success' }, messages: [{ author: 'bot', contentOrigin: 'DeepLeo' }],
    conversationExpiryTime: '2026-11-18T12:00:00Z', throttling: { numUserMessagesInConversation: 4, maxNumUserMessagesInConversation: 600 } } }) + '\x1e');
  await flush(); env.timeout(250);
  const saved = store.data.conversationReadings[keyOf(A)];
  assert.equal(saved.appliedTone, 'Gpt_6_Sol_Reasoning');
  assert.deepEqual(clone(saved.conversation), { used: 4, limit: 600 });
  assert.ok(!JSON.stringify(store.data).includes(A));
  // Opening the same chat in a new tab shows the last reply without a new message.
  const later = environment(store, { pathname: `/chat/conversation/${A}` }); const content = later.addContent(); await flush();
  const observed = content.snapshot().observed;
  assert.equal(observed.appliedTone, 'Gpt_6_Sol_Reasoning');
  assert.equal(observed.result, 'Success'); assert.equal(observed.origin, 'DeepLeo');
  assert.equal(observed.conversation.used, 4); assert.equal(observed.expiresAt, '2026-11-18T12:00:00.000Z');
  // A new, unsent chat shows none of it.
  later.location.pathname = '/chat';
  assert.equal(content.snapshot().observed.conversation, undefined);
});
