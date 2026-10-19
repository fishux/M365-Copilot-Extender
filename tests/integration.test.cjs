const { test } = require('node:test');
const assert = require('node:assert/strict');
const { settle, chat, endpoint, environment, createStore } = require('./harness.cjs');

test('saved tone applies before first chat; native identity, constants and subclassing survive', async () => {
  const env = environment(createStore({ selectedTone: 'Gpt_5_6_Reasoning' }));
  env.addContent(); await settle();
  const socket = new env.main.WebSocket(endpoint); socket.send(chat());
  await settle();
  assert.equal(JSON.parse(socket.sent[0].slice(0, -1)).arguments[0].tone, 'Gpt_5_6_Reasoning');
  assert.ok(socket instanceof env.main.WebSocket); assert.ok(socket instanceof env.FakeSocket);
  assert.equal(env.main.WebSocket.OPEN, 1);
  class SubSocket extends env.main.WebSocket {} assert.ok(new SubSocket(endpoint) instanceof SubSocket);
  assert.equal(env.panels[0].state.observed.appliedTone, 'Gpt_5_6_Reasoning');
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
  assert.equal(JSON.parse(socket.sent[1].slice(0, -1)).arguments[0].tone, 'Gpt_61_Sol_Reasoning');
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
test('selection and visibility changes propagate to two tabs and survive a new context', async () => {
  const store = createStore(); const a = environment(store), b = environment(store);
  a.addContent(); b.addContent(); await settle();
  await a.panels[0].callbacks.onSelect('Gpt_6_Reasoning'); await settle();
  assert.equal(b.panels[0].state.tone, 'Gpt_6_Reasoning');
  await store.chrome.storage.local.set({ inPageEnabled: false }); await settle();
  assert.equal(a.hosts[0].style.display, 'none'); assert.equal(b.hosts[0].style.display, 'none');
  const restarted = environment(store); restarted.addContent(); await settle();
  assert.equal(restarted.panels[0].state.tone, 'Gpt_6_Reasoning'); assert.equal(restarted.hosts[0].style.display, 'none');
  const socket = new b.main.WebSocket(endpoint); socket.send(chat());
  assert.equal(JSON.parse(socket.sent[0].slice(0, -1)).arguments[0].tone, 'Gpt_6_Reasoning');
});
test('failed saves retain the previous active preference', async () => {
  const env = environment(); env.addContent(); await settle();
  env.store.chrome.storage.local.set = async () => { throw new Error('disk full'); };
  await env.panels[0].callbacks.onSelect('Gpt_5_5_Reasoning'); await settle();
  assert.equal(env.panels[0].state.tone, 'Gpt_61_Sol_Reasoning');
  assert.match(env.panels[0].state.error, /Could not save/);
});
test('failed initial storage read passes requests through without falsely reporting an override', async () => {
  const store = createStore(); store.chrome.storage.local.get = async () => { throw new Error('storage unavailable'); };
  const env = environment(store); env.addContent(); await settle();
  const socket = new env.main.WebSocket(endpoint); socket.send(chat()); await settle();
  assert.equal(socket.sent[0], chat());
  assert.ok(!env.messages.some(message => message.kind === 'applied'));
  assert.match(env.panels[0].state.error, /could not be loaded/);
});
test('final response updates usage, honors zero, clears missing data and ignores older sockets', async () => {
  const env = environment(); env.addContent(); await settle();
  const first = new env.main.WebSocket(endpoint); first.send(chat()); await settle();
  const frame = allowance => JSON.stringify({ type: 2, item: { result: { value: 'Success' }, throttling: { metering: { GPT61SolQueryDaily: { remainingAllowance: allowance } } } } }) + '\x1e';
  first.receive(frame(0)); await settle(); assert.equal(env.panels[0].state.observed.metering.GPT61SolQueryDaily, 0);
  const second = new env.main.WebSocket(endpoint); second.send(chat()); await settle();
  second.receive(frame(5)); first.receive(frame(30)); await settle();
  assert.equal(env.panels[0].state.observed.metering.GPT61SolQueryDaily, 5);
  second.receive('{"type":2,"item":{"result":{"value":"Success"}}}\x1e'); await settle();
  assert.deepEqual(env.panels[0].state.observed.metering, {});
});
test('cross-tab changes arriving during startup are not overwritten by a stale initial read', async () => {
  const store = createStore(); let resolve;
  store.chrome.storage.local.get = () => new Promise(r => { resolve = r; });
  const env = environment(store); env.addContent();
  await store.chrome.storage.local.set({ selectedTone: 'Gpt_5_5_Reasoning', inPageEnabled: false });
  resolve({ selectedTone: 'Magic', inPageEnabled: true }); await settle();
  assert.equal(env.panels[0].state.tone, 'Gpt_5_5_Reasoning'); assert.equal(env.hosts[0].style.display, 'none');
});
