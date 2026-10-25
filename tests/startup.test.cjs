const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { environment, createStore, settle, endpoint, chat, clone } = require('./harness.cjs');

test('packaged entry points have no separate shared-core script dependency', () => {
  const root = path.join(__dirname, '../extension');
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'manifest.json'), 'utf8'));
  assert.deepEqual(manifest.content_scripts.map(group => group.js), [['page-hook.js'], ['content.js']]);
  const popup = fs.readFileSync(path.join(root, 'popup.html'), 'utf8');
  assert.deepEqual([...popup.matchAll(/<script src="([^"]+)"/g)].map(match => match[1]), ['popup.js']);
  assert.equal(fs.existsSync(path.join(root, 'core.js')), false);
});
test('actual runtime bundles start on /chat with no global core and apply saved tone', async () => {
  const env = environment(createStore({ selectedTone: 'Gpt_6_Sol_Reasoning' }), { bundled: true });
  assert.equal(env.main.M365ExtenderCore, undefined);
  const content = env.addContent(); await settle();
  assert.equal(content.M365ExtenderCore, undefined);
  assert.equal(content.snapshot().hookReady, true);
  const socket = new env.main.WebSocket(endpoint); socket.send(chat()); await settle();
  assert.equal(JSON.parse(socket.sent[0].slice(0, -1)).arguments[0].tone, 'Gpt_6_Sol_Reasoning');
  assert.equal(content.snapshot().observed.appliedTone, 'Gpt_6_Sol_Reasoning');
  socket.receive(JSON.stringify({ type: 2, item: { result: { value: 'Success' },
    conversationExpiryTime: '2026-11-18T12:00:00Z',
    throttling: { numUserMessagesInConversation: 2, maxNumUserMessagesInConversation: 600,
      metering: { GPT61SolQueryDaily: { remainingAllowance: 25 }, GPT61SolQueryWeekly: { remainingAllowance: 48 } } }
  } }) + '\x1e'); await settle();
  const observed = content.snapshot().observed;
  assert.equal(observed.metering.GPT61SolQueryWeekly, 48);
  assert.equal(observed.conversation.used, 2);
  assert.equal(observed.expiresAt, '2026-11-18T12:00:00.000Z');
});
test('actual bundle handles an early chat before saved settings finish loading', async () => {
  const store = createStore(); let resolve;
  store.chrome.storage.local.get = () => new Promise(done => { resolve = done; });
  const env = environment(store, { bundled: true }); env.addContent();
  const socket = new env.main.WebSocket(endpoint); socket.send(chat());
  assert.equal(socket.sent.length, 0);
  resolve({ selectedTone: 'Gpt_5_5_Reasoning' }); await settle();
  // A legacy saved default is ignored: the default model is fixed.
  assert.equal(JSON.parse(socket.sent[0].slice(0, -1)).arguments[0].tone, 'Gpt_6_Sol_Reasoning');
});


test('packaged snapshot persists through sparse final frames, idle sockets and polling until valid replacement data', async () => {
  const env = environment(createStore(), { bundled: true });
  const content = env.addContent(); await settle();
  const socket = new env.main.WebSocket(endpoint); socket.send(chat()); await settle();
  const receive = item => socket.receive(JSON.stringify({ type: 2, item }) + '\x1e');
  receive({ conversationExpiryTime: '2026-11-18T12:00:00Z', result: { value: 'Success' },
    messages: [{ author: 'bot', contentOrigin: 'DeepLeo' }],
    throttling: { numUserMessagesInConversation: 2, maxNumUserMessagesInConversation: 600,
      metering: { GPT61SolQueryDaily: { remainingAllowance: 25 }, GPT61SolQueryWeekly: { remainingAllowance: 48 } } } });
  await settle(); const original = content.snapshot().observed;
  receive({}); receive({ result: { value: 'Success' }, conversationExpiryTime: 'invalid', throttling: { metering: {} } });
  const idle = new env.main.WebSocket(endpoint);
  idle.receive(JSON.stringify({ type: 2, item: { result: { value: 'Throttled' } } }) + '\x1e');
  await settle();
  assert.deepEqual(content.snapshot().observed, original);
  // A follow-up request preserves the last reply while waiting for new readings.
  socket.send(chat()); await settle(); assert.deepEqual(content.snapshot().observed, original);
  receive({ conversationExpiryTime: '2026-11-19T12:00:00Z',
    throttling: { numUserMessagesInConversation: 3, maxNumUserMessagesInConversation: 600,
      metering: { GPT61SolQueryDaily: { remainingAllowance: 0 } } } });
  await settle(); const updated = content.snapshot().observed;
  assert.equal(updated.conversation.used, 3);
  assert.equal(updated.expiresAt, '2026-11-19T12:00:00.000Z');
  assert.deepEqual(updated.metering, { GPT61SolQueryDaily: 0, GPT61SolQueryWeekly: 48 });
  assert.equal(updated.origin, 'DeepLeo'); assert.equal(updated.result, 'Success');
  idle.send(chat()); await settle();
  assert.equal(content.snapshot().observed.conversation, undefined);
  // Allowances are account-wide, so they stay visible in the next conversation.
  assert.deepEqual(content.snapshot().observed.metering, { GPT61SolQueryDaily: 0, GPT61SolQueryWeekly: 48 });
  receive({ result: { value: 'ForbiddenRequest' } }); await settle();
  assert.equal(content.snapshot().observed.result, undefined);
  // Socket creation order does not prevent reusing an earlier connection.
  socket.send(chat()); receive({ result: { value: 'Success' } }); await settle();
  assert.equal(content.snapshot().observed.result, 'Success');
  // Only allowances are saved: these connections carry no conversation ID.
  assert.deepEqual(Object.keys(env.store.data), ['usage']);
  assert.deepEqual(clone(env.store.data.usage.metering), { GPT61SolQueryDaily: 0, GPT61SolQueryWeekly: 48 });
});
