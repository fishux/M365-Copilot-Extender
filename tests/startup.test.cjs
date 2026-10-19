const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { environment, createStore, settle, endpoint, chat } = require('./harness.cjs');

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
  assert.equal(JSON.parse(socket.sent[0].slice(0, -1)).arguments[0].tone, 'Gpt_5_5_Reasoning');
});
