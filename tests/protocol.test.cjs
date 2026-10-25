const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { load, clone, chat, endpoint } = require('./harness.cjs');
const context = vm.createContext({ URL }); load(context, 'core.js');
const core = context.M365ExtenderCore;

test('all six model tones work for invocation types 1 and 4 and preserve message content', () => {
  for (const model of core.MODELS) for (const type of [1, 4]) {
    const result = core.rewrite(chat('Magic', type), model.tone);
    const frame = JSON.parse(result.data.slice(0, -1));
    assert.equal(frame.arguments[0].tone, model.tone);
    assert.deepEqual(frame.arguments[0].message, { text: 'PRIVATE PROMPT' });
    assert.deepEqual(frame.arguments[0].optionsSets, ['keep-me']);
    assert.equal(result.tones.length, 1);
  }
});
test('preserves handshakes, pings, Metrics, malformed JSON, binary data, trailing separators', () => {
  const other = '{}\x1e{"type":6}\x1e{"type":1,"target":"Metrics","arguments":[{}]}\x1enot-json\x1e';
  const result = core.rewrite(other + chat(), core.DEFAULT_TONE);
  assert.ok(result.data.startsWith(other)); assert.ok(result.data.endsWith('\x1e'));
  assert.equal(core.rewrite(other, core.DEFAULT_TONE).data, other);
  const binary = new Uint8Array([1, 2]); assert.equal(core.rewrite(binary, core.DEFAULT_TONE).data, binary);
  assert.equal(core.rewrite(chat(), 'bogus').data, chat());
  // Chat records are counted even when no valid tone is configured (fail-open).
  assert.equal(result.chats, 1); assert.equal(core.rewrite(other, core.DEFAULT_TONE).chats, 0);
  assert.equal(core.rewrite(chat(), null).chats, 1); assert.equal(core.rewrite(chat(), null).tones.length, 0);
  assert.equal(core.rewrite(binary, core.DEFAULT_TONE).chats, 0);
});
test('rejects malformed chat shapes and leaves an already selected frame byte-identical', () => {
  for (const args of [null, {}, [], [null], [[]]]) {
    const data = JSON.stringify({ type: 4, target: 'chat', arguments: args });
    assert.equal(core.rewrite(data, core.DEFAULT_TONE).data, data);
  }
  assert.equal(core.rewrite(chat(core.DEFAULT_TONE), core.DEFAULT_TONE).data, chat(core.DEFAULT_TONE));
});
test('endpoint scoping excludes unrelated sockets and strips credentials', () => {
  assert.deepEqual(clone(core.connectionInfo(endpoint)), { scenario: 'OfficeWebPaidCopilot', conversationId: null });
  const id = '3F2504E0-4F89-11D3-9A0C-0305E82C3301';
  assert.deepEqual(clone(core.connectionInfo(endpoint + '&ConversationId=' + id)), { scenario: 'OfficeWebPaidCopilot', conversationId: id.toLowerCase() });
  assert.equal(core.connectionInfo(endpoint + '&conversationid=' + id).conversationId, id.toLowerCase());
  for (const bad of ['', 'short', '../../etc', id + '<script>']) assert.equal(core.connectionInfo(endpoint + '&ConversationId=' + encodeURIComponent(bad)).conversationId, null);
  for (const url of ['wss://evil.example/m365Copilot/Chathub/a', 'wss://substrate.office.com.evil.example/m365Copilot/Chathub/a', 'wss://substrate.office.com/telemetry', 'invalid']) assert.equal(core.connectionInfo(url), null);
});
test('reads zero allowances and rejects missing, negative, non-finite and string values', () => {
  const metering = core.parseMetering({ GPT61SolQueryDaily: { remainingAllowance: 0 }, GPT61SolQueryWeekly: { remainingAllowance: 12 }, Invalid: { remainingAllowance: -1 }, NaN: { remainingAllowance: NaN }, String: { remainingAllowance: '40' } });
  assert.deepEqual(clone(metering), { GPT61SolQueryDaily: 0, GPT61SolQueryWeekly: 12 });
  assert.deepEqual(clone(core.allowanceFor(core.DEFAULT_TONE, metering)), { daily: 0, weekly: 12 });
  assert.equal(core.allowanceFor('Gpt_6_Sol_Reasoning', metering), null);
  assert.equal(core.allowanceFor('Magic', metering), null);
  assert.equal(core.allowanceFor(core.DEFAULT_TONE, {}), null);
});
test('only associates unambiguous model-specific Daily/Weekly keys when observed', () => {
  assert.deepEqual(clone(core.allowanceFor('Gpt_6_Sol_Reasoning', { GPT6SolQueryDaily: 9 })), { daily: 9, weekly: null });
  assert.equal(core.allowanceFor('Gpt_6_Reasoning', { GPT61SolQueryDaily: 9 }), null);
  assert.equal(core.allowanceFor(core.DEFAULT_TONE, { ClaudeOpusQueryDaily: 40 }), null);
});
test('reads quota and refusals from final type 2 and streaming type 1, without leaking text', () => {
  const item = { throttling: { metering: { GPT61SolQueryWeekly: { remainingAllowance: 0 } }, numUserMessagesInConversation: 3, maxNumUserMessagesInConversation: 600 }, result: { value: 'OutOfCredits', message: 'PRIVATE REFUSAL' }, messages: [{ author: 'bot', contentOrigin: 'BotConnection', text: 'PRIVATE REPLY' }] };
  const output = clone(core.readMetadata({ type: 2, item }));
  assert.deepEqual(output, [{ metering: { GPT61SolQueryWeekly: 0 }, conversation: { used: 3, limit: 600 }, result: 'OutOfCredits', origin: 'BotConnection', final: true }]);
  assert.ok(!JSON.stringify(output).includes('PRIVATE'));
  assert.equal(core.readMetadata({ type: 1, target: 'update', arguments: [item] }).length, 1);
});
test('missing metadata remains absent in both final and streaming records', () => {
  assert.deepEqual(clone(core.readMetadata({ type: 2, item: { result: { value: 'Success' } } })), [{ result: 'Success', final: true }]);
  assert.deepEqual(clone(core.readMetadata({ type: 1, target: 'update', arguments: [{ throttling: {} }] })), []);
});
test('capacity meter uses server limits, handles zero, and clamps over-limit readings', () => {
  const empty = core.conversationView({ used: 0, limit: 800 });
  assert.equal(empty.capacity, '0 / 800 messages'); assert.equal(empty.percent, 0);
  assert.equal(empty.nearLimit, false);
  const almost = core.conversationView({ used: 720, limit: 800 });
  assert.equal(almost.nearLimit, true); assert.equal(almost.percent, 90);
  const full = core.conversationView({ used: 900, limit: 800 });
  assert.equal(full.percent, 100); assert.match(full.capacityTitle, /^0 messages remaining/);
  const zeroLimit = core.conversationView({ used: 0, limit: 0 });
  assert.equal(zeroLimit.percent, null); assert.equal(zeroLimit.capacity, '0 / 0 messages');
  for (const value of [undefined, { used: -1, limit: 600 }, { used: 4, limit: '600' }]) {
    assert.equal(core.conversationView(value).capacity, 'Awaiting reply');
    assert.equal(core.conversationView(value).percent, null);
  }
});
test('expiry parsing requires an explicit timezone and rejects malformed calendar dates', () => {
  assert.equal(core.parseExpiry('2026-11-18T12:30:00.1234567Z'), '2026-11-18T12:30:00.123Z');
  assert.equal(core.parseExpiry('2026-11-18T22:30:00+10:00'), '2026-11-18T12:30:00.000Z');
  for (const value of [null, 123456, '', 'tomorrow', '2026-11-18', '2026-11-18T12:30:00', '2026-02-30T00:00:00Z', '<script>alert(1)</script>']) assert.equal(core.parseExpiry(value), null);
});
test('expiry labels handle short intervals, day ranges, passed timestamps and absence', () => {
  const now = Date.parse('2026-10-19T00:00:00Z');
  assert.match(core.conversationView(null, '2026-10-19T00:10:00Z', now).expiryText, /<1h$/);
  assert.match(core.conversationView(null, '2026-10-20T00:00:00Z', now).expiryText, /24h$/);
  assert.match(core.conversationView(null, '2026-11-18T00:00:00Z', now).expiryText, /30d$/);
  assert.match(core.conversationView(null, '2026-10-18T00:00:00Z', now).expiryText, /Passed$/);
  assert.equal(core.conversationView(null, null, now).expiryText, 'Not reported');
  assert.match(core.conversationView(null, '2026-11-18T00:00:00Z', now).expiryTitle, /not a guaranteed data deletion date/);
});
test('final and streaming metadata carry only normalized server expiry timestamps', () => {
  const item = { conversationExpiryTime: '2026-11-18T22:30:00+10:00' };
  assert.equal(core.readMetadata({ type: 2, item })[0].expiresAt, '2026-11-18T12:30:00.000Z');
  assert.equal(core.readMetadata({ type: 1, target: 'update', arguments: [item] })[0].expiresAt, '2026-11-18T12:30:00.000Z');
  assert.deepEqual(clone(core.readMetadata({ type: 2, item: { conversationExpiryTime: 'invalid' } })), []);
});
