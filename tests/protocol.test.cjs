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
});
test('rejects malformed chat shapes and leaves an already selected frame byte-identical', () => {
  for (const args of [null, {}, [], [null], [[]]]) {
    const data = JSON.stringify({ type: 4, target: 'chat', arguments: args });
    assert.equal(core.rewrite(data, core.DEFAULT_TONE).data, data);
  }
  assert.equal(core.rewrite(chat(core.DEFAULT_TONE), core.DEFAULT_TONE).data, chat(core.DEFAULT_TONE));
});
test('endpoint scoping excludes unrelated sockets and strips credentials', () => {
  assert.deepEqual(clone(core.connectionInfo(endpoint)), { scenario: 'OfficeWebPaidCopilot' });
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
test('final missing quota clears old readings; streaming-only absence does not', () => {
  assert.deepEqual(clone(core.readMetadata({ type: 2, item: { result: { value: 'Success' } } })), [{ metering: {}, result: 'Success', final: true }]);
  assert.deepEqual(clone(core.readMetadata({ type: 1, target: 'update', arguments: [{ throttling: {} }] })), []);
});
