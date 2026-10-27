const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const { createStore, settle, models } = require('./harness.cjs');

// Runs the actual packaged worker with a recording chrome.action and a canvas stub.
function worker(store = createStore()) {
  const calls = [], listeners = {};
  class FakeContext {
    constructor(size) { this.size = size; this.fills = []; }
    scale() {} beginPath() {} roundRect() {} fill() { this.fills.push(this.fillStyle); } stroke() {} arc() {}
    createLinearGradient() { const stops = []; return { stops, addColorStop: (_, color) => stops.push(color) }; }
    getImageData() { return { size: this.size, fills: this.fills.map(fill => fill.stops || fill) }; }
  }
  const record = name => async details => { calls.push({ name, ...JSON.parse(JSON.stringify(details)) }); };
  const chrome = {
    ...store.chrome,
    runtime: { id: 'test-extension', onMessage: { addListener(fn) { listeners.message = fn; } } },
    tabs: { onUpdated: { addListener(fn) { listeners.updated = fn; } } },
    action: Object.fromEntries(['setBadgeText', 'setBadgeBackgroundColor', 'setBadgeTextColor', 'setTitle', 'setIcon'].map(name => [name, record(name)]))
  };
  const context = vm.createContext({ chrome, console, OffscreenCanvas: class { constructor(w) { this.w = w; } getContext() { return new FakeContext(this.w); } }, Path2D: class {} });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../extension/background.js'), 'utf8'), context);
  const send = (message, tabId = 5) => listeners.message(message, { id: 'test-extension', tab: { id: tabId } });
  const last = (name, tabId) => calls.filter(call => call.name === name && call.tabId === tabId).at(-1);
  return { calls, listeners, send, last, store };
}

test('global badge shows the fixed default model and ignores a legacy saved default', async () => {
  const env = worker(createStore({ selectedTone: 'Gpt_5_6_Reasoning' })); await settle();
  assert.equal(env.last('setBadgeText').text, '6S');
  assert.match(env.last('setTitle').title, /Default: GPT-6 Sol/);
  await env.store.chrome.storage.local.set({ selectedTone: 'Magic' }); await settle();
  assert.equal(env.last('setBadgeText').text, '6S');
});

test('tab badge colour and icon turn orange at 20 and red at 10 or below', async () => {
  const env = worker(); await settle();
  const cases = [[null, '#6254d9', 'path'], [21, '#6254d9', 'path'], [20, '#e07b16', 'warn'], [11, '#e07b16', 'warn'], [10, '#d93025', 'critical'], [0, '#d93025', 'critical']];
  for (const [remaining, color, icon] of cases) {
    env.send({ type: 'M365_EXTENDER_TAB', tone: 'Gpt_61_Sol_Reasoning', remaining }); await settle();
    assert.equal(env.last('setBadgeBackgroundColor', 5).color, color, `remaining ${remaining}`);
    assert.equal(env.last('setBadgeText', 5).text, '6.1');
    const setIcon = env.last('setIcon', 5);
    if (icon === 'path') assert.deepEqual(setIcon.path, { 16: 'icons/icon16.png', 32: 'icons/icon32.png' });
    else assert.deepEqual(setIcon.imageData[16].fills[0], icon === 'warn' ? ['#f7a94f', '#d9690f'] : ['#f2665c', '#c5221f']);
  }
  assert.match(env.last('setTitle', 5).title, /GPT-6\.1 Sol · 0 messages left/);
});

test('tab badge resets when the tab navigates and ignores foreign or malformed messages', async () => {
  const env = worker(); await settle();
  env.send({ type: 'M365_EXTENDER_TAB', tone: 'Gpt_6_Reasoning', remaining: 3 }); await settle();
  env.listeners.updated(5, { status: 'loading' }); await settle();
  assert.equal(env.last('setBadgeText', 5).text, null);
  assert.deepEqual(env.last('setIcon', 5).path, { 16: 'icons/icon16.png', 32: 'icons/icon32.png' });
  const before = env.calls.length;
  env.listeners.message({ type: 'M365_EXTENDER_TAB', tone: 'Magic' }, { id: 'someone-else', tab: { id: 5 } });
  env.listeners.message({ type: 'M365_EXTENDER_TAB', tone: 'Magic' }, { id: 'test-extension' });
  await settle();
  assert.equal(env.calls.length, before);
  env.send({ type: 'M365_EXTENDER_TAB', tone: 'not-a-model', remaining: -4 }); await settle();
  assert.equal(env.last('setBadgeBackgroundColor', 5).color, '#6254d9');
});

test('a Work IQ tab adds W to the badge and its tooltip', async () => {
  const env = worker(); await settle();
  env.send({ type: 'M365_EXTENDER_TAB', tone: 'Gpt_6_Sol_Reasoning', remaining: null, workIQ: true }); await settle();
  assert.equal(env.last('setBadgeText', 5).text, '6SW');
  assert.match(env.last('setTitle', 5).title, /GPT-6 Sol · Work IQ$/);
  // Only a literal true counts; anything else is a normal tab.
  env.send({ type: 'M365_EXTENDER_TAB', tone: 'Gpt_6_Sol_Reasoning', remaining: null, workIQ: 'yes' }); await settle();
  assert.equal(env.last('setBadgeText', 5).text, '6S');
});

test('every model has a badge and the packaged bundle carries models.json in order', () => {
  const bundle = fs.readFileSync(path.join(__dirname, '../extension/background.js'), 'utf8');
  assert.ok(bundle.includes(`shared.models = ${JSON.stringify(models)};`));
  for (const model of models) assert.ok(model.badge.length >= 1 && model.badge.length <= 4);
});
