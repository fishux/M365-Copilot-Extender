// Simulated WebSocket; never connects to a network or Microsoft.
window.WebSocket = class FixtureSocket extends EventTarget {
  static OPEN = 1;
  constructor(url) { super(); this.url = url; this.readyState = 1; this.lastTone = null; }
  send(data) {
    const record = JSON.parse(data.split('\x1e')[0]);
    if (record.target === 'chat') {
      this.lastTone = record.arguments[0].tone;
      document.getElementById('wire').textContent = 'Outgoing tone: ' + this.lastTone;
    }
  }
  receive(frame) { this.dispatchEvent(new MessageEvent('message', { data: JSON.stringify(frame) + '\x1e' })); }
};
