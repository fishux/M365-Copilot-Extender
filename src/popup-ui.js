/* Toolbar popup UI. Never injected into Copilot. */
(() => {
  "use strict";
  const core = globalThis.M365ExtenderCore;
  const STALE_MS = 5 * 60 * 1000;
  const css = `
    :host { all: initial; color-scheme: light dark; --bg:#fff; --ink:#1c2132;
      --muted:#697185; --line:#e7e9f0; --soft:#f6f7fb; --accent:#6254d9;
      --good:#248568; --warn:#b15722; font:13px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif; color:var(--ink); }
    * { box-sizing:border-box; } [hidden] { display:none !important; }
    .card { width:308px; background:var(--bg); overflow:hidden; }
    .header { display:flex; align-items:center; gap:8px; padding:14px 14px 10px; }
    .logo { display:grid; place-items:center; width:28px; height:28px; flex:none; border-radius:9px;
      color:white; background:linear-gradient(145deg,#8c76ed,#5145bb); box-shadow:0 3px 9px #6254d926; }
    .logo svg { width:18px; height:18px; }
    .brand { flex:1; } h2 { margin:0; font-size:14px; font-weight:650; letter-spacing:-.2px; }
    .eyebrow { font-size:10px; letter-spacing:1.1px; text-transform:uppercase; color:var(--muted); }
    .body { padding:0 14px 13px; }
    label.field-label { display:block; margin:0 0 5px; font-size:11px; font-weight:600; color:var(--muted); }
    .select-wrap { position:relative; } select { display:block; width:100%; min-height:39px;
      appearance:none; padding:8px 32px 8px 11px; border:1px solid var(--line); border-radius:11px;
      background:var(--soft); color:var(--ink); font:inherit; font-size:13px; font-weight:600; cursor:pointer; }
    select:disabled { cursor:wait; opacity:.6; } .chevron { position:absolute; right:12px; top:10px; pointer-events:none; color:var(--muted); }
    select:focus-visible { outline:3px solid var(--accent); outline-offset:3px; }
    .hint { display:flex; align-items:center; gap:6px; margin:6px 0 0; font-size:11px; color:var(--muted); }
    .hint svg { width:12px; height:12px; flex:none; color:var(--accent); }
    .usage { margin-top:12px; padding:11px; border:1px solid var(--line); border-radius:12px; }
    .usage-head { display:flex; align-items:center; justify-content:space-between; gap:8px; margin-bottom:7px; }
    .usage-head strong { font-size:11px; font-weight:600; } .live { font-size:10px; letter-spacing:.3px; text-transform:uppercase; color:var(--good); }
    .counts { display:grid; grid-template-columns:1fr 1fr; gap:10px; }
    .counts > div + div { border-left:1px solid var(--line); padding-left:11px; }
    .number { display:block; font-size:24px; line-height:1.2; font-weight:600; letter-spacing:-1px; font-variant-numeric:tabular-nums; }
    .number.zero { color:var(--warn); } .period { font-size:10px; color:var(--muted); }
    .usage-note { margin:7px 0 0; font-size:10px; color:var(--muted); } .pending { margin:0; font-size:11px; color:var(--muted); }
    .status { display:flex; align-items:flex-start; gap:7px; margin:10px 0 0; font-size:11px; color:var(--muted); }
    .dot { width:6px; height:6px; margin-top:5px; border-radius:50%; background:#a7aec0; flex:none; }
    .dot.good { background:var(--good); } .dot.warn { background:var(--warn); }
    .connection-details { margin-top:9px; padding-top:8px; border-top:1px solid var(--line); }
    .connection-title { margin:0; color:var(--muted); font-size:10px; font-weight:500; }
    dl { margin:7px 0 0; display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr); gap:5px 10px; font-size:10px; }
    dt { color:var(--muted); font-size:10px; } dd { margin:0; overflow-wrap:anywhere; }
    .disclaimer { margin:8px 0 0; font-size:10px; color:var(--muted); }
    .error { color:var(--warn); font-size:11px; margin:10px 0 0; }
    .conversation-info { margin-top:10px; padding-top:9px; border-top:1px solid var(--line); font-size:10px; }
    .conversation-row { display:flex; justify-content:space-between; align-items:center; gap:8px; color:var(--muted); }
    .capacity-count { font-weight:500; font-variant-numeric:tabular-nums; color:var(--ink); }
    .capacity-count.near-limit { color:var(--warn); }
    .capacity-track { height:3px; margin:5px 0 6px; border-radius:4px; background:var(--line); overflow:hidden; }
    .capacity-fill { height:100%; border-radius:4px; background:var(--accent); }
    .capacity-fill.near-limit { background:var(--warn); }
    .expiry-row { margin-top:5px; } .expiry-value { color:var(--ink); font-variant-numeric:tabular-nums; }
    .expiry-value:focus-visible { outline:2px solid var(--accent); outline-offset:2px; }
    @media(prefers-color-scheme:dark) { :host { --bg:#1c1d25; --ink:#f0f0f7; --muted:#a2a6b8; --line:#353641; --soft:#252631; --accent:#b5a6ff; --good:#76d3af; --warn:#f2b07d; } }
    @media(prefers-reduced-motion:reduce) { * { transition:none !important; } }
  `;
  const logo = '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 3v18M3 12h18M5.6 5.6l12.8 12.8M5.6 18.4L18.4 5.6" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/><circle cx="12" cy="12" r="4" fill="currentColor"/></svg>';
  const lock = '<svg viewBox="0 0 16 16" fill="none" aria-hidden="true"><rect x="3.5" y="7" width="9" height="7" rx="2" stroke="currentColor"/><path d="M5.5 7V4.5a2.5 2.5 0 015 0V7" stroke="currentColor"/></svg>';

  function mount(host, { onSelect }) {
    const root = host.attachShadow({ mode: "open" });
    const style = document.createElement("style"); style.textContent = css; root.append(style);
    const card = document.createElement("section");
    card.className = "card";
    card.setAttribute("aria-label", "M365 Copilot Extender");
    // This template is static; server and preference values are set with textContent.
    card.innerHTML = `<header class="header"><span class="logo">${logo}</span><div class="brand"><div class="eyebrow">M365 Copilot</div><h2>Extender</h2></div></header>
      <div class="body"><label class="field-label" for="model">Preferred model</label>
      <div class="select-wrap"><select id="model" disabled></select><span class="chevron" aria-hidden="true">⌄</span></div>
      <p class="hint">${lock}<span class="save-label">Loading saved preference…</span></p>
      <section class="usage" aria-label="Model usage allowance" hidden><div class="usage-head"><strong>Messages remaining</strong><span class="live">Server reported</span></div>
      <div class="counts"><div><span class="number daily">—</span><span class="period">Daily allowance</span></div><div><span class="number weekly">—</span><span class="period">Weekly allowance</span></div></div>
      <p class="pending" hidden>Send a message to check your allowance.</p><p class="usage-note"></p></section>
      <section class="conversation-info" aria-label="Conversation capacity and expiry" hidden>
      <div class="conversation-row"><span>Conversation</span><span class="capacity-count">Awaiting reply</span></div>
      <div class="capacity-track" role="meter" aria-label="Conversation messages used" aria-valuemin="0" hidden><div class="capacity-fill"></div></div>
      <div class="conversation-row expiry-row"><span>Server expiry</span><time class="expiry-value" tabindex="0">Not reported</time></div></section>
      <p class="status" role="status" aria-live="polite"><span class="dot"></span><span class="status-text">Waiting for Copilot</span></p>
      <p class="error" role="alert" hidden></p>
      <section class="connection-details" aria-label="Connection details"><h3 class="connection-title">Connection details</h3><dl><div><dt>Route</dt><dd class="route">Not observed</dd></div><div><dt>Last request</dt><dd class="requested">Not observed</dd></div><div><dt>Response</dt><dd class="response">Not observed</dd></div><div><dt>Origin</dt><dd class="origin">Not observed</dd></div></dl>
      <p class="disclaimer">Requests your selected model. Microsoft controls final model routing.</p></section></div>`;
    root.append(card);
    const find = selector => card.querySelector(selector);
    const select = find("select");
    for (const model of core.MODELS) {
      const option = document.createElement("option"); option.value = model.tone;
      option.textContent = model.label + (model.subtitle === "Think deeper" ? " · Think deeper" : "");
      select.append(option);
    }
    select.addEventListener("change", () => onSelect(select.value));

    function update(state) {
      select.value = state.tone;
      select.disabled = !state.loaded || state.saving;
      find(".save-label").textContent = state.saving ? "Saving…" : state.loaded ? "Saved across tabs & restarts" : "Loading saved preference…";
      find(".error").hidden = !state.error;
      find(".error").textContent = state.error || "";
      const observed = state.observed || {};
      const quota = core.allowanceFor(state.tone, observed.metering);
      const model = core.modelFor(state.tone);
      const knownPaidCap = model?.capped && observed.scenario !== "OfficeWebIncludedCopilot";
      const showUsage = Boolean(quota || knownPaidCap);
      find(".usage").hidden = !showUsage;
      find(".counts").hidden = !quota;
      find(".live").hidden = !quota;
      find(".pending").hidden = Boolean(quota);
      find(".pending").textContent = observed.quotaAt ? "Copilot did not report this allowance." : "Send a message to check your allowance.";
      for (const period of ["daily", "weekly"]) {
        const value = quota?.[period];
        find(`.${period}`).textContent = value === null || value === undefined ? "—" : String(value);
        find(`.${period}`).classList.toggle("zero", value === 0);
      }
      const age = observed.quotaAt ? Date.now() - observed.quotaAt : 0;
      find(".usage-note").textContent = quota ? `${age > STALE_MS ? "Last known · " : "Updated "}${new Date(observed.quotaAt).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })} · Refreshes after a reply` : "Allowances appear when Copilot reports them.";
      let text = state.connected === false ? "Open Copilot to see connection and usage details." : "Ready for your next message";
      let status = "";
      if (state.connected !== false && !state.hookReady) { text = "Connecting to Copilot…"; }
      if (observed.appliedTone === state.tone) { text = "Preference applied to the last request"; status = "good"; }
      if (observed.result === "OutOfCredits") { text = "The last request reached a usage limit."; status = "warn"; }
      else if (observed.result && observed.result !== "Success") { text = `Copilot response: ${observed.result}`; status = "warn"; }
      if (observed.warning) { text = observed.warning; status = "warn"; }
      find(".status-text").textContent = text;
      find(".dot").className = `dot ${status}`;
      find(".route").textContent = observed.scenario === "OfficeWebPaidCopilot" ? "Paid Copilot" : observed.scenario === "OfficeWebIncludedCopilot" ? "Included Copilot" : "Not observed";
      find(".requested").textContent = core.modelFor(observed.appliedTone)?.label || "Not observed";
      find(".response").textContent = observed.result || "Not observed";
      find(".origin").textContent = observed.origin || "Not observed";
      const conversation = core.conversationView(observed.conversation, observed.expiresAt);
      find(".conversation-info").hidden = state.connected === false;
      find(".capacity-count").textContent = conversation.capacity;
      find(".capacity-count").title = conversation.capacityTitle;
      find(".capacity-count").classList.toggle("near-limit", conversation.nearLimit);
      const meter = find(".capacity-track");
      meter.hidden = conversation.percent === null;
      if (conversation.percent !== null) {
        meter.setAttribute("aria-valuemax", String(conversation.limit));
        meter.setAttribute("aria-valuenow", String(Math.min(conversation.used, conversation.limit)));
        meter.setAttribute("aria-valuetext", conversation.capacity);
      } else {
        for (const attribute of ["aria-valuemax", "aria-valuenow", "aria-valuetext"]) meter.removeAttribute(attribute);
      }
      find(".capacity-fill").style.width = `${conversation.percent ?? 0}%`;
      find(".capacity-fill").classList.toggle("near-limit", conversation.nearLimit);
      find(".expiry-value").textContent = conversation.expiryText;
      find(".expiry-value").title = conversation.expiryTitle;
      find(".expiry-value").setAttribute("aria-label", conversation.expiryTitle);
      if (conversation.expiry) find(".expiry-value").setAttribute("datetime", conversation.expiry);
      else find(".expiry-value").removeAttribute("datetime");
    }
    return { update };
  }
  globalThis.M365ExtenderPanel = Object.freeze({ mount });
})();
