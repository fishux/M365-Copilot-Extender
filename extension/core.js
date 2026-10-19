/* Shared, dependency-free protocol helpers and panel renderer. */
(() => {
  "use strict";
  const MODELS = Object.freeze([
    { tone: "Gpt_61_Sol_Reasoning", label: "GPT-6.1 Sol", subtitle: "Reasoning", prefix: "GPT61SolQuery", capped: true },
    { tone: "Gpt_6_Sol_Reasoning", label: "GPT-6 Sol", subtitle: "Reasoning", prefix: "GPT6SolQuery" },
    { tone: "Gpt_6_Reasoning", label: "GPT-6", subtitle: "Think deeper", prefix: "GPT6Query" },
    { tone: "Gpt_5_6_Reasoning", label: "GPT-5.6", subtitle: "Think deeper", prefix: "GPT56Query" },
    { tone: "Gpt_5_5_Reasoning", label: "GPT-5.5", subtitle: "Think deeper", prefix: "GPT55Query" },
    { tone: "Magic", label: "Auto", subtitle: "Let Copilot decide", prefix: null }
  ].map(Object.freeze));
  const DEFAULT_TONE = MODELS[0].tone;
  const modelFor = tone => MODELS.find(model => model.tone === tone);
  const isRecord = value => value !== null && typeof value === "object" && !Array.isArray(value);
  const count = value => Number.isSafeInteger(value) && value >= 0 ? value : null;
  const RS = "\x1e";

  function connectionInfo(rawURL) {
    try {
      const url = new URL(rawURL);
      if (url.protocol !== "wss:" || url.hostname !== "substrate.office.com" ||
          !/^\/m365Copilot\/Chathub\//i.test(url.pathname)) return null;
      // Never return full URLs: they may contain access tokens and account IDs.
      const scenario = url.searchParams.get("scenario");
      return { scenario: ["OfficeWebPaidCopilot", "OfficeWebIncludedCopilot"].includes(scenario) ? scenario : "Unknown" };
    } catch { return null; }
  }

  function isChat(frame) {
    return isRecord(frame) && [1, 4].includes(frame.type) && frame.target === "chat" &&
      Array.isArray(frame.arguments) && isRecord(frame.arguments[0]);
  }

  function parseRecords(data) {
    if (typeof data !== "string") return [];
    return data.split(RS).flatMap(part => {
      try { const frame = JSON.parse(part); return isRecord(frame) ? [frame] : []; }
      catch { return []; }
    });
  }

  function rewrite(data, tone) {
    const tones = [];
    if (typeof data !== "string" || !modelFor(tone)) return { data, tones };
    const parts = data.split(RS).map(part => {
      try {
        const frame = JSON.parse(part);
        if (!isChat(frame)) return part;
        tones.push(tone);
        if (frame.arguments[0].tone === tone) return part;
        frame.arguments[0].tone = tone;
        return JSON.stringify(frame);
      } catch { return part; }
    });
    return { data: tones.length ? parts.join(RS) : data, tones };
  }

  function parseMetering(raw) {
    if (!isRecord(raw)) return {};
    const metering = {};
    for (const [key, value] of Object.entries(raw).slice(0, 100)) {
      // Only budget identifiers and numeric remaining allowances cross the bridge.
      if (!/^[A-Za-z][A-Za-z0-9_]{0,79}$/.test(key)) continue;
      const remaining = count(value?.remainingAllowance);
      if (remaining !== null) metering[key] = remaining;
    }
    return metering;
  }

  function allowanceFor(tone, metering) {
    const model = modelFor(tone);
    if (!model?.prefix) return null;
    // GPT61Sol keys are documented. Other model-specific prefixes are only
    // recognized if the server actually supplies them; they are never invented.
    // No totals or reset schedules are inferred from a budget name.
    const normalize = key => key.replaceAll("_", "").toLowerCase();
    const entries = Object.entries(metering || {});
    const read = suffix => {
      const entry = entries.find(([key]) => normalize(key) === normalize(model.prefix + suffix));
      return entry ? count(entry[1]) : null;
    };
    const daily = read("Daily"), weekly = read("Weekly");
    return daily === null && weekly === null ? null : { daily, weekly };
  }

  const allowedOrigins = ["DeepLeo", "BotConnection", "3PDeclarativeAgent", "CodeGenerator"];
  const allowedResults = ["Success", "OutOfCredits", "Throttled", "ForbiddenRequest", "InternalError", "NotFound"];
  function readMetadata(frame) {
    const candidates = [];
    if (frame.type === 2 && isRecord(frame.item)) candidates.push({ value: frame.item, final: true });
    if (frame.type === 1 && frame.target === "update" && Array.isArray(frame.arguments)) {
      for (const value of frame.arguments) if (isRecord(value)) candidates.push({ value, final: false });
    }
    return candidates.flatMap(({ value, final }) => {
      const output = {};
      if (isRecord(value.throttling)) {
        if (isRecord(value.throttling.metering) || final) {
          output.metering = parseMetering(value.throttling.metering);
        }
        const used = count(value.throttling.numUserMessagesInConversation);
        const limit = count(value.throttling.maxNumUserMessagesInConversation);
        if (used !== null && limit !== null) output.conversation = { used, limit };
      } else if (final) {
        output.metering = {};
      }
      if (allowedResults.includes(value.result?.value)) output.result = value.result.value;
      else if (typeof value.result?.value === "string") output.result = "Other";
      for (const message of Array.isArray(value.messages) ? value.messages : []) {
        if (message?.author === "bot" && allowedOrigins.includes(message.contentOrigin)) output.origin = message.contentOrigin;
      }
      if (Object.keys(output).length === 0) return [];
      return [{ ...output, final }];
    });
  }

  globalThis.M365ExtenderCore = Object.freeze({ MODELS, DEFAULT_TONE, modelFor, connectionInfo,
    isChat, parseRecords, rewrite, parseMetering, allowanceFor, readMetadata, count,
    allowedOrigins: Object.freeze(allowedOrigins), allowedResults: Object.freeze([...allowedResults, "Other"]) });
})();

/* Shared UI for the in-page panel and toolbar popup. */
(() => {
  "use strict";
  const core = globalThis.M365ExtenderCore;
  const STALE_MS = 5 * 60 * 1000;
  const css = `
    :host { all: initial; color-scheme: light dark; --bg:#fff; --ink:#1c2132;
      --muted:#697185; --line:#e7e9f0; --soft:#f6f7fb; --accent:#6254d9;
      --tint:#efedff; --good:#248568; --warn:#b15722; font:13px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif; color:var(--ink); }
    * { box-sizing:border-box; } [hidden] { display:none !important; }
    .card { width:304px; max-width:calc(100vw - 24px); background:var(--bg); border:1px solid var(--line); border-radius:18px;
      box-shadow:0 14px 50px #17213b1a,0 2px 8px #17213b08; overflow:hidden; }
    .popup { width:328px; max-width:none; border:0; border-radius:0; box-shadow:none; }
    .header { display:flex; align-items:center; gap:10px; padding:18px 18px 14px; }
    .logo { display:grid; place-items:center; width:32px; height:32px; flex:none; border-radius:10px;
      color:white; background:linear-gradient(145deg,#8c76ed,#5145bb); box-shadow:0 3px 9px #6254d926; }
    .logo svg { width:20px; height:20px; }
    .brand { flex:1; } h2 { margin:0; font-size:14px; font-weight:650; letter-spacing:-.2px; }
    .eyebrow { font-size:10px; letter-spacing:1.1px; text-transform:uppercase; color:var(--muted); }
    button,select,input { font:inherit; } button { cursor:pointer; }
    .collapse { width:28px; height:28px; border:0; border-radius:8px; color:var(--muted); background:transparent; font-size:20px; }
    .collapse:hover { background:var(--soft); } .body { padding:0 18px 17px; }
    label.field-label { display:block; margin:0 0 7px; font-size:11px; font-weight:600; color:var(--muted); }
    .select-wrap { position:relative; } select { display:block; width:100%; min-height:45px;
      appearance:none; padding:11px 35px 11px 13px; border:1px solid var(--line); border-radius:11px;
      background:var(--soft); color:var(--ink); font-size:14px; font-weight:600; cursor:pointer; }
    select:disabled { cursor:wait; opacity:.6; } .chevron { position:absolute; right:14px; top:14px; pointer-events:none; color:var(--muted); }
    :is(button,select,input,summary):focus-visible { outline:3px solid var(--accent); outline-offset:3px; }
    .hint { display:flex; align-items:center; gap:6px; margin:8px 0 0; font-size:11px; color:var(--muted); }
    .hint svg { width:12px; height:12px; flex:none; color:var(--accent); }
    .usage { margin-top:16px; padding:13px; border:1px solid var(--line); border-radius:12px; }
    .usage-head { display:flex; align-items:center; justify-content:space-between; gap:8px; margin-bottom:9px; }
    .usage-head strong { font-size:11px; font-weight:600; } .live { font-size:9px; letter-spacing:.6px; text-transform:uppercase; color:var(--good); }
    .counts { display:grid; grid-template-columns:1fr 1fr; gap:12px; }
    .counts > div + div { border-left:1px solid var(--line); padding-left:14px; }
    .number { display:block; font-size:26px; line-height:1.2; font-weight:600; letter-spacing:-1px; font-variant-numeric:tabular-nums; }
    .number.zero { color:var(--warn); } .period { font-size:10px; color:var(--muted); }
    .usage-note { margin:10px 0 0; font-size:10px; color:var(--muted); } .pending { margin:0; font-size:11px; color:var(--muted); }
    .status { display:flex; align-items:flex-start; gap:7px; margin:14px 0 0; font-size:11px; color:var(--muted); }
    .dot { width:6px; height:6px; margin-top:5px; border-radius:50%; background:#a7aec0; flex:none; }
    .dot.good { background:var(--good); } .dot.warn { background:var(--warn); }
    details { margin-top:13px; padding-top:12px; border-top:1px solid var(--line); }
    summary { cursor:pointer; color:var(--muted); font-size:11px; } dl { margin:12px 0 0; display:grid; grid-template-columns:80px 1fr; gap:6px 8px; font-size:10px; }
    dt { color:var(--muted); } dd { margin:0; overflow-wrap:anywhere; } .disclaimer { margin:11px 0 0; font-size:10px; color:var(--muted); }
    .settings { padding:14px 18px; border-top:1px solid var(--line); background:var(--soft); }
    .toggle-row { display:flex; align-items:center; justify-content:space-between; gap:12px; cursor:pointer; font-size:12px; font-weight:500; }
    .toggle-row small { display:block; margin-top:2px; font-size:10px; font-weight:400; color:var(--muted); }
    input[type=checkbox] { appearance:none; width:32px; height:19px; margin:0; border-radius:20px; background:#b7becd; position:relative; cursor:pointer; flex:none; }
    input[type=checkbox]:before { content:""; position:absolute; top:3px; left:3px; width:13px; height:13px; border-radius:50%; background:white; transition:transform .15s; }
    input[type=checkbox]:checked { background:var(--accent); } input[type=checkbox]:checked:before { transform:translateX(13px); }
    .error { color:var(--warn); font-size:11px; margin:10px 0 0; }
    .compact .body { display:none; } .compact .header { padding-bottom:18px; }
    @media(prefers-color-scheme:dark) { :host { --bg:#1c1d25; --ink:#f0f0f7; --muted:#a2a6b8; --line:#353641; --soft:#252631; --accent:#b5a6ff; --tint:#302943; --good:#76d3af; --warn:#f2b07d; } }
    @media(prefers-reduced-motion:reduce) { * { transition:none !important; } }
  `;
  const logo = '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 3v18M3 12h18M5.6 5.6l12.8 12.8M5.6 18.4L18.4 5.6" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/><circle cx="12" cy="12" r="4" fill="currentColor"/></svg>';
  const lock = '<svg viewBox="0 0 16 16" fill="none" aria-hidden="true"><rect x="3.5" y="7" width="9" height="7" rx="2" stroke="currentColor"/><path d="M5.5 7V4.5a2.5 2.5 0 015 0V7" stroke="currentColor"/></svg>';

  function mount(host, { popup = false, onSelect, onVisibility }) {
    const root = host.attachShadow({ mode: "open" });
    const style = document.createElement("style"); style.textContent = css; root.append(style);
    const card = document.createElement("section");
    card.className = `card${popup ? " popup" : ""}`;
    card.setAttribute("aria-label", "M365 Copilot Extender");
    // This template is static; server and preference values are set with textContent.
    card.innerHTML = `<header class="header"><span class="logo">${logo}</span><div class="brand"><div class="eyebrow">M365 Copilot</div><h2>Extender</h2></div>
      <button class="collapse" type="button" aria-label="Collapse panel" aria-expanded="true" ${popup ? "hidden" : ""}>−</button></header>
      <div class="body"><label class="field-label" for="model">Preferred model</label>
      <div class="select-wrap"><select id="model" disabled></select><span class="chevron" aria-hidden="true">⌄</span></div>
      <p class="hint">${lock}<span class="save-label">Loading saved preference…</span></p>
      <section class="usage" aria-label="Model usage allowance" hidden><div class="usage-head"><strong>Messages remaining</strong><span class="live">Server reported</span></div>
      <div class="counts"><div><span class="number daily">—</span><span class="period">Daily allowance</span></div><div><span class="number weekly">—</span><span class="period">Weekly allowance</span></div></div>
      <p class="pending" hidden>Send a message to check your allowance.</p><p class="usage-note"></p></section>
      <p class="status" role="status" aria-live="polite"><span class="dot"></span><span class="status-text">Waiting for Copilot</span></p>
      <p class="error" role="alert" hidden></p>
      <details><summary>Connection details</summary><dl><dt>Route</dt><dd class="route">Not observed</dd><dt>Last request</dt><dd class="requested">Not observed</dd><dt>Response</dt><dd class="response">Not observed</dd><dt>Origin</dt><dd class="origin">Not observed</dd><dt>Conversation</dt><dd class="conversation">Not reported</dd></dl>
      <p class="disclaimer">Requests your selected model. Microsoft controls final model routing.</p></details></div>
      <footer class="settings" ${popup ? "" : "hidden"}><label class="toggle-row"><span>Show panel on Copilot<small>Applies to all Copilot tabs</small></span><input type="checkbox" class="visibility" role="switch" checked disabled></label></footer>`;
    root.append(card);
    const find = selector => card.querySelector(selector);
    const select = find("select");
    for (const model of core.MODELS) {
      const option = document.createElement("option"); option.value = model.tone;
      option.textContent = model.label + (model.subtitle === "Think deeper" ? " · Think deeper" : "");
      select.append(option);
    }
    select.addEventListener("change", () => onSelect(select.value));
    find(".visibility").addEventListener("change", event => onVisibility?.(event.target.checked));
    find(".collapse").addEventListener("click", () => {
      const compact = card.classList.toggle("compact");
      find(".collapse").textContent = compact ? "+" : "−";
      find(".collapse").setAttribute("aria-expanded", String(!compact));
      find(".collapse").setAttribute("aria-label", compact ? "Expand panel" : "Collapse panel");
    });

    function update(state) {
      select.value = state.tone;
      select.disabled = !state.loaded || state.saving;
      find(".visibility").checked = state.inPageEnabled;
      find(".visibility").disabled = !state.loaded || state.saving;
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
      find(".conversation").textContent = observed.conversation ? `${observed.conversation.used} / ${observed.conversation.limit} messages` : "Not reported";
    }
    return { update };
  }
  globalThis.M365ExtenderPanel = Object.freeze({ mount });
})();
