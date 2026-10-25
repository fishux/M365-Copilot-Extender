// @ts-check
/* Toolbar popup UI. Never injected into Copilot. */
(() => {
  "use strict";
  const core = globalThis.M365ExtenderCore;
  const STALE_MS = 5 * 60 * 1000;
  const COPILOT_URL = "https://m365.cloud.microsoft/chat";
  const css = `
    :host { all: initial; color-scheme: light dark; --bg:#fbfaf6; --bar:#f1efe8; --line:#d8d5cc; --soft:#e2dfd6;
      --ink:#24292f; --muted:#57606a; --good:#1a7f37; --low:#9a6700; --critical:#cf222e; --accent:#0550ae;
      --hover:#eef1f5; --current:#e3eaf4; --warn-bg:#fff8e6; --warn-line:#e8cf8f;
      font:12px/20px ui-monospace,"SF Mono","Cascadia Mono",Consolas,Menlo,monospace; color:var(--ink); }
    * { box-sizing:border-box; } [hidden] { display:none !important; }
    button { font:inherit; color:inherit; }
    button:focus-visible, a:focus-visible, time:focus-visible { outline:2px solid var(--accent); outline-offset:-2px; border-radius:4px; }
    .card { width:340px; background:var(--bg); }
    .header { display:flex; justify-content:space-between; align-items:center; padding:8px 14px; background:var(--bar); border-bottom:1px solid var(--line); }
    .brand { color:var(--accent); } .route { color:var(--muted); }
    .body { padding:10px 14px 12px; display:flex; flex-direction:column; gap:12px; }
    .status { display:flex; align-items:flex-start; gap:8px; margin:0; }
    .dot { width:7px; height:7px; margin-top:7px; border-radius:50%; background:var(--muted); flex:none; }
    .dot.good { background:var(--good); } .dot.warn { background:var(--low); } .dot.critical { background:var(--critical); }
    .status.good { color:var(--good); } .status.warn { color:var(--low); } .status.critical { color:var(--critical); }
    .heading { display:flex; justify-content:space-between; gap:8px; margin:0; font:inherit; color:var(--muted); }
    .line { display:flex; align-items:center; min-height:20px; } .key { width:84px; flex:none; }
    .number { width:40px; text-align:right; font-weight:700; font-variant-numeric:tabular-nums; }
    .level { margin-left:12px; } .low { color:var(--low); } .critical { color:var(--critical); } .ok { color:var(--good); }
    .stale .number, .stale .level { opacity:.7; }
    .note { margin:0; color:var(--muted); }
    .models { display:flex; flex-direction:column; margin:2px -6px 0; }
    .model-row { display:flex; align-items:center; border-radius:4px; }
    .model-row:hover { background:var(--hover); } .model-row.active { background:var(--current); }
    .pick { flex:1; min-width:0; display:flex; align-items:center; height:32px; padding:0 6px; border:0; background:none; text-align:left; cursor:pointer; font-size:13px; }
    .pick:disabled { cursor:default; } .models.locked .model-row:not(.active) { opacity:.6; }
    .models.locked .model-row:hover { background:none; } .models.locked .model-row.active:hover { background:var(--current); }
    .mark { width:14px; flex:none; color:var(--good); }
    .badge { width:48px; flex:none; color:var(--accent); } .name { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
    .tag { color:var(--muted); font-size:11px; } .tag.selected { color:var(--low); }
    .chip { height:24px; padding:0 8px; border:1px solid var(--line); border-radius:5px; background:none; cursor:pointer; font-size:11px; }
    .chip:hover { background:var(--hover); }
    .capacity-count { font-variant-numeric:tabular-nums; } .capacity-count.near-limit { color:var(--low); }
    .capacity-track { width:80px; height:8px; margin-right:8px; flex:none; border-radius:2px; background:var(--soft); overflow:hidden; position:relative; }
    .capacity-track::after { content:""; position:absolute; inset:0; background:repeating-linear-gradient(90deg,transparent 0 7px,var(--bg) 7px 8px); }
    .capacity-fill { height:100%; background:var(--ink); } .capacity-fill.near-limit { background:var(--low); }
    .box { padding:8px 10px; border-radius:6px; font-size:11px; line-height:17px; }
    .warning { margin:0; border:1px solid var(--warn-line); background:var(--warn-bg); color:var(--low); }
    .error { margin:0; color:var(--critical); }
    .open-copilot { display:flex; align-items:center; justify-content:center; height:34px; border:1px solid var(--line); border-radius:6px; color:var(--accent); text-decoration:none; }
    .open-copilot:hover { background:var(--hover); }
    .details { padding:8px 10px; border-radius:6px; background:var(--bar); }
    dl { margin:4px 0 0; } dl > div { display:flex; align-items:center; min-height:24px; border-bottom:1px dashed var(--line); }
    dt { width:84px; flex:none; color:var(--muted); } dd { margin:0; flex:1; min-width:0; overflow-wrap:anywhere; }
    .disclaimer { margin:8px 0 0; font-size:11px; line-height:17px; color:var(--muted); }
    .footer { display:flex; align-items:center; justify-content:space-between; padding-top:8px; border-top:1px solid var(--line); }
    .save-label { color:var(--muted); font-size:11px; }
    .toggle { display:flex; align-items:center; gap:6px; }
    @media(prefers-color-scheme:dark) { :host { --bg:#0d1117; --bar:#161b22; --line:#30363d; --soft:#21262d; --ink:#e6edf3; --muted:#8b949e;
      --good:#7ee787; --low:#f2cc60; --critical:#ff7b72; --accent:#79c0ff; --hover:#161e29; --current:#1f2937; --warn-bg:#19160d; --warn-line:#5c4a17; } }
    @media(prefers-reduced-motion:reduce) { * { transition:none !important; } }
  `;
  const levelTag = { normal: "[ok]", warn: "[low]", critical: "[crit]" };
  const levelClass = { normal: "ok", warn: "low", critical: "critical" };

  function mount(host, { onConversationSelect }) {
    const root = host.attachShadow({ mode: "open" });
    const style = document.createElement("style"); style.textContent = css; root.append(style);
    const card = document.createElement("section");
    card.className = "card";
    card.setAttribute("aria-label", "M365 Copilot Extender");
    // This template is static; server and saved values are set with textContent.
    card.innerHTML = `<header class="header"><span class="brand">copilot-extender</span><span class="route"></span></header>
      <div class="body">
      <p class="status" role="status" aria-live="polite"><span class="dot"></span><span class="status-text">Waiting for Copilot</span></p>
      <p class="error" role="alert" hidden></p>
      <section class="usage" aria-label="Model usage allowance" hidden><h2 class="heading"><span class="usage-title"># allowance</span><span class="usage-age"></span></h2>
      <div class="counts"><div class="line"><span class="key">daily</span><span class="number daily">—</span><span class="level daily-level"></span></div>
      <div class="line"><span class="key">weekly</span><span class="number weekly">—</span><span class="level weekly-level"></span></div></div>
      <p class="note pending" hidden>Send a message to check your allowance.</p></section>
      <p class="warning box" hidden></p>
      <section aria-label="Models"><h2 class="heading"><span># model · this chat</span><span class="models-hint"></span></h2>
      <div class="models" role="radiogroup" aria-label="Model for this chat"></div></section>
      <section class="conversation-info" aria-label="This conversation" hidden><h2 class="heading"># chat</h2>
      <div class="line"><span class="key">model</span><span class="chat-model"></span></div>
      <div class="line"><span class="key">capacity</span><span class="capacity-track" role="meter" aria-label="Conversation messages used" aria-valuemin="0" hidden><span class="capacity-fill"></span></span><span class="capacity-count">Awaiting reply</span></div>
      <div class="line"><span class="key">expiry</span><time class="expiry-value" tabindex="0">Not reported</time></div></section>
      <section class="no-tab" hidden><p class="note">Readings come from the active Copilot tab.</p></section>
      <a class="open-copilot" href="${COPILOT_URL}" target="_blank" rel="noopener" hidden>open m365 copilot ↗</a>
      <section class="details" id="details" aria-label="Connection details" hidden><h2 class="heading"># connection</h2>
      <dl><div><dt>route</dt><dd class="route-value">Not observed</dd></div><div><dt>request</dt><dd class="requested">Not observed</dd></div>
      <div><dt>response</dt><dd class="response">Not observed</dd></div><div><dt>origin</dt><dd class="origin">Not observed</dd></div></dl>
      <p class="disclaimer">Requests your selected model. Microsoft controls final routing; origin doesn't verify the backend model.</p></section>
      <footer class="footer"><span class="save-label">Loading saved choices…</span>
      <button type="button" class="chip toggle" aria-expanded="false" aria-controls="details">details <span class="arrow" aria-hidden="true">▾</span></button></footer>
      </div>`;
    root.append(card);
    /** @type {(selector: string) => HTMLElement} */
    const find = selector => /** @type {HTMLElement} */ (card.querySelector(selector));
    /** @type {{ tone: string, row: HTMLElement, pick: HTMLButtonElement, mark: HTMLElement, tag: HTMLElement }[]} */
    const rows = [];
    for (const model of core.MODELS) {
      const row = document.createElement("div"); row.className = "model-row"; row.dataset.tone = model.tone;
      const pick = document.createElement("button"); pick.type = "button"; pick.className = "pick"; pick.setAttribute("role", "radio");
      const mark = document.createElement("span"); mark.className = "mark"; mark.setAttribute("aria-hidden", "true");
      const badge = document.createElement("span"); badge.className = "badge"; badge.textContent = `[${model.badge}]`;
      const name = document.createElement("span"); name.className = "name"; name.textContent = model.menuLabel;
      const tag = document.createElement("span"); tag.className = "tag";
      pick.append(mark, badge, name, tag);
      // Choosing a model keeps this chat on it until another is chosen.
      pick.addEventListener("click", () => onConversationSelect(model.tone));
      row.append(pick);
      find(".models").append(row);
      rows.push({ tone: model.tone, row, pick, mark, tag });
    }
    const toggle = find(".toggle");
    toggle.addEventListener("click", () => {
      const open = toggle.getAttribute("aria-expanded") !== "true";
      toggle.setAttribute("aria-expanded", String(open));
      find(".arrow").textContent = open ? "▴" : "▾";
      find(".details").hidden = !open;
      // Chrome caps popups at 600px, so the open drawer may need scrolling to.
      if (open) find(".details").scrollIntoView({ block: "nearest" });
    });
    function update(state) {
      const observed = state.observed || {};
      const connected = state.connected !== false;
      // A chat's own selection overrides the default; new chats start on the default.
      const tone = state.conversationTone || state.tone;
      const locked = !state.loaded || state.saving || !state.conversation;
      find(".models").classList.toggle("locked", locked);
      find(".models-hint").textContent = state.conversation ? "click to switch" : connected ? "new chat: send a message first" : "open a chat to switch";
      for (const { tone: rowTone, row, pick, mark, tag } of rows) {
        const model = core.modelFor(rowTone);
        const isDefault = rowTone === state.tone, isActive = rowTone === tone;
        row.classList.toggle("active", isActive);
        mark.textContent = isActive ? "›" : "";
        pick.disabled = locked;
        pick.setAttribute("aria-checked", String(isActive));
        pick.setAttribute("aria-label", `${model?.label}${isDefault ? ", default" : ""}`);
        pick.title = state.conversation ? `Use ${model?.label} in this chat from now on` : "Send a message in a chat to choose its model";
        tag.textContent = isDefault ? "default" : isActive ? "this chat" : "";
        tag.classList.toggle("selected", isActive && !isDefault);
      }
      find(".save-label").textContent = state.saving ? "Saving…" : state.loaded ? "Saved per chat · all tabs" : "Loading saved choices…";
      find(".error").hidden = !state.error;
      find(".error").textContent = state.error || "";

      const quota = core.allowanceFor(tone, observed.metering);
      const model = core.modelFor(tone);
      const knownPaidCap = model?.knownPaidCap && observed.scenario !== "OfficeWebIncludedCopilot";
      // Allowances are saved across tabs, so they show even outside Copilot.
      find(".usage").hidden = !(quota || (connected && knownPaidCap));
      find(".usage-title").textContent = `# allowance · ${model?.label ?? "—"}`;
      find(".counts").hidden = !quota;
      find(".pending").hidden = Boolean(quota);
      find(".pending").textContent = observed.quotaAt ? "Copilot did not report this allowance." : "Send a message to check your allowance.";
      const stale = Boolean(quota && observed.quotaAt && Date.now() - observed.quotaAt > STALE_MS);
      find(".usage").classList.toggle("stale", stale);
      find(".usage-age").textContent = quota && observed.quotaAt
        ? `${stale ? "last known" : "updated"} ${new Date(observed.quotaAt).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}` : "";
      for (const period of ["daily", "weekly"]) {
        const value = quota?.[period];
        const known = typeof value === "number";
        const level = core.allowanceLevel(known ? value : null);
        find(`.${period}`).textContent = known ? String(value) : "—";
        find(`.${period}`).classList.toggle("low", level === "warn");
        find(`.${period}`).classList.toggle("critical", level === "critical");
        find(`.${period}-level`).textContent = known ? levelTag[level] : "";
        find(`.${period}-level`).className = `level ${period}-level ${known ? levelClass[level] : ""}`;
      }

      let text = connected ? "Ready for your next message" : "No Copilot tab";
      let status = "";
      if (connected && !state.hookReady) text = "Connecting to Copilot…";
      const applied = core.modelFor(observed.appliedTone);
      if (connected && applied) {
        text = `Last request sent via ${applied.label}`;
        if (applied.tone !== tone) text += ` · next: ${model?.label}`;
        status = "good";
      }
      if (observed.result === "OutOfCredits") { text = "The last request reached a usage limit."; status = "critical"; }
      else if (observed.result && observed.result !== "Success") { text = `Copilot response: ${observed.result}`; status = "warn"; }
      find(".status-text").textContent = text;
      find(".dot").className = `dot ${status}`;
      find(".status").className = `status ${status}`;
      find(".warning").hidden = !observed.warning;
      find(".warning").textContent = observed.warning || "";

      const scenario = observed.scenario === "OfficeWebPaidCopilot" ? "Paid Copilot" : observed.scenario === "OfficeWebIncludedCopilot" ? "Included Copilot" : "";
      find(".route").textContent = scenario ? scenario.split(" ")[0].toLowerCase() : "";
      find(".route-value").textContent = scenario || "Not observed";
      find(".requested").textContent = core.modelFor(observed.appliedTone)?.label || "Not observed";
      find(".response").textContent = observed.result || "Not observed";
      find(".origin").textContent = observed.origin || "Not observed";

      find(".no-tab").hidden = connected;
      find(".open-copilot").hidden = connected;
      find(".conversation-info").hidden = !connected;
      find(".chat-model").textContent = `${model?.label ?? "—"} (${state.conversationTone ? "selected" : "default"})`;
      const conversation = core.conversationView(observed.conversation, observed.expiresAt);
      find(".capacity-count").textContent = conversation.capacity;
      find(".capacity-count").title = conversation.capacityTitle;
      find(".capacity-count").classList.toggle("near-limit", conversation.nearLimit);
      const meter = find(".capacity-track");
      meter.hidden = conversation.percent === null;
      if (conversation.percent !== null) {
        meter.setAttribute("aria-valuemax", String(conversation.limit));
        meter.setAttribute("aria-valuenow", String(Math.min(conversation.used ?? 0, conversation.limit ?? 0)));
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
