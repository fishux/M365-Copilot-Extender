/* Shared, dependency-free protocol helpers. No chat text leaves these helpers. */
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
