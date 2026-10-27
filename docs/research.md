# Research baseline

Reviewed 19 October 2026. This is source inspection, not an independent live Copilot test. The owner's prior successful DevTools override is user-reported evidence.

## Primary sources reviewed

- `cramt/m365-copilot-proxy` at [`65905e881d6597040966aff9b087aa9a908b358a`](https://github.com/cramt/m365-copilot-proxy/tree/65905e881d6597040966aff9b087aa9a908b358a): [API field notes](https://github.com/cramt/m365-copilot-proxy/blob/65905e881d6597040966aff9b087aa9a908b358a/docs/m365-copilot-api.md), [model mapping and scenario routing](https://github.com/cramt/m365-copilot-proxy/blob/65905e881d6597040966aff9b087aa9a908b358a/packages/core/src/copilot.ts), [session transport](https://github.com/cramt/m365-copilot-proxy/blob/65905e881d6597040966aff9b087aa9a908b358a/packages/core/src/session.ts), and [quota parsing](https://github.com/cramt/m365-copilot-proxy/blob/65905e881d6597040966aff9b087aa9a908b358a/packages/core/src/priority-access.ts).
- `aweussom/i-want-my-model-picker-back` at [`c4e8982ddfcc844286395a91fd097dbcbf88fd67`](https://github.com/aweussom/i-want-my-model-picker-back/tree/c4e8982ddfcc844286395a91fd097dbcbf88fd67): [README and controlled probes](https://github.com/aweussom/i-want-my-model-picker-back/blob/c4e8982ddfcc844286395a91fd097dbcbf88fd67/README.md) and [`force_tone.js`](https://github.com/aweussom/i-want-my-model-picker-back/blob/c4e8982ddfcc844286395a91fd097dbcbf88fd67/force_tone.js).

The picker repository is currently a browser automation / Electron application and API wrapper, not an MV3 extension template.

## Transport and overrides

Both sources describe SignalR JSON over WebSockets with record separator `\x1e`. A payload can contain multiple records. The picker interceptor wraps `WebSocket.prototype.send`, accepts type `1` or `4` with `target === "chat"` and an object at `arguments[0]`, then changes only its `tone`. It preserves unrelated records and the trailing separator, reports successful overrides, and reports unsupported binary data rather than pretending to modify it.

The proxy sends type `4` chat invocations alongside a Metrics record. An extension modifying an existing browser request should preserve the browser's other records and fields. Early injection into the page's JavaScript world is an implementation requirement to investigate; isolated-world interception must not be assumed to reach the page's sockets.

## Model tones

| Intended selection | Tone in current proxy source |
| --- | --- |
| GPT-6.1 Sol | `Gpt_61_Sol_Reasoning` |
| GPT-6 Sol | `Gpt_6_Sol_Reasoning` |
| GPT-6 Think Deeper | `Gpt_6_Reasoning` |
| GPT-5.6 Think Deeper | `Gpt_5_6_Reasoning` |
| GPT-5.5 Think Deeper | `Gpt_5_5_Reasoning` |
| Auto | `magic` |

Auto casing differs between sources: the proxy sends lowercase `magic`; the picker research captures `Magic`. Case equivalence is not established here. These are requested tones, not independently authenticated model identities.

## Routing and quotas

The proxy distinguishes `OfficeWebPaidCopilot` and `OfficeWebIncludedCopilot`. Its findings say scenario, account entitlements, and agent context can affect service behavior; changing `licenseType` alone does not grant a license. The owner's account was observed on the paid scenario. Preserve that observed route unless a later requirement authorizes and validates a change.

The proxy field notes report paid-scenario GPT-6.1 Sol allowances at `throttling.metering.GPT61SolQueryDaily.remainingAllowance` and `throttling.metering.GPT61SolQueryWeekly.remainingAllowance`, with reported limits of 40 daily and 75 weekly. Its observations say included-scenario turns lack this metering map and did not decrement these counters. That is a source claim requiring account-specific validation, not a guarantee of unlimited usage. Missing fields must mean unknown, not zero or unlimited.

The proxy's typed priority-access budgets currently cover Opus and Sonnet 5.5, even though its docs discuss GPT-6.1 Sol counters. Do not assume the existing helper already tracks GPT-6.1 Sol. This review did not establish a separate GPT-6 Sol allowance.

Quota refusals may appear only in the final type `2` item with `result.value === "OutOfCredits"` and a `BotConnection` reply. Streaming text alone is insufficient to detect all quota outcomes. GPT-6.1 Sol reset behavior and refusal wording still need live confirmation.

## Verification limits and next checks

`contentOrigin: "DeepLeo"` is used by the proxy as evidence of a serving route; `BotConnection` can indicate a canned refusal. Neither authenticates a precise model version. The proxy also uses model self-identification prompts, which are not reliable independent proof. Distinguish requested tone, observed routing/response metadata, and verified model identity.

Before implementing claims about current backend behavior, capture sanitized observations in the owner's real Chrome/Copilot session: original and modified tone, scenario, override confirmation, response result and content origin, and available quota fields before/after a turn. A control request and an invalid-tone comparison can distinguish a working interceptor from an ignored override. Account for new conversations versus follow-up turns and any agent-specific routing. Do not publish access tokens, full socket URLs, chat text, tenant/account identifiers, or raw captures.

## Follow-up review, 25 October 2026

Same upstream revisions as above (`65905e8`, `c4e8982`; neither had new commits). Source inspection only.

### Conversation identity

- The proxy's API notes (§1, §8) list `ConversationId` as a query parameter on the Chathub socket URL: a client-generated UUID reused across turns to keep server-side context. Each proxy turn opens a fresh WebSocket with the same `ConversationId`; `isStartOfSession` is true only on the first turn.
- `m365-gui-emulate.mjs` captures the real web client's Chathub URL and replaces its `ConversationId`, which shows the browser client sends one too.
- The picker's Electron driver (`driver.js`, `currentConversationId`) reads the conversation from the page URL `/chat/conversation/<id>` and falls back to the socket URL's `ConversationId`, noting that navigation can lag a turn. It treats both as the same ID.
- The picker README says `tone` is per conversation: a new chat resets to `Magic`, and server-rendered history carries a tone per conversation.

Conclusion: a per-conversation model is feasible from the socket URL alone, consistently from the first message of a new chat. Live confirmation of the parameter name and casing in the current client is still required (see the architecture acceptance checks).

### Auto tone casing

The picker captured `Magic` from the web client's own `modelSelectorMetadata` and `availableTones`. The proxy sends lowercase `magic` and is served. The extension keeps sending the client's `Magic` and treats tones case-insensitively when comparing.

### Metering keys reviewed 25 October 2026

Every keyed allowance observed upstream on the paid scenario:

| Key(s) | Model | Notes |
| --- | --- | --- |
| `GPT61SolQueryDaily` / `GPT61SolQueryWeekly` | GPT-6.1 Sol (`Gpt_61_Sol_Reasoning`) | 40/day, 75/week; spent only on the paid scenario |
| `ClaudeOpusQueryDaily` / `ClaudeOpusQuery75` | Claude Opus (`Claude_Opus`) | 40/day; the weekly key is named `…75`, not `…Weekly`; resets at midnight UTC (weekly on Monday) |
| `ClaudeSonnet55QueryDaily` / `ClaudeSonnet55QueryWeekly` | Claude Sonnet 5.5 (`Claude_Sonnet`, paid) | 80/day, 150/week; the weekly key was seen resetting daily |
| `ClaudeOpusQuery`, `…Dev`, `…HourlyDev`, `…DailyWord`, `…WeeklyWord`, `…C1`, `DeepResearch` | Not tied to a selectable tone | Present in the map; meaning unestablished |

Findings for the models this extension offers:

- **GPT-6 (`Gpt_6_Reasoning`)** has no priority-access allowance; it is limited by the per-conversation cap and the thread-rate throttle. Its candidate prefix was removed.
- **GPT-6 Sol and GPT-5.5** turns lowered neither the Opus nor the Sonnet 5.5 keys, and no GPT-6 Sol or GPT-5.5 key appears in the observed maps. They are most likely unmetered; the name-based prefixes are kept only so a future key would be shown.
- **GPT-5.6** has no observed key and no direct observation either way.
- **Included-scenario** turns carry no `metering` map at all.

Not done: the Claude tones are not offered by this extension.

## Work IQ capture, 27 October 2026

Owner's sanitized capture of outgoing chat requests (extension disabled, paid account, GPT-6 Sol then Auto chosen in Copilot's picker): structure only, no prompt text, tokens or IDs.

- Work IQ replaced the Work/Web toggle in mid-2026 (third-party write-ups; Microsoft's support pages describe the switch but not the protocol). Switching it starts a new chat.
- **Off:** Chathub URL `agent=web`, `scenario=OfficeWebPaidCopilot`; no `gpts`; `clientPreferences.executionControls` empty.
- **On:** `agent=work`, `scenario=officeweb`; `gpts: [{ id: "bizchat-as-gpt-scenario", source: "BuiltInAgents", clientOverrides: { capabilities: GraphConnectors (tenant connections), WebSearch, WorkSearch } }]`; matching `executionControls`; 17 `enterprise_*`/plugin `optionsSets` added and 7 removed; a few URL `variants` differ.
- **The model still travels in `tone`.** GPT-6 Sol sent `Gpt_6_Sol_Chat` in both modes (not the `Gpt_6_Sol_Reasoning` this extension sends); Auto sent `Magic`.
- `bizchat-as-gpt-scenario` appears as the built-in `DeclarativeCopilot` agent in HEXUXIU/M365-Copilot2API's HAR notes (`docs/har-mining/02-hidden-endpoints.md`). The cramt proxy found an attached agent overrides non-default and `*_Reasoning` tones with GPT-5, which matches the reported Work IQ symptom.

Out of scope by decision (27 October 2026): Work IQ tones for GPT-6.1 Sol, GPT-6, GPT-5.6 and GPT-5.5. Only GPT-6 Sol is supported with Work IQ; the others fall back to it. Still open: whether `Gpt_6_Sol_Chat` has replaced `Gpt_6_Sol_Reasoning` outside Work IQ. Copilot's picker shows the same model list with Work IQ on and off. Opus's weekly key would need an explicit key name in `models.json` (its suffix is `75`, not `Weekly`) if Claude models are added later.
