# M365 Copilot Extender

A lightweight Chrome Manifest V3 extension for persistent model selection in [Microsoft 365 Copilot](https://m365.cloud.microsoft).

Every chat starts on **GPT-6 Sol**, the fixed default. Choose a different model for a chat in the Chrome toolbar popup and that chat stays on it from then on, in every tab and across browser restarts.

The toolbar icon shows which model the current tab is using and turns orange, then red, as that model's remaining allowance runs low. The popup lists every model on one screen in a compact console style: click a row to switch the chat you're viewing. A status line shows which model the last request was sent via. It follows the system light/dark appearance, and connection details sit in an expandable drawer. Nothing is displayed over the Copilot page. See the [changelog](CHANGELOG.md) for what's new.

## Install on Windows

1. Download `M365-Copilot-Extender-v<version>.zip` from the [latest release](https://github.com/fishux/M365-Copilot-Extender/releases/latest) and extract it to a permanent folder. (Alternatively, use **Code → Download ZIP** on this repository.)
2. In Google Chrome, open `chrome://extensions` and enable **Developer mode**.
3. Click **Load unpacked**. Select the extracted folder containing `manifest.json` (or this repository's `extension` folder).
4. Pin **M365 Copilot Extender** from Chrome's Extensions menu.
5. Reload any already-open Copilot tabs. Send a message, then use the toolbar popup to switch the chat's model.

When updating, replace the extension files in the same folder, click **Reload** on its `chrome://extensions` card, then reload Copilot. Keep the unpacked folder in place so Chrome can load it on restart. Saved choices and readings belong to this Chrome profile; they do not synchronize across separate profiles or devices.

## Models

<!-- models:start (generated from src/models.json by npm run build) -->
| Selection | Badge | Requested tone | Work IQ tone |
| --- | --- | --- | --- |
| GPT-6.1 Sol | `6.1` | `Gpt_61_Sol_Reasoning` | None: uses GPT-6 Sol |
| GPT-6 Sol (default) | `6S` | `Gpt_6_Sol_Reasoning` | `Gpt_6_Sol_Chat` |
| GPT-6 · Think deeper | `6` | `Gpt_6_Reasoning` | None: uses GPT-6 Sol |
| GPT-5.6 · Think deeper | `5.6` | `Gpt_5_6_Reasoning` | None: uses GPT-6 Sol |
| GPT-5.5 · Think deeper | `5.5` | `Gpt_5_5_Reasoning` | None: uses GPT-6 Sol |
| Auto | `A` | `Magic` | `Magic` |
<!-- models:end -->

Model matching ignores case: Copilot's web client sends `Magic` for Auto, other clients `magic`, and the service accepts both.

### Work IQ

Work IQ is Copilot's switch for grounding answers in Teams, Outlook, SharePoint and other work data. Turn it on or off in Copilot's page as usual; the extension follows it and has no switch of its own.

**Supported with Work IQ: GPT-6 Sol only** (plus Auto, which is Copilot's own default). This is deliberate scope, not a pending gap: other models are not planned for Work IQ.

- A Work IQ chat attaches Copilot's built-in `bizchat-as-gpt-scenario` agent. That agent ignores the `*_Reasoning` tones the extension normally sends and falls back to a GPT-5 model, which is why Work IQ used to drop chats to GPT-5.
- In a Work IQ chat the extension sends the model's **Work IQ tone** from the table above: `Gpt_6_Sol_Chat` for GPT-6 Sol, the value Copilot's own picker sends.
- Any other model chosen for a Work IQ chat is sent as **GPT-6 Sol**. In a Work IQ chat the popup dims those rows; the chat line, status line and toolbar badge name GPT-6 Sol. To use another model, turn Work IQ off for that chat.
- When the last request went through Work IQ, the status line shows a small `[work iq]` tag, the toolbar badge reads `6SW` (tooltip: "GPT-6 Sol · Work IQ"), and the **details** drawer has a `work iq` row (`On`, `Off` or `Not observed`).

A chat is treated as Work IQ when its Chathub connection carries `agent=work` or a request carries the built-in agent. Copilot starts a new chat whenever Work IQ is switched, so one chat never mixes the two.

### Adding or removing a model

Edit [`src/models.json`](src/models.json) — exactly one entry has `"default": true` — then run `npm run build`. The build inlines the list into every runtime script and regenerates the tables above and in the [architecture notes](docs/architecture.md). Each entry needs a `tone` (the exact value Copilot expects), `workIQTone` (the value Copilot's own picker sends for that model with Work IQ on, or `null` to use the default's; only GPT-6 Sol, the default, and Auto have one, and the default must), `label`, `menuLabel` (popup text), `badge` (toolbar text, at most 4 characters), `budgetPrefix` (a metering key prefix such as `GPT61SolQuery`, or `null`), `knownPaidCap` (show a waiting usage box before the first reading), and `budgetEvidence` (why that prefix is believed). Saved choices for a removed model fall back to the default.

The extension changes only `arguments[0].tone` on type `1`/`4` chat invocations to Copilot's Chathub socket. It preserves other SignalR records, request fields, Work IQ's agent and connectors, and the connection's existing paid/included scenario. It does not add license entitlements, change routing scenarios, or automatically switch models when a budget is exhausted.

Microsoft controls final model routing. An applied tone and response `contentOrigin` do **not** independently verify the exact backend model. Try a new conversation if changing a model in an existing conversation is refused or behaves unexpectedly.

## Usage allowances

Counters show **remaining messages** reported by Copilot, refreshed after a response. They do not estimate usage or assume that the reported 40/day and 75/week apply to every account.

- GPT-6.1 Sol uses the documented `throttling.metering.GPT61SolQueryDaily.remainingAllowance` and `GPT61SolQueryWeekly.remainingAllowance` fields. The parser reads both final type `2` items and type `1` update records.
- Other selectable models have no hardcoded quota. If Copilot supplies directly matching model-specific daily/weekly keys, the UI can display them; the mappings and their evidence limits are described in [architecture notes](docs/architecture.md).
- Models without a documented or observed budget have no usage box. GPT-6.1 Sol shows a waiting state before its first paid-route reading. Missing data is unknown, never zero or unlimited. Included-route GPT-6.1 Sol has no counter unless an allowance is actually reported.
- Zero remains visible. `OutOfCredits` is surfaced even if the refusal appears only in the final result. There is no guessed reset countdown or automatic fallback.
- Allowances are per account, so they stay visible when you switch to another conversation in the same tab.
- The last reported allowances are **saved in `chrome.storage.local`** and shared by every tab in this Chrome profile. Each new reply that reports an allowance updates the values it reports; others keep their last value. Allowances also show when the popup is opened outside Copilot. If you use two Microsoft accounts in one Chrome profile, the most recent reply's allowances are shown. Old readings are marked **Last known** after five minutes; they refresh only with a new response.

The status line at the top of the popup names the model the last request was sent via (and the next one, right after you switch), or the last error. The **details** drawer shows the observed scenario, the last requested model, result code, and known content origin. If a request reaches a usage limit, the status line says so; the popup never switches models automatically.

## Conversation capacity and expiry

The toolbar popup includes a compact chat section below the model list:

- **Capacity:** the server's used/maximum message counts and a slim meter. It highlights at 90% capacity. The maximum is read from Copilot, never assumed to be 600, and is separate from model allowances. Hover the count for remaining conversation messages.
- **Server expiry:** a short local date and time remaining, read from `item.conversationExpiryTime`. Hover or keyboard-focus the date for its exact local timestamp. A passed timestamp is labelled **Passed**; it is not a claim that the chat has been deleted. The server timestamp does not establish Microsoft's retention policy.

These readings belong to the conversation shown in the page. Each conversation keeps its own last observed reply in this tab, so switching chats in the sidebar shows that chat's readings, and a new, unsent chat shows none. Each chat's last readings (the model the last request was sent via, result, origin, capacity and expiry) are saved under a hash of its conversation ID, so reopening the chat in another tab or after a restart shows them without a new message. They are not a live query: a conversation that has never replied with the extension installed shows **Awaiting reply** for capacity and **Not reported** for expiry. The last valid values remain visible until Copilot reports replacements, including across partial or final replies that omit fields. Opening a background socket does not clear them, and replies on a connection where nothing was sent are ignored. Readings for connections without a conversation ID stay in that tab only. Countdown labels refresh locally without Microsoft requests.

## Per-conversation model

Every chat starts on the default, GPT-6 Sol. Click another model's row in the popup and the chat you're viewing uses it for every later turn, in any tab, and across restarts; other conversations are unaffected. Clicking the default's row returns the chat to the default. Chats that were never switched always follow the default, so if a future version changes the default they follow it too.

- Switching is available immediately, including on a new chat before its first message: the choice is held by the tab and applied to, and saved with, the chat when the first message is sent. Picking a model on a new chat that is later abandoned is simply forgotten.
- Conversations are identified by the `ConversationId` that Copilot's own client attaches to its Chathub connection (and shows in `/chat/conversation/<id>` URLs). Only a SHA-256 hash of that ID is saved, together with the model and a timestamp. The 2,000 most recently used chats keep their choice and readings; older entries are dropped.
- Before the first chat request on a connection, the hook briefly waits for that conversation's choice. If it never arrives within 1.5 seconds, the default model is used and a warning is shown.
- Copilot's own picker also stores a model per conversation on the server. The extension overrides what is sent; it does not change what Copilot's picker displays.

## Toolbar badge

The toolbar badge shows the model a Copilot tab is using (see the **Badge** column above). On other tabs it shows the default model. Hover the icon for the model's name and its lowest reported remaining allowance.

When the active model's lowest reported daily or weekly allowance is **20 or fewer**, the icon and badge turn **orange**; at **10 or fewer** they turn **red**. Popup counters use the same colours. Models with no reported allowance keep the normal colour.

## Privacy and permissions

The only extension API permission is `storage`, with content scripts scoped to `https://m365.cloud.microsoft/*`. There are no host permissions, third-party requests, analytics, credentials, remote scripts, or runtime dependencies. A background service worker only sets the toolbar badge, title and icon; it makes no network requests. `chrome.storage.local` holds per-conversation model choices and last readings keyed by hashed conversation IDs, plus the last reported allowances. The page hook parses chat frames locally, but forwards only allowlisted metadata, numeric quota values and the conversation ID; it does not forward or save prompts, replies, access tokens, full socket URLs, or account identifiers.

The page bridge is observable by scripts on the Copilot page. It is an instrumentation channel, not an authenticated model-identity or security boundary. See [architecture and limitations](docs/architecture.md).

## Development and validation

No build or dependency installation is required to load the prebuilt `extension/` folder in Chrome. Edit `src/` for development and run `npm run build` to regenerate the four self-contained runtime scripts (page hook, content script, popup, background worker) and the generated model tables; there is no separate core script dependency at runtime. Node.js 18+ is needed only for development checks; Python 3 is needed only for packaging or the local preview. Type checking needs `npm ci` once for the pinned TypeScript and Chrome type definitions; nothing from `node_modules` is shipped.

```sh
npm run build
npm run check
npm test
npm run typecheck
npm run package
python3 -m http.server 8765 --bind 127.0.0.1
```

For the last command, open `http://127.0.0.1:8765/tests/preview/index.html` and click **Run UI checks**. The preview uses the actual extension UI and hook with a simulated Chrome API and WebSocket transport. Its test doubles are excluded from the extension package. The package command writes `outputs/M365-Copilot-Extender-v<version>.zip` with `manifest.json` in the ZIP root.

GitHub Actions runs `npm run check`, `npm test` and `npm run typecheck` on Node 20, 22 and 24 for every pull request and push to `main` ([workflow](.github/workflows/ci.yml)). The browser preview checks are still run by hand.

Automated coverage includes multi-record framing, all six models, tone casing, native socket behavior, settings-load races, cross-tab propagation of chat choices, restart-style reloads, absent/zero quotas, final-item refusals, stale socket isolation, per-conversation models and readings, hashed conversation storage, badge thresholds, and metadata privacy. Persistence tests cover allowances shared across tabs and per-chat readings restored in a new tab. The browser fixture checks the model rows, the GPT-6 Sol default, switching a chat's model, the status line, usage states and colours, saved allowances and readings, and absence of injected page controls.

**Live account validation remains required.** These checks establish extension behavior against fixtures; they do not establish current Microsoft availability, backend identity, or account-specific quota policies. No Microsoft account credentials are needed for the test suite.

## Technical references

- [cramt/m365-copilot-proxy](https://github.com/cramt/m365-copilot-proxy): protocol, model tones, routing, response metadata, and quota research.
- [aweussom/i-want-my-model-picker-back](https://github.com/aweussom/i-want-my-model-picker-back): picker research and WebSocket interception experiments.
- [Research baseline](docs/research.md): source revisions, findings, and outstanding live checks.

No upstream implementation has been incorporated.

## Troubleshooting startup errors

If the badge or model switching doesn't work after updating, reload the extension at `chrome://extensions` and then reload every open Copilot tab: the page hook only attaches when the page loads.

The runtime scripts are self-contained, so there is no shared script that must load first. This removed the dependency implicated in an earlier reported `Cannot read properties of undefined (reading DEFAULT_TONE)` error; the exact cause in that installation was never confirmed. Startup tests run the packaged scripts without a global core. If you still see errors after updating, clear old errors on the extension's `chrome://extensions` card before sending one new message so stale errors are not confused with current ones, and disable any other model-picker extension while testing.
