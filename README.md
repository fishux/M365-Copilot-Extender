# M365 Copilot Extender

A lightweight Chrome Manifest V3 extension for persistent model selection in [Microsoft 365 Copilot](https://m365.cloud.microsoft).

Choose a default model once using the Chrome toolbar popup. Your choice applies to outgoing chat requests across Copilot tabs and survives browser restarts. GPT-6.1 Sol is the initial default. You can also pin a different model to an individual conversation.

The toolbar icon shows which model the current tab is using and turns orange, then red, as that model's remaining allowance runs low. The popup lists every model on one screen in a compact console style: click a row to set the default, or its pin to use that model for the current chat only. It follows the system light/dark appearance, and connection details sit in an expandable drawer. Nothing is displayed over the Copilot page. See the [changelog](CHANGELOG.md) for what's new.

## Install on Windows

1. Download this repository using **Code → Download ZIP**, then extract it to a permanent folder. Alternatively, extract a packaged extension ZIP.
2. In Google Chrome, open `chrome://extensions` and enable **Developer mode**.
3. Click **Load unpacked**. Select this repository's `extension` folder, or the extracted package folder containing `manifest.json`.
4. Pin **M365 Copilot Extender** from Chrome's Extensions menu.
5. Reload any already-open Copilot tabs. Choose a model in the toolbar popup and send a message.

When updating, replace the extension files, click **Reload** on its `chrome://extensions` card, then reload Copilot. Keep the unpacked folder in place so Chrome can load it on restart. Settings belong to this Chrome profile; they do not synchronize across separate profiles or devices.

## Models

<!-- models:start (generated from src/models.json by npm run build) -->
| Selection | Badge | Requested tone |
| --- | --- | --- |
| GPT-6.1 Sol | `6.1` | `Gpt_61_Sol_Reasoning` |
| GPT-6 Sol | `6S` | `Gpt_6_Sol_Reasoning` |
| GPT-6 · Think deeper | `6` | `Gpt_6_Reasoning` |
| GPT-5.6 · Think deeper | `5.6` | `Gpt_5_6_Reasoning` |
| GPT-5.5 · Think deeper | `5.5` | `Gpt_5_5_Reasoning` |
| Auto | `A` | `Magic` |
<!-- models:end -->

Model matching ignores case: Copilot's web client sends `Magic` for Auto, other clients `magic`, and the service accepts both.

### Adding or removing a model

Edit [`src/models.json`](src/models.json) — the first entry is the initial default — then run `npm run build`. The build inlines the list into every runtime script and regenerates the tables above and in the [architecture notes](docs/architecture.md). Each entry needs a `tone` (the exact value Copilot expects), `label`, `menuLabel` (popup text), `badge` (toolbar text, at most 4 characters), `budgetPrefix` (a metering key prefix such as `GPT61SolQuery`, or `null`), `knownPaidCap` (show a waiting usage box before the first reading), and `budgetEvidence` (why that prefix is believed). Saved choices for a removed model fall back to the default.

The extension changes only `arguments[0].tone` on type `1`/`4` chat invocations to Copilot's Chathub socket. It preserves other SignalR records, request fields, and the connection's existing paid/included scenario. It does not add license entitlements, change routing scenarios, or automatically switch models when a budget is exhausted.

Microsoft controls final model routing. An applied tone and response `contentOrigin` do **not** independently verify the exact backend model. Try a new conversation if changing a model in an existing conversation is refused or behaves unexpectedly.

## Usage allowances

Counters show **remaining messages** reported by Copilot, refreshed after a response. They do not estimate usage or assume that the reported 40/day and 75/week apply to every account.

- GPT-6.1 Sol uses the documented `throttling.metering.GPT61SolQueryDaily.remainingAllowance` and `GPT61SolQueryWeekly.remainingAllowance` fields. The parser reads both final type `2` items and type `1` update records.
- Other selectable models have no hardcoded quota. If Copilot supplies directly matching model-specific daily/weekly keys, the UI can display them; the mappings and their evidence limits are described in [architecture notes](docs/architecture.md).
- Models without a documented or observed budget have no usage box. GPT-6.1 Sol shows a waiting state before its first paid-route reading. Missing data is unknown, never zero or unlimited. Included-route GPT-6.1 Sol has no counter unless an allowance is actually reported.
- Zero remains visible. `OutOfCredits` is surfaced even if the refusal appears only in the final result. There is no guessed reset countdown or automatic fallback.
- Allowances are per account, so they stay visible when you switch to another conversation in the same tab.
- Readings are kept **in memory in each tab**. The popup reads the active Copilot tab; opening it on another website still lets you change preferences, but shows no connection data. Send a message in a new or reloaded Copilot tab to obtain its readings. Old readings are marked **Last known** after five minutes; they refresh only with a new response.

The status line at the top of the popup shows whether your preference was applied to the last request, or the last error. The **details** drawer shows the observed scenario, the last requested model, result code, and known content origin, each with a copy button; **copy all** copies only those fields, never conversation IDs. If a request reaches a usage limit, the popup offers one-click buttons to pin another model to the chat; it never switches automatically.

## Conversation capacity and expiry

The toolbar popup includes a compact conversation section below model usage:

- **Capacity:** the server's used/maximum message counts and a slim meter. It highlights at 90% capacity. The maximum is read from Copilot, never assumed to be 600, and is separate from model allowances. Hover the count for remaining conversation messages.
- **Server expiry:** a short local date and time remaining, read from `item.conversationExpiryTime`. Hover or keyboard-focus the date for its exact local timestamp. A passed timestamp is labelled **Passed**; it is not a claim that the chat has been deleted. The server timestamp does not establish Microsoft's retention policy.

These readings belong to the conversation shown in the page. Each conversation keeps its own last observed reply in this tab, so switching chats in the sidebar shows that chat's readings, and a new, unsent chat shows none. They are not a live query: a conversation that has not replied in this tab shows **Awaiting reply** for capacity and **Not reported** for expiry. The last valid values remain visible until Copilot reports replacements, including across partial or final replies that omit fields. Opening a background socket does not clear them, and replies on a connection where nothing was sent are ignored. Conversation readings are never saved to disk. Countdown labels refresh locally without Microsoft requests.

## Per-conversation model

Each model row has a pin button that pins that model to the chat you're viewing. Click the filled pin again, or **unpin** in the chat section, to return the chat to the default. The choice applies to every later turn in that conversation, in any tab, and survives restarts; other conversations keep using the default.

- Pins become available once the conversation has an ID: after its first message is sent, or when you open an existing chat from the sidebar.
- Conversations are identified by the `ConversationId` that Copilot's own client attaches to its Chathub connection (and shows in `/chat/conversation/<id>` URLs). Only a SHA-256 hash of that ID is saved, together with the model and a timestamp; the 200 most recent choices are kept.
- Before the first chat request on a connection, the hook briefly waits for that conversation's choice, as it already does for the default. If it never arrives within 1.5 seconds, the default model is used and a warning is shown.
- Copilot's own picker also stores a model per conversation on the server. The extension overrides what is sent; it does not change what Copilot's picker displays.

## Toolbar badge

The toolbar badge shows the model a Copilot tab is using (see the **Badge** column above). On other tabs it shows the default model. Hover the icon for the model's name and its lowest reported remaining allowance.

When the active model's lowest reported daily or weekly allowance is **20 or fewer**, the icon and badge turn **orange**; at **10 or fewer** they turn **red**. Popup counters use the same colours. Models with no reported allowance keep the normal colour.

## Privacy and permissions

The only extension API permission is `storage`, with content scripts scoped to `https://m365.cloud.microsoft/*`. There are no host permissions, third-party requests, analytics, credentials, remote scripts, or runtime dependencies. A background service worker only sets the toolbar badge, title and icon; it makes no network requests. `chrome.storage.local` holds the default model and per-conversation choices keyed by hashed conversation IDs. The page hook parses chat frames locally, but forwards only allowlisted metadata, numeric quota values and the conversation ID; it does not forward or save prompts, replies, access tokens, full socket URLs, or account identifiers.

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

Automated coverage includes multi-record framing, all six models, tone casing, native socket behavior, settings-load races, cross-tab preference propagation, restart-style reloads, absent/zero quotas, final-item refusals, stale socket isolation, per-conversation models and readings, hashed conversation storage, badge thresholds, and metadata privacy. The browser fixture checks the model rows and pins, usage states and colours, the conversation override, active-tab readings, and absence of injected page controls.

**Live account validation remains required.** These checks establish extension behavior against fixtures; they do not establish current Microsoft availability, backend identity, or account-specific quota policies. No Microsoft account credentials are needed for the test suite.

## Technical references

- [cramt/m365-copilot-proxy](https://github.com/cramt/m365-copilot-proxy): protocol, model tones, routing, response metadata, and quota research.
- [aweussom/i-want-my-model-picker-back](https://github.com/aweussom/i-want-my-model-picker-back): picker research and WebSocket interception experiments.
- [Research baseline](docs/research.md): source revisions, findings, and outstanding live checks.

No upstream implementation has been incorporated.

## Troubleshooting startup errors

If the badge or per-conversation pins don't appear after updating, reload the extension at `chrome://extensions` and then reload every open Copilot tab: the page hook only attaches when the page loads.

Since v0.1.7 the runtime scripts are self-contained. This removes the separate shared-core loading dependency implicated by a reported `Cannot read properties of undefined (reading DEFAULT_TONE)` error. The exact reason the old shared script was unavailable in that installation remains unconfirmed. New startup tests execute the actual packaged scripts without a global core. Live validation remains necessary.

To try this update, extract it into a new folder, point the Chrome extension at that folder (or replace every file in its current folder), reload the extension at `chrome://extensions`, then reload the Copilot tab. Clear old extension errors before sending one new message so stale errors are not confused with current ones. Disable any older model-picker extension during the test to keep the comparison controlled.
