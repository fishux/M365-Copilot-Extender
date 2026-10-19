# M365 Copilot Extender

A lightweight Chrome Manifest V3 extension for persistent model selection in [Microsoft 365 Copilot](https://m365.cloud.microsoft).

Choose a model once using the Chrome toolbar popup or the floating Copilot panel. Your choice applies to outgoing chat requests across Copilot tabs and survives browser restarts. GPT-6.1 Sol is the initial default.

The popup includes a **Show panel on Copilot** switch. Hiding the panel leaves the override active. Both surfaces support system light/dark appearance, keyboard navigation, and always-visible connection details.

## Install on Windows

1. Download this repository using **Code → Download ZIP**, then extract it to a permanent folder. Alternatively, extract a packaged extension ZIP.
2. In Google Chrome, open `chrome://extensions` and enable **Developer mode**.
3. Click **Load unpacked**. Select this repository's `extension` folder, or the extracted package folder containing `manifest.json`.
4. Pin **M365 Copilot Extender** from Chrome's Extensions menu.
5. Reload any already-open Copilot tabs. Choose a model in either selector and send a message.

When updating, replace the extension files, click **Reload** on its `chrome://extensions` card, then reload Copilot. Keep the unpacked folder in place so Chrome can load it on restart. Settings belong to this Chrome profile; they do not synchronize across separate profiles or devices.

## Models

| Selection | Requested tone |
| --- | --- |
| GPT-6.1 Sol | `Gpt_61_Sol_Reasoning` |
| GPT-6 Sol | `Gpt_6_Sol_Reasoning` |
| GPT-6 · Think deeper | `Gpt_6_Reasoning` |
| GPT-5.6 · Think deeper | `Gpt_5_6_Reasoning` |
| GPT-5.5 · Think deeper | `Gpt_5_5_Reasoning` |
| Auto | `Magic` |

The extension changes only `arguments[0].tone` on type `1`/`4` chat invocations to Copilot's Chathub socket. It preserves other SignalR records, request fields, and the connection's existing paid/included scenario. It does not add license entitlements, change routing scenarios, or automatically switch models when a budget is exhausted.

Microsoft controls final model routing. An applied tone and response `contentOrigin` do **not** independently verify the exact backend model. Try a new conversation if changing a model in an existing conversation is refused or behaves unexpectedly.

## Usage allowances

Counters show **remaining messages** reported by Copilot, refreshed after a response. They do not estimate usage or assume that the reported 40/day and 75/week apply to every account.

- GPT-6.1 Sol uses the documented `throttling.metering.GPT61SolQueryDaily.remainingAllowance` and `GPT61SolQueryWeekly.remainingAllowance` fields. The parser reads both final type `2` items and type `1` update records.
- Other selectable models have no hardcoded quota. If Copilot supplies directly matching model-specific daily/weekly keys, the UI can display them; the mappings and their evidence limits are described in [architecture notes](docs/architecture.md).
- Models without a documented or observed budget have no usage box. GPT-6.1 Sol shows a waiting state before its first paid-route reading. Missing data is unknown, never zero or unlimited. Included-route GPT-6.1 Sol has no counter unless an allowance is actually reported.
- Zero remains visible. `OutOfCredits` is surfaced even if the refusal appears only in the final result. There is no guessed reset countdown or automatic fallback.
- Readings are kept **in memory in each tab**. The popup reads the active Copilot tab; opening it on another website still lets you change preferences, but shows no connection data. Send a message in a new or reloaded Copilot tab to obtain its readings. Old readings are marked **Last known** after five minutes; they refresh only with a new response.

Connection details show the observed scenario, the last requested model, result code, and known content origin.

## Conversation capacity and expiry

Both interfaces include a compact conversation section below model usage:

- **Capacity:** the server's used/maximum message counts and a slim meter. It highlights at 90% capacity. The maximum is read from Copilot, never assumed to be 600, and is separate from model allowances. Hover the count for remaining conversation messages.
- **Server expiry:** a short local date and time remaining, read from `item.conversationExpiryTime`. Hover or keyboard-focus the date for its exact local timestamp. A passed timestamp is labelled **Passed**; it is not a claim that the chat has been deleted. The server timestamp does not establish Microsoft's retention policy.

These are readings from the last observed reply in this tab, not a live query of whichever history item is open. Before a response, missing capacity shows **Awaiting reply** and missing or invalid expiry shows **Not reported**. Final replies with missing fields clear old values. New socket connections also clear the previous readings. No per-conversation data is saved to disk. Countdown labels refresh locally without Microsoft requests.

## Privacy and permissions

The only extension API permission is `storage`, with content scripts scoped to `https://m365.cloud.microsoft/*`. There are no third-party requests, analytics, credentials, remote scripts, service workers, or runtime dependencies. Only the selected model and panel visibility are saved in `chrome.storage.local`. The page hook parses chat frames locally, but forwards only allowlisted metadata and numeric quota values; it does not forward or save prompts, replies, access tokens, full socket URLs, or account identifiers.

The page bridge is observable by scripts on the Copilot page. It is an instrumentation channel, not an authenticated model-identity or security boundary. See [architecture and limitations](docs/architecture.md).

## Development and validation

No build or dependency installation is required to load `extension/` in Chrome. Node.js 18+ is needed only for development checks; Python 3 is needed only for packaging or the local preview.

```sh
npm run check
npm test
npm run package
python3 -m http.server 8765 --bind 127.0.0.1
```

For the last command, open `http://127.0.0.1:8765/tests/preview/index.html` and click **Run UI checks**. The preview uses the actual extension UI and hook with a simulated Chrome API and WebSocket transport. Its test doubles are excluded from the extension package. The package command writes `outputs/M365-Copilot-Extender-v0.1.4.zip` with `manifest.json` in the ZIP root.

Automated coverage includes multi-record framing, all six models, native socket behavior, settings-load races, cross-tab preference and visibility propagation, restart-style reloads, absent/zero quotas, final-item refusals, stale socket isolation, and metadata privacy. The browser fixture checks both selectors, panel visibility, usage states, collapse behavior, and active-tab popup readings.

**Live account validation remains required.** These checks establish extension behavior against fixtures; they do not establish current Microsoft availability, backend identity, or account-specific quota policies. No Microsoft account credentials are needed for the test suite.

## Technical references

- [cramt/m365-copilot-proxy](https://github.com/cramt/m365-copilot-proxy): protocol, model tones, routing, response metadata, and quota research.
- [aweussom/i-want-my-model-picker-back](https://github.com/aweussom/i-want-my-model-picker-back): picker research and WebSocket interception experiments.
- [Research baseline](docs/research.md): source revisions, findings, and outstanding live checks.

No upstream implementation has been incorporated.
