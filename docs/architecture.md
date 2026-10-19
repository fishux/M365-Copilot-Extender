# MVP architecture and validation boundaries

## Components

`manifest.json` installs two static `document_start` content-script groups on `https://m365.cloud.microsoft/*`. `core.js` is loaded independently in each world. `page-hook.js` runs in MAIN so it can intercept the page's WebSockets; `content.js` and `panel.js` run in ISOLATED so they can use extension storage. Chrome's toolbar popup uses `popup.js` and the same panel renderer.

The hook proxies the WebSocket constructor to attach passive response listeners early. It also wraps the native prototype's `send`, checking only `wss://substrate.office.com/m365Copilot/Chathub/…` sockets. It preserves the native constructor's prototype, static constants, `instanceof`, and subclass behavior. Unrelated sockets pass through. JSON payloads are split/rejoined on `\x1e`; only type `1`/`4`, target `chat`, with an object at `arguments[0]` are modified. Every qualifying request uses the current preference, including follow-up turns.

## Startup and persistence

`selectedTone` and `inPageEnabled` are separate `chrome.storage.local` keys. Separate writes avoid replacing unrelated preferences when two tabs change settings. `chrome.storage.onChanged` updates both UIs and the hook. Per-key revision checks prevent a stale asynchronous initial read from overwriting a newer cross-tab update.

Because MAIN cannot read extension storage, the hook briefly holds the first chat request while the isolated script loads preferences. Later sends on the same socket retain their order behind it. If settings never arrive, the queue flushes after 1.5 seconds using Copilot's original tone and emits a visible warning; it never silently chooses a different forced model. Handshakes and pre-chat pings pass immediately. Socket close drops queued requests; native CONNECTING errors are preserved. No pending request is retried automatically.

## Quota metadata

The parser reads `frame.item` for type `2` final results and each `frame.arguments` value for type `1`, target `update`. It extracts only valid nonnegative integer `remainingAllowance` values, known result/origin enums, and numeric conversation counts. Final responses replace the metering snapshot, including clearing it when absent; streaming updates without a metering map do not erase an earlier snapshot.

Only `GPT61SolQueryDaily` and `GPT61SolQueryWeekly` have direct evidence in the reviewed references for the six offered models. That model is marked as having a known paid-route cap. The other tone mappings conservatively recognize the exact `Daily`/`Weekly` suffixes of the following model-specific prefixes **only if those keys are actually observed**:

| Selected model | Candidate budget prefix | Evidence |
| --- | --- | --- |
| GPT-6.1 Sol | `GPT61SolQuery` | Documented by upstream live probes |
| GPT-6 Sol | `GPT6SolQuery` | Name-based association only; not verified by upstream research |
| GPT-6 | `GPT6Query` | Name-based association only; not verified by upstream research |
| GPT-5.6 | `GPT56Query` | Name-based association only; not verified by upstream research |
| GPT-5.5 | `GPT55Query` | Name-based association only; not verified by upstream research |
| Auto | None | Cannot attribute a budget to Auto's unknown backend choice |

The comparison ignores underscore differences and letter case. Unrecognized keys do not become model counters. No allowance is inferred from another model's map, no limits are hardcoded, and a missing value never becomes zero. New key formats require further research rather than guesswork. This limits automatic support to the budget shapes the extension can associate with a selected model; it cannot guarantee discovery of every future Microsoft usage cap.

Quotas are not persisted or shared between tabs, so accounts in different tabs are not mixed. Each new socket clears the prior snapshot; sequential local socket IDs let the isolated script ignore late results from an older connection. No IDs from the actual connection URL are forwarded. These IDs are bookkeeping only, not account identity. The popup polls only the active tab's content script while it is open; this makes no Microsoft requests.

## Bridge and UI

Both worlds communicate via same-window, same-origin `postMessage` with an explicit channel and direction. The isolated script validates and filters incoming values. Page messages cannot directly write persistent preferences: only the popup and in-page control handlers do so. MAIN uses a model allowlist for configuration. Host-page scripts can observe or spoof this bridge, as they can alter Copilot's own JavaScript, so it does not establish trusted model identity. Only sanitized metadata crosses it. Values rendered from settings or responses use `textContent`; the only HTML template is a static local UI string.

The in-page panel uses a shadow root and fixed positioning, with a collapse button. A saved popup switch hides it without removing the hook. The popup can change preferences even outside Copilot; tab metadata is available only when its content script is reachable. Neither UI claims that `DeepLeo` identifies a precise model version.

## Current limits and live acceptance checks

The MVP covers the top-level Copilot page's JSON Chathub WebSockets on the observed host. It does not intercept workers, cross-origin iframes, other Copilot domains, future transports, or MessagePack/binary chat frames. Reload the page after installation or extension updates. Site changes, tenant policies, agent routing, and Microsoft entitlement checks may affect behavior.

The Node suite exercises actual protocol and controller scripts in isolated VM contexts with mocked transport/storage. The browser fixture exercises the actual UI, popup controller, page hook, and isolated controller with simulated APIs in regular web frames. Neither test installs the MV3 extension into the user's real Chrome profile. A fresh VM with shared saved settings is used for restart-style tests; no physical browser restart is claimed.

Before declaring live support, load the unpacked extension on Windows Chrome, test all six selections in new and existing chats, and confirm that switching tabs and restarting Chrome preserves both preferences. Check that the observed applied tone matches outgoing DevTools frames and that usage values match sanitized type `2` final-item metering. Verify zero/absence handling and final-item refusal status. Do not publish raw frames or account-bearing socket URLs. Live backend identity remains unresolved even when a request succeeds.

References: [upstream research revisions](research.md), [Chrome content-script worlds and timing](https://developer.chrome.com/docs/extensions/develop/concepts/content-scripts), [storage persistence and events](https://developer.chrome.com/docs/extensions/reference/api/storage), and [tabs API permission boundaries](https://developer.chrome.com/docs/extensions/reference/api/tabs).
