# Changelog

All notable changes to M365 Copilot Extender. Versions match `extension/manifest.json`.

## 0.3.1 — 2026-10-25

- Refreshed the install notes (`INSTALL.txt`), README, architecture notes and manifest description, which still described the old default selector, pins and GPT-6.1 Sol default.
- The release now ships an extension-only ZIP (`manifest.json` at its root) alongside GitHub's source archive.

## 0.3.0 — 2026-10-25

- **GPT-6 Sol is the fixed default.** Every chat starts on it. The saved default and the default selector are retired.
- **A chat's model choice sticks.** Click a model in the popup and that chat uses it from then on, in every tab and across restarts. Clicking the default returns the chat to it. Up to 2,000 chats keep their choice (was 200).
- **Console-style popup.** All six models on one screen with a status line naming the model the last request was sent via, a details drawer for connection info, and light and dark themes that follow the system.
- **Readings persist.** Each chat's last readings (model, result, origin, capacity, expiry) and the last reported allowances are saved, shared across tabs and shown after a restart, including outside Copilot.

## 0.2.0 — 2026-10-25

- **Per-conversation model.** Override the default for a single chat. Only a SHA-256 hash of the `ConversationId` is saved.
- **Toolbar badge** shows the model a tab is using, with a tooltip of the model and its remaining allowance. The icon and badge turn orange at 20 or fewer remaining messages and red at 10 or fewer.
- Conversation capacity, expiry and other readings are tracked per conversation and follow the chat shown in the page. Allowances stay account-wide.
- Model matching ignores case, so Copilot's `Magic` and other clients' `magic` both work.
- `src/models.json` is the single model list; `npm run build` inlines it and regenerates the README and architecture tables.
- Added `npm run typecheck` and GitHub Actions CI (check, test and typecheck on Node 20, 22 and 24).

## 0.1.x — 2026-10-19 to 2026-10-20

- First MVP (0.1.0): persistent model selection for Copilot's Chathub requests, with server-reported usage allowances.
- Conversation capacity meter and server expiry in the popup (0.1.4).
- In-page panel removed; all controls live in the toolbar popup (0.1.7).
- Self-contained runtime scripts with no shared-core dependency, fixing a `DEFAULT_TONE` startup error (0.1.5, 0.1.7).
- Copilot readings stay visible until fresh ones arrive (0.1.6).
