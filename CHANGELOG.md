# Changelog

All notable changes to M365 Copilot Extender. Versions match `extension/manifest.json`.

## 0.2.0 — 2026-10-25

### Added
- **Per-conversation model.** A "This conversation" selector in the popup overrides the default model for one chat. The choice persists across tabs and restarts. Conversations are identified by the `ConversationId` Copilot already puts on its Chathub connection; only a SHA-256 hash of it is saved, and the 200 most recent choices are kept.
- **Toolbar badge.** The toolbar icon shows the model a tab is using (`6.1`, `6S`, `6`, `5.6`, `5.5`, `A`). Its tooltip shows the model and remaining allowance.
- **Low-allowance colours.** The toolbar icon and badge turn orange at 20 or fewer remaining messages and red at 10 or fewer, based on the lowest reported daily or weekly allowance for the tab's model. Popup counters use the same colours.
- Background service worker (`background.js`) for the badge and icon. It makes no network requests; no new permissions.
- `src/models.json` as the single list of models. `npm run build` inlines it into every runtime script and regenerates the model tables in `README.md` and `docs/architecture.md`.
- `npm run typecheck`: JSDoc and `// @ts-check` type checking of `src/` with strict null checks (TypeScript and Chrome types as pinned dev dependencies only).
- `CHANGELOG.md`.
- GitHub Actions CI: `npm run check`, `npm test` and `npm run typecheck` on Node 20, 22 and 24 for pull requests and pushes to `main`.

### Changed
- Conversation readings (capacity, expiry, result, origin, last request) are tracked per conversation and follow the conversation shown in the page, including when switching chats in the sidebar. Model allowances are account-wide and stay visible across conversations.
- Every warning from the page hook carries its socket and conversation. Warnings about a just-sent chat make that conversation current.
- Model matching ignores case. Copilot's web client sends `Magic` for Auto and other clients send `magic`; an equivalent tone is left byte-identical.
- GPT-6 no longer has a candidate budget prefix: upstream research found it has no priority-access allowance.
- The popup's first selector is labelled "Default model".

### Maintenance
- The page hook parses each outgoing payload once instead of up to three times.
- Shared scenario allowlist; removed dead popup CSS and branches.
- `npm run check` verifies that the `manifest.json`, `package.json` and `INSTALL.txt` versions match.
- Fixed preview fixture paths that escaped the repository root.

## 0.1.7 — 2026-10-20
- Removed the in-page panel; all controls live in the toolbar popup.
- Restored self-contained runtime scripts, fixing a `DEFAULT_TONE` startup error.

## 0.1.6 — 2026-10-20
- Copilot metadata stays visible until fresh readings arrive.

## 0.1.5 — 2026-10-20
- Self-contained runtime scripts with no shared-core script dependency.

## 0.1.4 — 2026-10-19
- Conversation capacity meter and server expiry in a compact details section.

## 0.1.1 — 2026-10-19
- Consolidated the shared renderer and inlined popup styles.

## 0.1.0 — 2026-10-19
- First MVP: persistent model selection for Copilot's Chathub requests and server-reported usage allowances.
