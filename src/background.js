// @ts-check
/* Background service worker: toolbar badge, title and icon colour only.
   It makes no network requests and reads no page content. */
(() => {
  "use strict";
  const core = globalThis.M365ExtenderCore;
  const TITLE = "M365 Copilot Extender";
  const BADGE_COLORS = { normal: "#6254d9", warn: "#e07b16", critical: "#d93025" };
  // Same artwork as icons/source.svg, recoloured when an allowance runs low.
  const GRADIENTS = { warn: ["#f7a94f", "#d9690f"], critical: ["#f2665c", "#c5221f"] };
  const DEFAULT_ICON = { 16: "icons/icon16.png", 32: "icons/icon32.png" };
  /** @type {Partial<Record<"warn" | "critical", Record<number, ImageData>>>} */
  const icons = {};
  const defaultTone = core.DEFAULT_TONE;

  /** @param {"warn" | "critical"} level @param {number} size */
  function drawIcon(level, size) {
    const canvas = new OffscreenCanvas(size, size);
    const context = /** @type {OffscreenCanvasRenderingContext2D} */ (canvas.getContext("2d"));
    context.scale(size / 128, size / 128);
    const gradient = context.createLinearGradient(4, 4, 124, 124);
    gradient.addColorStop(0, GRADIENTS[level][0]);
    gradient.addColorStop(1, GRADIENTS[level][1]);
    context.fillStyle = gradient;
    context.beginPath();
    context.roundRect(4, 4, 120, 120, 32);
    context.fill();
    context.strokeStyle = "#fff";
    context.lineWidth = 6;
    context.lineCap = "round";
    context.stroke(new Path2D("M64 28v72M28 64h72M38.5 38.5l51 51M38.5 89.5l51-51"));
    context.fillStyle = "#fff";
    context.beginPath();
    context.arc(64, 64, 16, 0, Math.PI * 2);
    context.fill();
    return context.getImageData(0, 0, size, size);
  }

  /** @param {"warn" | "critical"} level */
  const iconFor = level => (icons[level] ??= { 16: drawIcon(level, 16), 32: drawIcon(level, 32) });

  /** @param {string} tone @param {number | null} remaining */
  function badgeView(tone, remaining) {
    const model = core.modelFor(tone) || core.modelFor(core.DEFAULT_TONE);
    const level = core.allowanceLevel(remaining);
    const left = remaining === null ? "" : ` · ${remaining} message${remaining === 1 ? "" : "s"} left`;
    return { text: model?.badge ?? "", color: BADGE_COLORS[level], level, title: `${TITLE} — ${model?.label ?? "Unknown"}${left}` };
  }

  async function applyDefault() {
    const view = badgeView(defaultTone, null);
    await Promise.all([
      chrome.action.setBadgeText({ text: view.text }),
      chrome.action.setBadgeBackgroundColor({ color: view.color }),
      chrome.action.setBadgeTextColor({ color: "#ffffff" }),
      chrome.action.setTitle({ title: `${TITLE} — Default: ${core.modelFor(defaultTone)?.label}` })
    ]);
  }

  /** @param {number} tabId @param {string} tone @param {number | null} remaining */
  async function applyTab(tabId, tone, remaining) {
    const view = badgeView(tone, remaining);
    await Promise.all([
      chrome.action.setBadgeText({ tabId, text: view.text }),
      chrome.action.setBadgeBackgroundColor({ tabId, color: view.color }),
      chrome.action.setTitle({ tabId, title: view.title }),
      view.level === "normal" ? chrome.action.setIcon({ tabId, path: DEFAULT_ICON }) : chrome.action.setIcon({ tabId, imageData: iconFor(view.level) })
    ]);
  }

  /** Leaving Copilot (or reloading) returns the tab to the default badge. */
  async function resetTab(tabId) {
    await Promise.all([
      chrome.action.setBadgeText({ tabId, text: /** @type {any} */ (null) }),
      chrome.action.setBadgeBackgroundColor({ tabId, color: BADGE_COLORS.normal }),
      chrome.action.setTitle({ tabId, title: `${TITLE} — Default: ${core.modelFor(defaultTone)?.label}` }),
      chrome.action.setIcon({ tabId, path: DEFAULT_ICON })
    ]);
  }

  chrome.runtime.onMessage.addListener((message, sender) => {
    const tabId = sender.tab?.id;
    if (sender.id !== chrome.runtime.id || message?.type !== "M365_EXTENDER_TAB" || typeof tabId !== "number") return;
    const remaining = core.count(message.remaining);
    applyTab(tabId, message.tone, remaining).catch(() => { /* Tab closed. */ });
  });
  chrome.tabs.onUpdated.addListener((tabId, info) => {
    if (info.status === "loading") resetTab(tabId).catch(() => { /* Tab closed. */ });
  });
  // The default model is fixed in models.json; per-chat selections arrive from tabs.
  applyDefault().catch(() => {});
})();
