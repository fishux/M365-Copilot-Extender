// Type declarations for `npm run typecheck`. Not shipped: the build inlines the
// runtime files and these globals become private bundle variables.

interface Model {
  readonly tone: string;
  readonly label: string;
  readonly menuLabel: string;
  readonly badge: string;
  readonly budgetPrefix: string | null;
  readonly knownPaidCap: boolean;
  readonly budgetEvidence: string;
}

type Scenario = "OfficeWebPaidCopilot" | "OfficeWebIncludedCopilot" | "Unknown";
type AllowanceLevel = "normal" | "warn" | "critical";

/** Saved per-conversation choices, keyed by a hash of the ConversationId. */
type ConversationTones = Record<string, { tone: string; at: number }>;

interface ConnectionInfo { scenario: Scenario; conversationId: string | null }

interface SocketState extends ConnectionInfo {
  socketId: number;
  tone: string | null;
  queue: unknown[];
  timer: ReturnType<typeof setTimeout> | null;
}

interface CopilotMetadata {
  final: boolean;
  expiresAt?: string;
  metering?: Record<string, number>;
  conversation?: { used: number; limit: number };
  result?: string;
  origin?: string;
}

/** Readings for one conversation, as kept by the isolated script. */
interface Observed {
  appliedTone?: string;
  result?: string;
  origin?: string;
  conversation?: { used: number; limit: number };
  expiresAt?: string;
  warning?: string;
}

interface ConversationView {
  capacity: string;
  capacityTitle: string;
  used: number | null;
  limit: number | null;
  percent: number | null;
  nearLimit: boolean;
  expiry: string | null;
  expiryText: string;
  expiryTitle: string;
}

interface Core {
  MODELS: readonly Model[];
  DEFAULT_TONE: string;
  ALLOWANCE_LEVELS: Readonly<{ warn: number; critical: number }>;
  modelFor(tone: unknown): Model | undefined;
  connectionInfo(rawURL: string): ConnectionInfo | null;
  safeScenario(value: unknown): Scenario;
  safeConversationId(value: unknown): string | null;
  conversationFromPath(pathname: unknown): string | null;
  allowanceLevel(remaining: number | null): AllowanceLevel;
  lowestAllowance(tone: unknown, metering: Record<string, number> | undefined): number | null;
  isChat(frame: unknown): boolean;
  parseRecords(data: unknown): any[];
  rewrite<T>(data: T, tone: unknown): { data: T | string; tones: string[]; chats: number };
  parseMetering(raw: unknown): Record<string, number>;
  allowanceFor(tone: unknown, metering: Record<string, number> | undefined): { daily: number | null; weekly: number | null } | null;
  readMetadata(frame: any): CopilotMetadata[];
  count(value: unknown): number | null;
  parseExpiry(value: unknown): string | null;
  conversationView(conversation: unknown, expiresAt: unknown, now?: number): ConversationView;
  allowedOrigins: readonly string[];
  allowedResults: readonly string[];
}

interface PopupState {
  tone: string;
  loaded: boolean;
  saving: boolean;
  connected: boolean;
  hookReady: boolean;
  error: string;
  observed: Observed & { metering?: Record<string, number>; quotaAt?: number; scenario?: string };
  conversation: { key: string } | null;
  conversationTone: string | null;
}

interface Panel {
  mount(host: HTMLElement, handlers: { onSelect(tone: string): void; onConversationSelect(tone: string | null): void }): { update(state: PopupState): void };
}

declare var M365ExtenderModels: Model[];
declare var M365ExtenderCore: Core;
declare var M365ExtenderPanel: Panel;
