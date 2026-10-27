// Type declarations for `npm run typecheck`. Not shipped: the build inlines the
// runtime files and these globals become private bundle variables.

interface Model {
  readonly tone: string;
  /** Tone sent when Work IQ is on, or null when the model has none (the default is used). */
  readonly workIQTone: string | null;
  readonly default?: boolean;
  readonly label: string;
  readonly menuLabel: string;
  readonly badge: string;
  readonly budgetPrefix: string | null;
  readonly knownPaidCap: boolean;
  readonly budgetEvidence: string;
}

type Scenario = "OfficeWebPaidCopilot" | "OfficeWebIncludedCopilot" | "officeweb" | "Unknown";
type AllowanceLevel = "normal" | "warn" | "critical";

/** Saved per-conversation choices, keyed by a hash of the ConversationId. */
type ConversationTones = Record<string, { tone: string; at: number }>;

interface ConnectionInfo { scenario: Scenario; conversationId: string | null; workIQ: boolean }

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
  /** Whether the last request in this conversation used Work IQ. */
  workIQ?: boolean;
  result?: string;
  origin?: string;
  conversation?: { used: number; limit: number };
  expiresAt?: string;
  warning?: string;
  /** When this reading last changed (ms). Newer saved or live readings win. */
  at?: number;
}

/** Saved per-conversation readings, keyed by a hash of the ConversationId. */
type ConversationReadings = Record<string, Observed & { at: number }>;

/** Last reported allowances, shared by every tab. */
interface Usage { metering: Record<string, number>; quotaAt: number }

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
  MAX_SAVED_CONVERSATIONS: number;
  ALLOWANCE_LEVELS: Readonly<{ warn: number; critical: number }>;
  modelFor(tone: unknown): Model | undefined;
  connectionInfo(rawURL: string): ConnectionInfo | null;
  safeScenario(value: unknown): Scenario;
  safeConversationId(value: unknown): string | null;
  conversationFromPath(pathname: unknown): string | null;
  allowanceLevel(remaining: number | null): AllowanceLevel;
  lowestAllowance(tone: unknown, metering: Record<string, number> | undefined): number | null;
  isChat(frame: unknown): boolean;
  usesWorkIQ(request: unknown): boolean;
  requestFor(tone: unknown, workIQ?: boolean): { model: Model; tone: string } | null;
  parseRecords(data: unknown): any[];
  rewrite<T>(data: T, tone: unknown, workIQ?: boolean): { data: T | string; tones: string[]; chats: number; workIQ: boolean };
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
  newChat: { tone: string | null } | null;
  conversationTone: string | null;
}

interface Panel {
  mount(host: HTMLElement, handlers: { onConversationSelect(tone: string | null): void }): { update(state: PopupState): void };
}

declare var M365ExtenderModels: Model[];
declare var M365ExtenderCore: Core;
declare var M365ExtenderPanel: Panel;
