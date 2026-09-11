import type { JsonValue, SerializedAiCdlError } from "./protocol.js";

/** Stable AI-CDL error codes. */
export type AiCdlErrorCode =
  | "AI_CDL_ACTIVE_TURN"
  | "AI_CDL_AGENT_PROTOCOL_INVALID"
  | "AI_CDL_APPROVAL_NOT_FOUND"
  | "AI_CDL_BUDGET_EXCEEDED"
  | "AI_CDL_CANCELLED"
  | "AI_CDL_INVALID_STATE_TRANSITION"
  | "AI_CDL_POLICY_DENIED"
  | "AI_CDL_TIMEOUT"
  | "AI_CDL_TOOL_ARGUMENTS_INVALID"
  | "AI_CDL_TOOL_EXECUTION_FAILED"
  | "AI_CDL_TOOL_NOT_FOUND";

/** Options for a structured, safe public error. */
export interface AiCdlErrorOptions {
  readonly code: AiCdlErrorCode | (string & {});
  readonly message: string;
  readonly context?: Readonly<Record<string, JsonValue>>;
  readonly probableCause?: string;
  readonly suggestedAction?: string;
  readonly documentationUrl?: string;
  readonly cause?: unknown;
}

/** Base error carrying a stable code and remediation fields. */
export class AiCdlError extends Error {
  readonly code: string;
  readonly context: Readonly<Record<string, JsonValue>>;
  readonly probableCause: string | undefined;
  readonly suggestedAction: string | undefined;
  readonly documentationUrl: string | undefined;

  constructor(options: AiCdlErrorOptions) {
    super(options.message, { cause: options.cause });
    this.name = "AiCdlError";
    this.code = options.code;
    this.context = options.context ?? {};
    this.probableCause = options.probableCause;
    this.suggestedAction = options.suggestedAction;
    this.documentationUrl = options.documentationUrl;
  }

  /** Return a JSON-safe form that deliberately omits the cause chain. */
  toJSON(): SerializedAiCdlError {
    return {
      code: this.code,
      message: this.message,
      context: this.context,
      ...(this.probableCause === undefined ? {} : { probableCause: this.probableCause }),
      ...(this.suggestedAction === undefined ? {} : { suggestedAction: this.suggestedAction }),
      ...(this.documentationUrl === undefined ? {} : { documentationUrl: this.documentationUrl }),
    };
  }
}

/** Convert an unknown failure into a safe structured error. */
export function asAiCdlError(
  error: unknown,
  fallbackCode = "AI_CDL_TOOL_EXECUTION_FAILED",
): AiCdlError {
  if (error instanceof AiCdlError) return error;
  return new AiCdlError({
    code: fallbackCode,
    message: error instanceof Error ? error.message : "An unknown AI-CDL error occurred.",
    suggestedAction: "Inspect the associated trace event and retry the operation.",
    cause: error,
  });
}
