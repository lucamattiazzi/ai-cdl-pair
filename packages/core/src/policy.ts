import type { ToolClassification } from "./protocol.js";
import type { ToolPreview } from "./tool.js";

/** Policy decision for one validated tool call. */
export type PolicyDecision =
  | { readonly action: "allow" }
  | { readonly action: "require_approval"; readonly reason: string }
  | { readonly action: "deny"; readonly reason: string };

/** Safe metadata supplied to a policy engine. */
export interface PolicyContext {
  readonly toolName: string;
  readonly classification: ToolClassification;
  readonly preview?: ToolPreview;
}

/** Replaceable authorization policy for tool execution. */
export interface ApprovalPolicy {
  decide(context: PolicyContext): PolicyDecision | Promise<PolicyDecision>;
}

/** Reads run automatically; writes and structural operations require explicit approval. */
export function defaultApprovalPolicy(): ApprovalPolicy {
  return {
    decide: ({ classification }) =>
      classification === "read"
        ? { action: "allow" }
        : {
            action: "require_approval",
            reason: "Workbook mutations require user approval by default.",
          },
  };
}

/** Explicitly unsafe policy useful only for controlled automation and tests. */
export function allowAllPolicy(): ApprovalPolicy {
  return { decide: () => ({ action: "allow" }) };
}
