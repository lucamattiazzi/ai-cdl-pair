import type { z } from "zod";
import type { JsonSchema, JsonValue, ToolCapability, ToolClassification } from "./protocol.js";
import { toJsonSchema } from "./protocol.js";

/** Bounded, UI-friendly preview of a proposed mutation. */
export interface ToolPreview {
  readonly summary: string;
  readonly target?: string;
  readonly affectedCells: number;
  readonly sample?: JsonValue;
  readonly warnings: readonly string[];
}

/** Execution context passed to a workbook tool. */
export interface ToolExecutionContext {
  readonly signal: AbortSignal;
  readonly sessionId: string;
  readonly turnId: string;
  readonly toolCallId: string;
}

/** Typed executable tool definition. */
export interface ToolDefinition<TInput = unknown, TOutput extends JsonValue = JsonValue>
  extends ToolCapability {
  readonly input: z.ZodType<TInput>;
  execute(input: TInput, context: ToolExecutionContext): Promise<TOutput>;
  preview?(input: TInput, context: ToolExecutionContext): Promise<ToolPreview>;
  estimateCells?(input: TInput): { readonly read: number; readonly written: number };
}

/** Options for defining a validated tool without repeating its JSON Schema. */
export interface DefineToolOptions<TInput, TOutput extends JsonValue> {
  readonly name: string;
  readonly description: string;
  readonly classification: ToolClassification;
  readonly input: z.ZodType<TInput>;
  readonly execute: (input: TInput, context: ToolExecutionContext) => Promise<TOutput>;
  readonly preview?: (input: TInput, context: ToolExecutionContext) => Promise<ToolPreview>;
  readonly estimateCells?: (input: TInput) => { readonly read: number; readonly written: number };
}

/** Define a tool and derive its advertised JSON Schema from the runtime validator. */
export function defineTool<TInput, TOutput extends JsonValue>(
  options: DefineToolOptions<TInput, TOutput>,
): ToolDefinition<TInput, TOutput> {
  return { ...options, inputSchema: toJsonSchema(options.input) as JsonSchema };
}
