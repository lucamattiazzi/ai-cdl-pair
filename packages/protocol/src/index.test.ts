import { describe, expect, expectTypeOf, it } from "vitest";
import {
  METHOD_NAMES,
  methodSchemas,
  PROTOCOL_VERSION,
  type ProtocolMethodMap,
  type ProtocolRequest,
  parseProtocolMessage,
  pinnedRangeSchema,
  protocolMessageSchema,
  toolDescriptorRegistry,
} from "./index.js";

const methods = [
  "excel.context.get",
  "excel.workbook.describe",
  "excel.sheet.list",
  "excel.sheet.describe",
  "excel.range.read",
  "excel.range.write",
  "excel.range.clear",
  "excel.chart.list",
  "excel.chart.create",
  "excel.range.select",
  "excel.table.list",
  "excel.table.read",
  "excel.operation.preview",
  "excel.operation.commit",
  "excel.operation.undo",
  "excel.user.ask",
  "excel.user.notify",
] as const;

describe("protocol registry", () => {
  it("publishes exactly the vNext methods from one descriptor registry", () => {
    expect(PROTOCOL_VERSION).toBe("0.2");
    expect(METHOD_NAMES).toEqual(methods);
    expect(Object.keys(toolDescriptorRegistry)).toEqual(methods);
    expect(methodSchemas).toBe(toolDescriptorRegistry);
    expect(Object.values(toolDescriptorRegistry).map(({ capability }) => capability)).toContain(
      "interact",
    );
  });

  it("maps method params and results to TypeScript", () => {
    expectTypeOf<ProtocolMethodMap["excel.range.read"]["params"]>().toMatchObjectType<{
      range: { sheetId: string; address: string };
    }>();
    expectTypeOf<ProtocolRequest<"excel.user.notify">["params"]>().toMatchObjectType<{
      message: string;
    }>();
  });
});

describe("protocol validation", () => {
  it("parses a method request and rejects invalid method params", () => {
    const request = {
      protocolVersion: "0.2",
      type: "request",
      id: "request-1",
      method: "excel.range.read",
      params: { range: { sheetId: "sheet-1", address: "A1:B2" } },
    };

    expect(parseProtocolMessage(request)).toEqual(request);
    expect(() =>
      parseProtocolMessage({ ...request, params: { range: { sheetId: "sheet-1" } } }),
    ).toThrow();
  });

  it("validates response, error, selection, operation, and audit envelopes", () => {
    const messages = [
      {
        protocolVersion: "0.2",
        type: "response",
        id: "request-1",
        method: "excel.range.select",
        result: { selected: true },
      },
      {
        protocolVersion: "0.2",
        type: "error",
        id: "request-2",
        error: { code: "RANGE_NOT_FOUND", message: "Range not found" },
      },
      {
        protocolVersion: "0.2",
        type: "event",
        event: "excel.selection.changed",
        data: { range: { sheetId: "sheet-1", address: "C3" } },
      },
      {
        protocolVersion: "0.2",
        type: "event",
        event: "excel.operation.changed",
        data: { operationId: "op-1", status: "committed", occurredAt: "2026-08-31T10:00:00Z" },
      },
      {
        protocolVersion: "0.2",
        type: "event",
        event: "excel.audit.recorded",
        data: {
          auditId: "audit-1",
          action: "excel.range.write",
          outcome: "success",
          occurredAt: "2026-08-31T10:00:00Z",
        },
      },
      {
        protocolVersion: "0.2",
        type: "event",
        event: "pair.peer.changed",
        data: { role: "agent", connected: true },
      },
      {
        protocolVersion: "0.2",
        type: "event",
        event: "pair.chat.message",
        data: {
          messageId: "message-1",
          sender: "user",
          content: "Summarise the selected range.",
          occurredAt: "2026-08-31T10:00:00Z",
        },
      },
      {
        protocolVersion: "0.2",
        type: "event",
        event: "pair.chat.message",
        data: {
          messageId: "message-2",
          sender: "agent",
          content: "Sheet1!A1:D20 holds twenty rows of sales.",
          occurredAt: "2026-08-31T10:00:01Z",
        },
      },
    ];

    for (const message of messages) expect(protocolMessageSchema.parse(message)).toEqual(message);
  });

  it("supports all context modes and bounded pinned ranges", () => {
    expect(pinnedRangeSchema.parse({ sheetId: "sheet-1", address: "A1:B4" })).toEqual({
      sheetId: "sheet-1",
      address: "A1:B4",
    });
    for (const mode of ["follow-selection", "whole-workbook", "manual"]) {
      expect(
        toolDescriptorRegistry["excel.context.get"].resultSchema.parse({
          mode,
          pinnedRanges: [],
        }),
      ).toMatchObject({ mode });
    }
  });

  it("rejects unknown methods, events, versions, and extra envelope properties", () => {
    const base = { protocolVersion: "0.2", type: "request", id: "r1", params: {} };
    expect(() => parseProtocolMessage({ ...base, method: "excel.unknown" })).toThrow();
    expect(() => parseProtocolMessage({ ...base, protocolVersion: "0.1" })).toThrow();
    expect(() =>
      parseProtocolMessage({ ...base, method: "excel.sheet.list", extra: true }),
    ).toThrow();

    const chat = {
      protocolVersion: "0.2",
      type: "event",
      event: "pair.chat.message",
      data: {
        messageId: "message-1",
        sender: "user",
        content: "Hello",
        occurredAt: "2026-08-31T10:00:00Z",
      },
    };
    expect(() => parseProtocolMessage({ ...chat, data: { ...chat.data, content: "" } })).toThrow();
    expect(() =>
      parseProtocolMessage({ ...chat, data: { ...chat.data, sender: "relay" } }),
    ).toThrow();
  });
});
