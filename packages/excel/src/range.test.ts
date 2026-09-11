import { describe, expect, it } from "vitest";
import { assertMatrixShape, parseA1Range } from "./range.js";

describe("parseA1Range", () => {
  it("parses and canonicalizes bounded ranges", () => {
    expect(parseA1Range("$b$2:D4")).toMatchObject({
      address: "B2:D4",
      rowCount: 3,
      columnCount: 3,
      cellCount: 9,
    });
  });

  it("rejects whole-column and reversed ranges", () => {
    expect(() => parseA1Range("A:A")).toThrow(/Invalid bounded/);
    expect(() => parseA1Range("B2:A1")).toThrow(/Invalid bounded/);
  });

  it("validates matrix shape", () => {
    expect(() => assertMatrixShape([[1]], parseA1Range("A1:B1"), "values")).toThrow(/1 by 2/);
  });
});
