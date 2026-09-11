import { describe, expect, it, vi } from "vitest";
import { InMemoryExcelAdapter } from "./in-memory-adapter.js";

describe("InMemoryExcelAdapter", () => {
  it("reads, writes, clears, and selects bounded ranges", async () => {
    const adapter = new InMemoryExcelAdapter({
      sheets: [
        {
          name: "Data",
          values: [
            ["Item", "Value"],
            ["Copper", 9.5],
          ],
        },
      ],
    });

    await expect(adapter.readRange({ worksheet: "Data", range: "$a$2:B2" })).resolves.toEqual({
      worksheet: "Data",
      range: "A2:B2",
      values: [["Copper", 9.5]],
      rowCount: 1,
      columnCount: 2,
    });

    await expect(
      adapter.writeRange({ worksheet: "Data", range: "B2:B3", values: [[10], [11]] }),
    ).resolves.toEqual({ worksheet: "Data", range: "B2:B3", cellCount: 2 });
    expect((await adapter.readRange({ worksheet: "Data", range: "A2:B3" })).values).toEqual([
      ["Copper", 10],
      [null, 11],
    ]);

    await expect(adapter.clearRange({ worksheet: "Data", range: "A2:B2" })).resolves.toEqual({
      worksheet: "Data",
      range: "A2:B2",
      cellCount: 2,
    });
    expect((await adapter.readRange({ worksheet: "Data", range: "A2:B2" })).values).toEqual([
      [null, null],
    ]);

    await expect(adapter.selectRange({ worksheet: "Data", range: "B3" })).resolves.toMatchObject({
      worksheet: "Data",
      range: "B3:B3",
      rowCount: 1,
      columnCount: 1,
    });
    const context = await adapter.getContext();
    expect(context.activeWorksheet).toBe("Data");
    await expect(context.selection.getValues()).resolves.toEqual([[11]]);
  });

  it("validates write dimensions before changing cells", async () => {
    const adapter = new InMemoryExcelAdapter({ sheets: [{ name: "Sheet1" }] });

    await expect(
      adapter.writeRange({ worksheet: "Sheet1", range: "A1:B1", values: [[1]] }),
    ).rejects.toThrow(/1 by 2/);
    expect((await adapter.readRange({ worksheet: "Sheet1", range: "A1:B1" })).values).toEqual([
      [null, null],
    ]);
  });

  it("emits selection metadata with values loaded lazily", async () => {
    const adapter = new InMemoryExcelAdapter({
      sheets: [{ name: "Sheet1", values: [["selected"]] }],
    });
    const listener = vi.fn();
    const unsubscribe = await adapter.onSelectionChanged(listener);

    await adapter.selectRange({ worksheet: "Sheet1", range: "A1" });

    expect(listener).toHaveBeenCalledOnce();
    const selection = listener.mock.calls[0]?.[0];
    expect(selection).toMatchObject({
      worksheet: "Sheet1",
      range: "A1:A1",
      rowCount: 1,
      columnCount: 1,
    });
    await expect(selection?.getValues()).resolves.toEqual([["selected"]]);

    await unsubscribe();
    await adapter.selectRange({ worksheet: "Sheet1", range: "A1" });
    expect(listener).toHaveBeenCalledOnce();
  });

  it("describes sheets and reads named tables", async () => {
    const adapter = new InMemoryExcelAdapter({
      sheets: [
        {
          name: "Data",
          values: [
            ["Item", "Value"],
            ["Copper", 9.5],
            ["Tin", 3.1],
          ],
          tables: [{ name: "Prices", range: "A1:B3" }],
        },
        { name: "Notes", visibility: "hidden" },
      ],
    });

    await expect(adapter.describeWorkbook()).resolves.toMatchObject({
      sheets: [
        { name: "Data", position: 0, visibility: "visible", usedRange: "A1:B3" },
        { name: "Notes", position: 1, visibility: "hidden" },
      ],
      tables: [{ name: "Prices", worksheet: "Data", range: "A1:B3" }],
    });
    await expect(adapter.listSheets()).resolves.toHaveLength(2);
    await expect(adapter.describeSheet({ worksheet: "Data" })).resolves.toMatchObject({
      name: "Data",
      tables: [{ name: "Prices" }],
    });
    await expect(adapter.listTables({ worksheet: "Data" })).resolves.toEqual([
      { name: "Prices", worksheet: "Data", range: "A1:B3" },
    ]);
    await expect(adapter.readTable({ name: "Prices" })).resolves.toEqual({
      name: "Prices",
      worksheet: "Data",
      range: "A1:B3",
      headers: ["Item", "Value"],
      rows: [
        ["Copper", 9.5],
        ["Tin", 3.1],
      ],
    });
  });
});
