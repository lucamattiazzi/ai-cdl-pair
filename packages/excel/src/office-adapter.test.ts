import { afterEach, expect, it, vi } from "vitest";
import { OfficeJsExcelAdapter } from "./office-adapter.js";

afterEach(() => vi.unstubAllGlobals());

it("returns formulas alongside computed values", async () => {
  const range = {
    load: vi.fn(),
    values: [[30]],
    formulas: [["=SUM(B1:B2)"]],
    rowCount: 1,
    columnCount: 1,
  };
  const context = {
    sync: vi.fn(),
    workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
  };
  vi.stubGlobal("Excel", { run: (fn: (ctx: typeof context) => unknown) => fn(context) });
  expect(
    await new OfficeJsExcelAdapter().readRange({ worksheet: "Sheet1", range: "A1" }),
  ).toMatchObject({ values: [[30]], formulas: [["=SUM(B1:B2)"]] });
});

it("turns null into an actual empty cell instead of an Office.js no-op", async () => {
  const range = { values: [[99]] as unknown[][] };
  const context = {
    sync: vi.fn(),
    workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
  };
  vi.stubGlobal("Excel", { run: (fn: (ctx: typeof context) => unknown) => fn(context) });
  await new OfficeJsExcelAdapter().writeRange({
    worksheet: "Sheet1",
    range: "A1",
    values: [[null]],
  });
  expect(range.values).toEqual([[""]]);
});

it("does not mistake literal text beginning with equals for a formula", async () => {
  const range = {
    load: vi.fn(),
    values: [["=literal"]],
    formulas: [["=literal"]],
    rowCount: 1,
    columnCount: 1,
    getCell: () => ({
      getSpecialCellsOrNullObject: () => ({ load: vi.fn(), isNullObject: true }),
    }),
  };
  const context = {
    sync: vi.fn(),
    workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
  };
  vi.stubGlobal("Excel", { run: (fn: (ctx: typeof context) => unknown) => fn(context) });
  expect(
    await new OfficeJsExcelAdapter().readRange({ worksheet: "Sheet1", range: "A1" }),
  ).toMatchObject({ formulas: [[null]] });
});

it("restores formulas and constants without keeping overwritten values", async () => {
  const range = { formulas: [] as unknown[][] };
  const context = {
    sync: vi.fn(),
    workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
  };
  vi.stubGlobal("Excel", { run: (fn: (ctx: typeof context) => unknown) => fn(context) });
  await new OfficeJsExcelAdapter().restoreRange({
    worksheet: "Sheet1",
    range: "A1:C1",
    rowCount: 1,
    columnCount: 3,
    values: [[30, "=literal", null]],
    formulas: [["=SUM(B1:B2)", null, null]],
  });
  expect(range.formulas).toEqual([["=SUM(B1:B2)", "'=literal", ""]]);
});

it("creates charts through Office.js and refuses a colliding name before adding", async () => {
  const existing = { load: vi.fn(), isNullObject: false };
  const chart = { id: "chart-1", name: "", title: { text: "", visible: false }, load: vi.fn() };
  const add = vi.fn(() => chart);
  const range = {};
  const context = {
    sync: vi.fn(),
    workbook: {
      worksheets: {
        getItem: () => ({
          charts: { getItemOrNullObject: () => existing, add },
          getRange: () => range,
        }),
      },
    },
  };
  vi.stubGlobal("Excel", {
    run: (fn: (ctx: typeof context) => unknown) => fn(context),
    ChartType: { columnClustered: "ColumnClustered" },
    ChartSeriesBy: { auto: "Auto" },
  });
  const adapter = new OfficeJsExcelAdapter();
  const input = {
    worksheet: "Sales",
    range: "A1:B4",
    name: "Monthly",
    title: "Monthly sales",
    chartType: "column" as const,
  };
  await expect(adapter.createChart(input)).rejects.toThrow("already exists");
  expect(add).not.toHaveBeenCalled();
  existing.isNullObject = true;
  expect(await adapter.createChart(input)).toEqual({
    id: "chart-1",
    name: "Monthly",
    title: "Monthly sales",
  });
  expect(add).toHaveBeenCalledWith("ColumnClustered", range, "Auto");
  expect(chart.title.visible).toBe(true);
});
