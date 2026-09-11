/// <reference types="office-js" />

import type {
  ChartDescription,
  ClearRangeInput,
  ClearRangeResult,
  CreateChartInput,
  DescribeSheetInput,
  ExcelAdapter,
  ExcelAdapterUnsubscribe,
  ExcelContext,
  ExcelSelection,
  ExcelTableDescription,
  ListTablesInput,
  ReadRangeInput,
  ReadRangeResult,
  ReadTableInput,
  ReadTableResult,
  SelectionChangedListener,
  SelectRangeInput,
  SheetDescription,
  SheetVisibility,
  WorkbookDescription,
  WriteRangeInput,
  WriteRangeResult,
} from "./adapter.js";
import type { CellValue } from "./driver.js";
import { assertMatrixShape, parseA1Range } from "./range.js";

function abortIfNeeded(signal?: AbortSignal): void {
  if (signal?.aborted) throw signal.reason;
}

function valuesOf(values: unknown[][]): CellValue[][] {
  return values.map((row) =>
    row.map((value) =>
      typeof value === "string" || typeof value === "number" || typeof value === "boolean"
        ? value
        : null,
    ),
  );
}

function localAddress(address: string): string {
  return parseA1Range(address.split("!").pop()?.replaceAll("$", "") ?? address).address;
}

function visibilityOf(
  visibility: Excel.SheetVisibility | "Visible" | "Hidden" | "VeryHidden",
): SheetVisibility {
  if (visibility === Excel.SheetVisibility.hidden) return "hidden";
  if (visibility === Excel.SheetVisibility.veryHidden) return "veryHidden";
  return "visible";
}

/** Thin Office.js implementation for the workbook open in the current Excel host. */
export class OfficeJsExcelAdapter implements ExcelAdapter {
  async listCharts(worksheet: string): Promise<readonly ChartDescription[]> {
    return Excel.run(async (context) => {
      const charts = context.workbook.worksheets.getItem(worksheet).charts;
      charts.load("items/id,items/name,items/title/text");
      await context.sync();
      return charts.items.map((chart) => ({
        id: chart.id,
        name: chart.name,
        title: chart.title.text,
      }));
    });
  }
  async createChart(input: CreateChartInput): Promise<ChartDescription> {
    abortIfNeeded(input.signal);
    const range = parseA1Range(input.range);
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(input.worksheet);
      const existing = sheet.charts.getItemOrNullObject(input.name);
      existing.load("isNullObject");
      await context.sync();
      abortIfNeeded(input.signal);
      if (!existing.isNullObject) throw new Error("Chart name already exists.");
      const types = {
        column: Excel.ChartType.columnClustered,
        bar: Excel.ChartType.barClustered,
        line: Excel.ChartType.line,
        pie: Excel.ChartType.pie,
        scatter: Excel.ChartType.xyscatter,
      };
      const chart = sheet.charts.add(
        types[input.chartType],
        sheet.getRange(range.address),
        Excel.ChartSeriesBy.auto,
      );
      chart.name = input.name;
      chart.title.text = input.title;
      chart.title.visible = true;
      chart.load("id,name,title/text");
      await context.sync();
      return { id: chart.id, name: chart.name, title: chart.title.text };
    });
  }

  async getContext(signal?: AbortSignal): Promise<ExcelContext> {
    abortIfNeeded(signal);
    return Excel.run(async (context) => {
      const selected = context.workbook.getSelectedRange();
      selected.load("address,rowCount,columnCount");
      selected.worksheet.load("name");
      await context.sync();
      abortIfNeeded(signal);
      return {
        activeWorksheet: selected.worksheet.name,
        selection: this.#selection(
          selected.worksheet.name,
          localAddress(selected.address),
          selected.rowCount,
          selected.columnCount,
        ),
      };
    });
  }

  async describeWorkbook(signal?: AbortSignal): Promise<WorkbookDescription> {
    const sheets = await this.listSheets(signal);
    return { sheets, tables: sheets.flatMap((sheet) => sheet.tables) };
  }

  async listSheets(signal?: AbortSignal): Promise<readonly SheetDescription[]> {
    abortIfNeeded(signal);
    return Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load("items/name,items/position,items/visibility");
      await context.sync();
      const entries = worksheets.items.map((sheet) => {
        const usedRange = sheet.getUsedRangeOrNullObject(true);
        const tables = sheet.tables;
        usedRange.load("address,isNullObject");
        tables.load("items/name");
        return { sheet, usedRange, tables };
      });
      await context.sync();
      const tableEntries = entries.flatMap(({ sheet, tables }) =>
        tables.items.map((table) => {
          const range = table.getRange();
          range.load("address");
          return { worksheet: sheet.name, table, range };
        }),
      );
      await context.sync();
      abortIfNeeded(signal);
      return entries.map(({ sheet, usedRange, tables }) => ({
        name: sheet.name,
        position: sheet.position,
        visibility: visibilityOf(sheet.visibility),
        ...(usedRange.isNullObject ? {} : { usedRange: localAddress(usedRange.address) }),
        tables: tables.items.map((table) => {
          const entry = tableEntries.find((candidate) => candidate.table === table);
          return {
            name: table.name,
            worksheet: sheet.name,
            range: entry ? localAddress(entry.range.address) : "",
          };
        }),
      }));
    });
  }

  async describeSheet(input: DescribeSheetInput): Promise<SheetDescription> {
    const sheets = await this.listSheets(input.signal);
    const sheet = sheets.find((candidate) => candidate.name === input.worksheet);
    if (!sheet) throw new Error(`Worksheet not found: ${input.worksheet}.`);
    return sheet;
  }

  async readRange(input: ReadRangeInput): Promise<ReadRangeResult> {
    abortIfNeeded(input.signal);
    const parsed = parseA1Range(input.range);
    return Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem(input.worksheet).getRange(parsed.address);
      range.load("values,formulas,rowCount,columnCount");
      await context.sync();
      abortIfNeeded(input.signal);
      const ambiguous: { row: number; column: number; formulas: Excel.RangeAreas }[] = [];
      range.formulas.forEach((row: unknown[], r: number) => {
        row.forEach((formula, c) => {
          if (
            typeof formula === "string" &&
            formula.startsWith("=") &&
            formula === range.values[r]?.[c]
          ) {
            const formulas = range.getCell(r, c).getSpecialCellsOrNullObject("Formulas");
            formulas.load("isNullObject");
            ambiguous.push({ row: r, column: c, formulas });
          }
        });
      });
      if (ambiguous.length > 0) await context.sync();
      abortIfNeeded(input.signal);
      return {
        worksheet: input.worksheet,
        range: parsed.address,
        values: valuesOf(range.values),
        formulas: range.formulas.map((row: unknown[], r: number) =>
          row.map((value, c) => {
            const candidate = ambiguous.find((item) => item.row === r && item.column === c);
            return typeof value === "string" &&
              value.startsWith("=") &&
              candidate?.formulas.isNullObject !== true
              ? value
              : null;
          }),
        ),
        rowCount: range.rowCount,
        columnCount: range.columnCount,
      };
    });
  }

  async writeRange(input: WriteRangeInput): Promise<WriteRangeResult> {
    abortIfNeeded(input.signal);
    const parsed = parseA1Range(input.range);
    assertMatrixShape(input.values, parsed, "values");
    await Excel.run(async (context) => {
      context.workbook.worksheets.getItem(input.worksheet).getRange(parsed.address).values =
        input.values.map((row) => row.map((value) => value ?? ""));
      await context.sync();
      abortIfNeeded(input.signal);
    });
    return { worksheet: input.worksheet, range: parsed.address, cellCount: parsed.cellCount };
  }

  async restoreRange(snapshot: ReadRangeResult): Promise<void> {
    const parsed = parseA1Range(snapshot.range);
    assertMatrixShape(snapshot.values, parsed, "values");
    if (snapshot.formulas) assertMatrixShape(snapshot.formulas, parsed, "formulas");
    await Excel.run(async (context) => {
      const range = context.workbook.worksheets
        .getItem(snapshot.worksheet)
        .getRange(parsed.address);
      range.formulas = snapshot.values.map((row, r) =>
        row.map((value, c) => {
          const formula = snapshot.formulas?.[r]?.[c];
          if (formula) return formula;
          // The formulas property also accepts constants. Escape literal text on restoration.
          return typeof value === "string" && value.length > 0 ? `'${value}` : (value ?? "");
        }),
      );
      await context.sync();
    });
  }

  async clearRange(input: ClearRangeInput): Promise<ClearRangeResult> {
    abortIfNeeded(input.signal);
    const parsed = parseA1Range(input.range);
    await Excel.run(async (context) => {
      context.workbook.worksheets
        .getItem(input.worksheet)
        .getRange(parsed.address)
        .clear(Excel.ClearApplyTo.contents);
      await context.sync();
      abortIfNeeded(input.signal);
    });
    return { worksheet: input.worksheet, range: parsed.address, cellCount: parsed.cellCount };
  }

  async selectRange(input: SelectRangeInput): Promise<ExcelSelection> {
    abortIfNeeded(input.signal);
    const parsed = parseA1Range(input.range);
    await Excel.run(async (context) => {
      context.workbook.worksheets.getItem(input.worksheet).getRange(parsed.address).select();
      await context.sync();
      abortIfNeeded(input.signal);
    });
    return this.#selection(input.worksheet, parsed.address, parsed.rowCount, parsed.columnCount);
  }

  async listTables(input?: ListTablesInput): Promise<readonly ExcelTableDescription[]> {
    const sheets = await this.listSheets(input?.signal);
    return sheets
      .filter((sheet) => input?.worksheet === undefined || sheet.name === input.worksheet)
      .flatMap((sheet) => sheet.tables);
  }

  async readTable(input: ReadTableInput): Promise<ReadTableResult> {
    abortIfNeeded(input.signal);
    return Excel.run(async (context) => {
      const table = context.workbook.tables.getItem(input.name);
      const header = table.getHeaderRowRange();
      const body = table.getDataBodyRange();
      const full = table.getRange();
      table.load("name");
      table.worksheet.load("name");
      header.load("values");
      body.load("values");
      full.load("address");
      await context.sync();
      abortIfNeeded(input.signal);
      return {
        name: table.name,
        worksheet: table.worksheet.name,
        range: localAddress(full.address),
        headers: (valuesOf(header.values)[0] ?? []).map((value) =>
          value === null ? "" : String(value),
        ),
        rows: valuesOf(body.values),
      };
    });
  }

  async onSelectionChanged(listener: SelectionChangedListener): Promise<ExcelAdapterUnsubscribe> {
    const handler = async (event: Excel.WorksheetSelectionChangedEventArgs): Promise<void> => {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(event.worksheetId);
        sheet.load("name");
        await context.sync();
        const parsed = parseA1Range(event.address);
        listener(this.#selection(sheet.name, parsed.address, parsed.rowCount, parsed.columnCount));
      });
    };
    const registration = await Excel.run(async (context) => {
      const result = context.workbook.worksheets.onSelectionChanged.add(handler);
      await context.sync();
      return result;
    });
    return async () => {
      registration.remove();
      await registration.context.sync();
    };
  }

  #selection(
    worksheet: string,
    range: string,
    rowCount: number,
    columnCount: number,
  ): ExcelSelection {
    return {
      worksheet,
      range,
      rowCount,
      columnCount,
      getValues: async (signal?: AbortSignal) =>
        (await this.readRange({ worksheet, range, ...(signal ? { signal } : {}) })).values,
    };
  }
}
