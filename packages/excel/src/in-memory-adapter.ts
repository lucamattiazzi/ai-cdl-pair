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
import { assertMatrixShape, parseA1Range, toA1Cell } from "./range.js";

export interface InMemoryTableInput {
  readonly name: string;
  readonly range: string;
}

export interface InMemorySheetInput {
  readonly name: string;
  readonly values?: readonly (readonly CellValue[])[];
  readonly visibility?: SheetVisibility;
  readonly tables?: readonly InMemoryTableInput[];
}

export interface InMemoryExcelAdapterOptions {
  readonly sheets: readonly InMemorySheetInput[];
  readonly activeWorksheet?: string;
  readonly selection?: string;
}

interface MemorySheet {
  readonly name: string;
  readonly visibility: SheetVisibility;
  readonly values: CellValue[][];
  readonly tables: ExcelTableDescription[];
}

function abortIfNeeded(signal?: AbortSignal): void {
  if (signal?.aborted) throw signal.reason;
}

/** Deterministic Excel adapter for consumers and focused tests. */
export class InMemoryExcelAdapter implements ExcelAdapter {
  readonly #charts = new Map<string, ChartDescription[]>();
  readonly #sheets: MemorySheet[];
  readonly #listeners = new Set<SelectionChangedListener>();
  #activeWorksheet: string;
  #selection: string;

  constructor(options: InMemoryExcelAdapterOptions) {
    if (options.sheets.length === 0) throw new Error("At least one worksheet is required.");
    this.#sheets = options.sheets.map((sheet) => ({
      name: sheet.name,
      visibility: sheet.visibility ?? "visible",
      values: sheet.values?.map((row) => [...row]) ?? [],
      tables:
        sheet.tables?.map((table) => ({
          name: table.name,
          worksheet: sheet.name,
          range: parseA1Range(table.range).address,
        })) ?? [],
    }));
    this.#activeWorksheet = options.activeWorksheet ?? this.#sheets[0]?.name ?? "";
    this.#sheet(this.#activeWorksheet);
    this.#selection = parseA1Range(options.selection ?? "A1").address;
  }

  async listCharts(worksheet: string): Promise<readonly ChartDescription[]> {
    this.#sheet(worksheet);
    return (this.#charts.get(worksheet) ?? []).map((chart) => ({ ...chart }));
  }
  async createChart(input: CreateChartInput): Promise<ChartDescription> {
    abortIfNeeded(input.signal);
    await this.readRange(input);
    const charts = [...(await this.listCharts(input.worksheet))];
    if (charts.some((chart) => chart.name === input.name))
      throw new Error("Chart name already exists.");
    const chart = { id: crypto.randomUUID(), name: input.name, title: input.title };
    this.#charts.set(input.worksheet, [...charts, chart]);
    return chart;
  }

  async getContext(signal?: AbortSignal): Promise<ExcelContext> {
    abortIfNeeded(signal);
    return {
      activeWorksheet: this.#activeWorksheet,
      selection: this.#makeSelection(this.#activeWorksheet, this.#selection),
    };
  }

  async describeWorkbook(signal?: AbortSignal): Promise<WorkbookDescription> {
    abortIfNeeded(signal);
    return { sheets: this.#descriptions(), tables: this.#tables() };
  }

  async listSheets(signal?: AbortSignal): Promise<readonly SheetDescription[]> {
    abortIfNeeded(signal);
    return this.#descriptions();
  }

  async describeSheet(input: DescribeSheetInput): Promise<SheetDescription> {
    abortIfNeeded(input.signal);
    const index = this.#sheets.findIndex((sheet) => sheet.name === input.worksheet);
    const sheet = this.#sheets[index];
    if (!sheet) throw new Error(`Worksheet not found: ${input.worksheet}.`);
    return this.#description(sheet, index);
  }

  async readRange(input: ReadRangeInput): Promise<ReadRangeResult> {
    abortIfNeeded(input.signal);
    const sheet = this.#sheet(input.worksheet);
    const range = parseA1Range(input.range);
    const values = Array.from({ length: range.rowCount }, (_, row) =>
      Array.from(
        { length: range.columnCount },
        (_, column) => sheet.values[range.startRow + row]?.[range.startColumn + column] ?? null,
      ),
    );
    return {
      worksheet: sheet.name,
      range: range.address,
      values,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
    };
  }

  async writeRange(input: WriteRangeInput): Promise<WriteRangeResult> {
    abortIfNeeded(input.signal);
    const sheet = this.#sheet(input.worksheet);
    const range = parseA1Range(input.range);
    assertMatrixShape(input.values, range, "values");
    for (let row = 0; row < range.rowCount; row += 1) {
      const targetRow = range.startRow + row;
      if (!sheet.values[targetRow]) sheet.values[targetRow] = [];
      const target = sheet.values[targetRow];
      const source = input.values[row];
      if (!target || !source) throw new Error("Values must include every target row.");
      for (let column = 0; column < range.columnCount; column += 1) {
        const value = source[column];
        if (value === undefined) throw new Error("Values must include every target column.");
        target[range.startColumn + column] = value;
      }
    }
    return { worksheet: sheet.name, range: range.address, cellCount: range.cellCount };
  }

  async clearRange(input: ClearRangeInput): Promise<ClearRangeResult> {
    const range = parseA1Range(input.range);
    return this.writeRange({
      worksheet: input.worksheet,
      range: range.address,
      values: Array.from({ length: range.rowCount }, () =>
        Array.from({ length: range.columnCount }, () => null),
      ),
      ...(input.signal ? { signal: input.signal } : {}),
    });
  }

  async selectRange(input: SelectRangeInput): Promise<ExcelSelection> {
    abortIfNeeded(input.signal);
    this.#sheet(input.worksheet);
    this.#activeWorksheet = input.worksheet;
    this.#selection = parseA1Range(input.range).address;
    const selection = this.#makeSelection(this.#activeWorksheet, this.#selection);
    for (const listener of this.#listeners) listener(selection);
    return selection;
  }

  async listTables(input?: ListTablesInput): Promise<readonly ExcelTableDescription[]> {
    abortIfNeeded(input?.signal);
    if (input?.worksheet !== undefined) this.#sheet(input.worksheet);
    return this.#tables().filter(
      (table) => input?.worksheet === undefined || table.worksheet === input.worksheet,
    );
  }

  async readTable(input: ReadTableInput): Promise<ReadTableResult> {
    abortIfNeeded(input.signal);
    const table = this.#tables().find((candidate) => candidate.name === input.name);
    if (!table) throw new Error(`Table not found: ${input.name}.`);
    const data = await this.readRange({
      worksheet: table.worksheet,
      range: table.range,
      ...(input.signal ? { signal: input.signal } : {}),
    });
    return {
      ...table,
      headers: (data.values[0] ?? []).map((value) => (value === null ? "" : String(value))),
      rows: data.values.slice(1),
    };
  }

  async onSelectionChanged(listener: SelectionChangedListener): Promise<ExcelAdapterUnsubscribe> {
    this.#listeners.add(listener);
    return async () => {
      this.#listeners.delete(listener);
    };
  }

  #sheet(name: string): MemorySheet {
    const sheet = this.#sheets.find((candidate) => candidate.name === name);
    if (!sheet) throw new Error(`Worksheet not found: ${name}.`);
    return sheet;
  }

  #tables(): ExcelTableDescription[] {
    return this.#sheets.flatMap((sheet) => sheet.tables.map((table) => ({ ...table })));
  }

  #descriptions(): SheetDescription[] {
    return this.#sheets.map((sheet, index) => this.#description(sheet, index));
  }

  #description(sheet: MemorySheet, position: number): SheetDescription {
    const usedRange = this.#usedRange(sheet.values);
    return {
      name: sheet.name,
      position,
      visibility: sheet.visibility,
      ...(usedRange ? { usedRange } : {}),
      tables: sheet.tables.map((table) => ({ ...table })),
    };
  }

  #usedRange(values: readonly (readonly CellValue[])[]): string | undefined {
    let endRow = -1;
    let endColumn = -1;
    values.forEach((row, rowIndex) => {
      row.forEach((value, columnIndex) => {
        if (value !== null) {
          endRow = Math.max(endRow, rowIndex);
          endColumn = Math.max(endColumn, columnIndex);
        }
      });
    });
    return endRow < 0 ? undefined : `A1:${toA1Cell(endRow, endColumn)}`;
  }

  #makeSelection(worksheet: string, address: string): ExcelSelection {
    const range = parseA1Range(address);
    return {
      worksheet,
      range: range.address,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      getValues: async (signal?: AbortSignal) =>
        (await this.readRange({ worksheet, range: range.address, ...(signal ? { signal } : {}) }))
          .values,
    };
  }
}
