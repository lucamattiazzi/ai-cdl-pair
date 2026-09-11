import type { CellValue } from "./driver.js";

export interface ExcelOperationInput {
  readonly signal?: AbortSignal;
}

/** A bounded selection whose cell values are fetched only when requested. */
export interface ExcelSelection {
  readonly worksheet: string;
  readonly range: string;
  readonly rowCount: number;
  readonly columnCount: number;
  getValues(signal?: AbortSignal): Promise<readonly (readonly CellValue[])[]>;
}

export interface ExcelContext {
  readonly activeWorksheet: string;
  readonly selection: ExcelSelection;
}

export interface ExcelTableDescription {
  readonly name: string;
  readonly worksheet: string;
  readonly range: string;
}

export type SheetVisibility = "visible" | "hidden" | "veryHidden";

export interface SheetDescription {
  readonly name: string;
  readonly position: number;
  readonly visibility: SheetVisibility;
  readonly usedRange?: string;
  readonly tables: readonly ExcelTableDescription[];
}

export interface WorkbookDescription {
  readonly sheets: readonly SheetDescription[];
  readonly tables: readonly ExcelTableDescription[];
}

export interface DescribeSheetInput extends ExcelOperationInput {
  readonly worksheet: string;
}

export interface ReadRangeInput extends ExcelOperationInput {
  readonly worksheet: string;
  readonly range: string;
}

export interface ReadRangeResult {
  readonly worksheet: string;
  readonly range: string;
  readonly values: readonly (readonly CellValue[])[];
  readonly rowCount: number;
  readonly columnCount: number;
  readonly formulas?: readonly (readonly (string | null)[])[];
}

export interface WriteRangeInput extends ReadRangeInput {
  readonly values: readonly (readonly CellValue[])[];
}

export interface WriteRangeResult {
  readonly worksheet: string;
  readonly range: string;
  readonly cellCount: number;
}

export interface ChartDescription {
  readonly id: string;
  readonly name: string;
  readonly title: string;
}
export interface CreateChartInput extends ReadRangeInput {
  readonly name: string;
  readonly title: string;
  readonly chartType: "column" | "bar" | "line" | "pie" | "scatter";
}

export interface ClearRangeInput extends ReadRangeInput {}

export interface ClearRangeResult extends WriteRangeResult {}

export interface SelectRangeInput extends ReadRangeInput {}

export type SelectRangeResult = ExcelSelection;

export interface ListTablesInput extends ExcelOperationInput {
  readonly worksheet?: string;
}

export interface ReadTableInput extends ExcelOperationInput {
  readonly name: string;
}

export interface ReadTableResult extends ExcelTableDescription {
  readonly headers: readonly string[];
  readonly rows: readonly (readonly CellValue[])[];
}

export type SelectionChangedListener = (selection: ExcelSelection) => void;
export type ExcelAdapterUnsubscribe = () => Promise<void>;

/** Consumer-oriented Excel port for workbook context, ranges, tables, and selection events. */
export interface ExcelAdapter {
  createChart?(input: CreateChartInput): Promise<ChartDescription>;
  listCharts?(worksheet: string): Promise<readonly ChartDescription[]>;
  getContext(signal?: AbortSignal): Promise<ExcelContext>;
  describeWorkbook(signal?: AbortSignal): Promise<WorkbookDescription>;
  listSheets(signal?: AbortSignal): Promise<readonly SheetDescription[]>;
  describeSheet(input: DescribeSheetInput): Promise<SheetDescription>;
  readRange(input: ReadRangeInput): Promise<ReadRangeResult>;
  writeRange(input: WriteRangeInput): Promise<WriteRangeResult>;
  clearRange(input: ClearRangeInput): Promise<ClearRangeResult>;
  /** Restore values and formulas from a bounded snapshot, preserving current formats. */
  restoreRange?(snapshot: ReadRangeResult): Promise<void>;
  selectRange(input: SelectRangeInput): Promise<SelectRangeResult>;
  listTables(input?: ListTablesInput): Promise<readonly ExcelTableDescription[]>;
  readTable(input: ReadTableInput): Promise<ReadTableResult>;
  onSelectionChanged(listener: SelectionChangedListener): Promise<ExcelAdapterUnsubscribe>;
}
