import { expectTypeOf, it } from "vitest";
import type {
  ClearRangeInput,
  ClearRangeResult,
  ExcelAdapter,
  ExcelContext,
  ExcelSelection,
  InMemoryExcelAdapter,
  OfficeJsExcelAdapter,
  ReadRangeInput,
  ReadRangeResult,
  SelectRangeInput,
  SelectRangeResult,
  SheetDescription,
  WriteRangeInput,
  WriteRangeResult,
} from "./index.js";

it("exposes the additive Excel adapter contract", () => {
  expectTypeOf<InMemoryExcelAdapter>().toMatchTypeOf<ExcelAdapter>();
  expectTypeOf<OfficeJsExcelAdapter>().toMatchTypeOf<ExcelAdapter>();
  expectTypeOf<ExcelSelection>().toHaveProperty("getValues");
  expectTypeOf<ExcelContext>().toHaveProperty("selection");
  expectTypeOf<SheetDescription>().toHaveProperty("usedRange");
  expectTypeOf<ReadRangeInput>().toHaveProperty("range");
  expectTypeOf<ReadRangeResult>().toHaveProperty("values");
  expectTypeOf<WriteRangeInput>().toHaveProperty("values");
  expectTypeOf<WriteRangeResult>().toHaveProperty("cellCount");
  expectTypeOf<ClearRangeInput>().toHaveProperty("range");
  expectTypeOf<ClearRangeResult>().toHaveProperty("cellCount");
  expectTypeOf<SelectRangeInput>().toHaveProperty("range");
  expectTypeOf<SelectRangeResult>().toMatchTypeOf<ExcelSelection>();
});
