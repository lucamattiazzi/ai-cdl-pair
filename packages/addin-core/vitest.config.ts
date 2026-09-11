import { resolve } from "node:path";
import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      "@ai-cdl/protocol": resolve("../protocol/src/index.ts"),
      "@ai-cdl/excel": resolve("../excel/src/index.ts"),
    },
  },
  test: { include: ["src/**/*.test.ts"] },
});
