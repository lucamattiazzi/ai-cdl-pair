import { resolve } from "node:path";
import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      "@ai-cdl/protocol": resolve("../protocol/src/index.ts"),
      "@ai-cdl/transport": resolve("../transport/src/index.ts"),
      "@ai-cdl/excel": resolve("../excel/src/index.ts"),
      "@ai-cdl/addin-core": resolve("../addin-core/src/index.ts"),
    },
  },
  test: { include: ["src/**/*.test.ts"] },
});
