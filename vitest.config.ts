import { resolve } from "node:path";
import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      "@ai-cdl/protocol": resolve("packages/protocol/src/index.ts"),
      "@ai-cdl/transport": resolve("packages/transport/src/index.ts"),
      "@ai-cdl/core": resolve("packages/core/src/index.ts"),
      "@ai-cdl/excel": resolve("packages/excel/src/index.ts"),
      "@ai-cdl/addin-core": resolve("packages/addin-core/src/index.ts"),
      "@ai-cdl/pair": resolve("packages/pair/src/index.ts"),
      "@ai-cdl/pair-cli": resolve("packages/pair-cli/src/index.ts"),
      "@ai-cdl/agent-http": resolve("packages/agent-http/src/index.ts"),
      "@ai-cdl/testing": resolve("packages/testing/src/index.ts"),
      "@ai-cdl/cli/project": resolve("packages/cli/src/project.ts"),
      "@ai-cdl/cli/manifest": resolve("packages/cli/src/manifest.ts"),
    },
  },
  test: {
    include: ["packages/**/*.test.ts", "apps/**/*.test.ts"],
    coverage: { provider: "v8", reporter: ["text-summary", "json-summary"] },
  },
});
