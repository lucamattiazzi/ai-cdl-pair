import { defineConfig } from "vitest/config";
export default defineConfig({
  test: { name: "agent-http", include: ["src/**/*.test.ts"], environment: "node" },
});
