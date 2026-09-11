import { defineConfig } from "vitest/config";
export default defineConfig({
  test: { name: "excel", include: ["src/**/*.test.ts"], environment: "node" },
});
