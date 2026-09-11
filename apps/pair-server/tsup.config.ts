import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts", "src/main.ts"],
  format: ["cjs"],
  dts: true,
  sourcemap: true,
  clean: true,
  noExternal: ["@ai-cdl/protocol", "@ai-cdl/transport", "ws", "zod"],
});
