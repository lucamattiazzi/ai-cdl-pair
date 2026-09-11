import { resolve } from "node:path";
import { createPairServer } from "./index.js";

async function main(): Promise<void> {
  const publicOrigin = process.env.PAIR_PUBLIC_ORIGIN;
  if (!publicOrigin) throw new Error("PAIR_PUBLIC_ORIGIN is required.");

  const port = Number(process.env.PORT ?? "3000");
  if (!Number.isInteger(port) || port < 1 || port > 65_535) throw new Error("PORT is invalid.");

  const server = await createPairServer({
    host: process.env.HOST ?? "127.0.0.1",
    port,
    publicOrigin,
    ...(process.env.PAIR_ADDIN_ORIGIN ? { addinOrigin: process.env.PAIR_ADDIN_ORIGIN } : {}),
    ...(process.env.PAIR_AGENT_ORIGIN ? { agentOrigin: process.env.PAIR_AGENT_ORIGIN } : {}),
    staticDirectory: resolve(process.env.PAIR_STATIC_DIR ?? "apps/pair-addin/dist"),
    bridgeScript: resolve(
      process.env.PAIR_BRIDGE_SCRIPT ?? "skills/ai-cdl-pair/scripts/session.mjs",
    ),
  });

  console.log(`AI-CDL Pair server listening on ${server.host}:${server.port}`);

  const shutdown = async (): Promise<void> => {
    await server.close();
    process.exit(0);
  };
  process.once("SIGINT", () => void shutdown());
  process.once("SIGTERM", () => void shutdown());
}

void main();
