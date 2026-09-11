import { execFileSync, spawnSync } from "node:child_process";
import {
  copyFileSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { basename, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const temporaryRoot = mkdtempSync(join(tmpdir(), "ai-cdl-consumer-"));
const tarballDirectory = join(temporaryRoot, "tarballs");
mkdirSync(tarballDirectory);

function run(command, args, cwd = root) {
  return execFileSync(command, args, { cwd, encoding: "utf8", stdio: "pipe" }).trim();
}

const publicPackages = readdirSync(join(root, "packages"), { withFileTypes: true })
  .filter((entry) => entry.isDirectory())
  .map((entry) => ({
    directory: entry.name,
    manifest: JSON.parse(readFileSync(join(root, "packages", entry.name, "package.json"), "utf8")),
  }))
  .filter(({ manifest }) => manifest.private !== true)
  .sort((left, right) => left.manifest.name.localeCompare(right.manifest.name));

for (const { manifest } of publicPackages) {
  run("pnpm", ["--filter", manifest.name, "pack", "--pack-destination", tarballDirectory]);
}
const tarballs = readdirSync(tarballDirectory)
  .filter((file) => file.endsWith(".tgz"))
  .map((file) => join(tarballDirectory, file));
if (tarballs.length !== publicPackages.length) {
  throw new Error(`Expected ${publicPackages.length} tarballs, found ${tarballs.length}.`);
}
const tarballByPackage = new Map(
  publicPackages.map(({ manifest }) => {
    const filename = `${manifest.name.replace(/^@/, "").replace("/", "-")}-${manifest.version}.tgz`;
    const tarball = tarballs.find((candidate) => candidate.endsWith(filename));
    if (!tarball) throw new Error(`Missing tarball for ${manifest.name}.`);
    return [manifest.name, tarball];
  }),
);

const minimalDirectory = join(temporaryRoot, "minimal-cli-consumer");
mkdirSync(minimalDirectory);
writeFileSync(
  join(minimalDirectory, "package.json"),
  `${JSON.stringify({ name: "ai-cdl-minimal-smoke", private: true, type: "module" }, null, 2)}\n`,
);
const cliTarball = tarballs.find((file) => file.includes("ai-cdl-cli-"));
if (!cliTarball) throw new Error("Missing @ai-cdl/cli tarball.");
run(
  "npm",
  ["install", "--ignore-scripts", "--no-audit", "--no-fund", cliTarball],
  minimalDirectory,
);
const minimalTree = JSON.parse(run("npm", ["ls", "--all", "--json"], minimalDirectory));
const serializedTree = JSON.stringify(minimalTree);
for (const forbidden of [
  "@ai-cdl/agent-http",
  "@ai-cdl/evals",
  "@ai-cdl/testing",
  "office-addin-debugging",
  "office-addin-dev-certs",
  "office-addin-manifest",
]) {
  if (serializedTree.includes(forbidden)) {
    throw new Error(`Minimal CLI unexpectedly installs ${forbidden}.`);
  }
}
function dependencyCount(node) {
  return Object.values(node.dependencies ?? {}).reduce(
    (total, dependency) => total + 1 + dependencyCount(dependency),
    0,
  );
}
const minimalDependencyCount = dependencyCount(minimalTree);
if (minimalDependencyCount > 75) {
  throw new Error(`Minimal CLI dependency tree regressed to ${minimalDependencyCount} packages.`);
}
run(process.execPath, [join(minimalDirectory, "node_modules/@ai-cdl/cli/dist/cli.js"), "--help"]);

const singlePackageDirectory = join(temporaryRoot, "single-package-consumer");
const vendorDirectory = join(singlePackageDirectory, "vendor");
mkdirSync(vendorDirectory, { recursive: true });
for (const tarball of tarballs) copyFileSync(tarball, join(vendorDirectory, basename(tarball)));
const pairTarball = tarballByPackage.get("@ai-cdl/pair");
if (!pairTarball) throw new Error("Missing @ai-cdl/pair tarball.");
const overrides = Object.fromEntries(
  [...tarballByPackage].map(([name, tarball]) => [name, `file:./vendor/${basename(tarball)}`]),
);
writeFileSync(
  join(singlePackageDirectory, "package.json"),
  `${JSON.stringify(
    {
      name: "ai-cdl-single-tarball-smoke",
      private: true,
      type: "module",
      dependencies: {
        "@ai-cdl/pair": `file:./vendor/${basename(pairTarball)}`,
      },
    },
    null,
    2,
  )}\n`,
);
writeFileSync(
  join(singlePackageDirectory, "pnpm-workspace.yaml"),
  `packages:\n  - .\noverrides:\n${Object.entries(overrides)
    .map(([name, tarball]) => `  '${name}': ${tarball}`)
    .join("\n")}\n`,
);
run("pnpm", ["install", "--ignore-scripts"], singlePackageDirectory);
writeFileSync(
  join(singlePackageDirectory, "single.mjs"),
  `import { createPairClient } from "@ai-cdl/pair";\nif (typeof createPairClient !== "function") throw new Error("Missing single-package export");\n`,
);
run(process.execPath, ["single.mjs"], singlePackageDirectory);
const singleLockfile = readFileSync(join(singlePackageDirectory, "pnpm-lock.yaml"), "utf8");
for (const name of ["core", "excel", "addin-core", "protocol", "transport"]) {
  if (!singleLockfile.includes(`vendor/ai-cdl-${name}-`)) {
    throw new Error(`Single-package install did not resolve @ai-cdl/${name} from a local tarball.`);
  }
}

const consumerDirectory = join(temporaryRoot, "all-packages-consumer");
mkdirSync(consumerDirectory);
writeFileSync(
  join(consumerDirectory, "package.json"),
  `${JSON.stringify({ name: "ai-cdl-external-smoke", private: true, type: "module" }, null, 2)}\n`,
);
run(
  "npm",
  ["install", "--ignore-scripts", "--no-audit", "--no-fund", ...tarballs],
  consumerDirectory,
);
writeFileSync(
  join(consumerDirectory, "esm.mjs"),
  `import { createAgentSession } from "@ai-cdl/core";\nimport { HttpAgentAdapter } from "@ai-cdl/agent-http";\nimport { createExcelTools } from "@ai-cdl/excel";\nimport { runAdapterContract } from "@ai-cdl/testing";\nimport { defineConfig } from "@ai-cdl/cli/config";\nimport { createPairClient } from "@ai-cdl/pair";\nimport { createAddinController } from "@ai-cdl/addin-core";\nfor (const value of [createAgentSession, HttpAgentAdapter, createExcelTools, runAdapterContract, defineConfig, createPairClient, createAddinController]) {\n  if (typeof value !== "function") throw new Error("Missing ESM export");\n}\n`,
);
writeFileSync(
  join(consumerDirectory, "cjs.cjs"),
  `for (const name of ["core", "agent-http", "excel", "testing", "protocol", "transport", "addin-core", "pair", "pair-cli"]) {\n  const loaded = require("@ai-cdl/" + name);\n  if (!loaded || typeof loaded !== "object") throw new Error("Missing CJS export: " + name);\n}\nconst config = require("@ai-cdl/cli/config");\nif (typeof config.defineConfig !== "function") throw new Error("Missing CJS config export");\n`,
);
writeFileSync(
  join(consumerDirectory, "types.ts"),
  `import type { AgentAdapter } from "@ai-cdl/core";\nimport type { HttpAgentAdapterOptions } from "@ai-cdl/agent-http";\nimport { createPairClient } from "@ai-cdl/pair";\nconst adapter: AgentAdapter | undefined = undefined;\nconst http: HttpAgentAdapterOptions = { endpoint: "https://example.com" };\nconst pairOptions: Parameters<typeof createPairClient>[0] | undefined = undefined;\nvoid [adapter, http, pairOptions];\n`,
);
writeFileSync(
  join(consumerDirectory, "tsconfig.json"),
  `${JSON.stringify(
    {
      compilerOptions: {
        strict: true,
        target: "ES2022",
        module: "NodeNext",
        moduleResolution: "NodeNext",
        skipLibCheck: true,
        noEmit: true,
      },
      include: ["types.ts"],
    },
    null,
    2,
  )}\n`,
);
run(process.execPath, ["esm.mjs"], consumerDirectory);
run(process.execPath, ["cjs.cjs"], consumerDirectory);
run(join(root, "node_modules/.bin/tsc"), ["--project", "tsconfig.json"], consumerDirectory);
for (const binary of ["ai-cdl", "ai-cdl-contract", "ai-cdl-pair-agent"]) {
  run(join(consumerDirectory, `node_modules/.bin/${binary}`), ["--help"], consumerDirectory);
}

const adapterDist = join(consumerDirectory, "node_modules/@ai-cdl/pair-cli/dist");
for (const asset of [
  "skill/SKILL.md",
  "skill/scripts/session.mjs",
  "skill/scripts/lib/encrypted-socket.mjs",
]) {
  if (!readFileSync(join(adapterDist, asset), "utf8").length)
    throw new Error(`Missing adapter asset: ${asset}`);
}
const emptyProfiles = spawnSync(process.execPath, [join(adapterDist, "agent.js"), "list"], {
  cwd: consumerDirectory,
  encoding: "utf8",
  env: { ...process.env, AI_CDL_PAIR_HOME: join(temporaryRoot, "empty-profiles") },
});
if (emptyProfiles.status !== 0 || JSON.parse(emptyProfiles.stdout).terminals.length !== 0)
  throw new Error("Packed adapter cannot list profiles.");

// The existing Pair CLI reports usage with status 1 when no subcommand is supplied.
const pairUsage = spawnSync(join(consumerDirectory, "node_modules/.bin/ai-cdl-pair"), [], {
  cwd: consumerDirectory,
  encoding: "utf8",
});
if (pairUsage.status !== 1 || !pairUsage.stderr.includes("Usage: ai-cdl-pair relay")) {
  throw new Error("Packed Pair CLI did not expose its usage contract.");
}
writeFileSync(
  join(consumerDirectory, "pair.mjs"),
  `import assert from "node:assert/strict";
import { createPairClient, createPairAddinSession } from "@ai-cdl/pair";
import { createInMemoryTransportPair } from "@ai-cdl/transport";
import { createAddinController } from "@ai-cdl/addin-core";
import { InMemoryExcelAdapter } from "@ai-cdl/excel";
const [addinTransport, agentTransport] = createInMemoryTransportPair();
const adapter = new InMemoryExcelAdapter({ sheets: [{ name: "Sheet1", values: [[10]] }] });
let approvals = 0;
const controller = createAddinController({
  adapter,
  workbook: { id: "synthetic", name: "Synthetic.xlsx" },
  requestApproval: () => { approvals += 1; return true; },
});
const addin = createPairAddinSession({ transport: addinTransport, controller });
const client = createPairClient({ transport: agentTransport });
await addin.start();
await client.connect();
try {
  const range = { sheetId: "Sheet1", address: "A1" };
  assert.deepEqual((await client.request("excel.range.read", { range })).values, [[10]]);
  assert.equal(approvals, 0);
  await client.request("excel.range.write", { range, values: [[25]] });
  assert.equal(approvals, 1);
  assert.deepEqual((await client.request("excel.range.read", { range })).values, [[25]]);
} finally {
  await client.close();
  await addin.stop();
}
`,
);
run(process.execPath, ["pair.mjs"], consumerDirectory);

console.log(
  `PASS external consumer: ${publicPackages.length} packed packages, one-package local override install, ESM, CJS, types, binaries, Pair read/write approval round trip; minimal CLI ${minimalDependencyCount} dependencies.`,
);
rmSync(temporaryRoot, { recursive: true, force: true });
