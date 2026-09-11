# Pair v1 readiness — 2026-09-11

The requested v1 features are implemented in this checkout. This is **not yet a verified
AppSource release**: real Excel validation, interactive Claude channel validation, the OpenCode
provider configuration and publication ownership/hosting still need attention.

## Feature inventory

| Requirement | Implementation | Verification / limit |
| --- | --- | --- |
| Agent ↔ TaskPane | Codex App Server, OpenCode HTTP adapter, Claude MCP channel | Real Codex tool call + thread resume; real OpenCode MCP discovery; real stdio channel contract |
| Hosted and self-hosted service | Operator configures `PAIR_DOMAIN`; pairing remembers the chosen instance | End users launch only the adapter; no fixed public service domain |
| Reads and writes | Shared MCP tools generated from RPC descriptors | Encrypted integration with synthetic workbook, preview/approval/readback |
| Charts | List and create column/bar/line/pie/scatter from a bounded source | Controller approval, duplicate name protection, Office.js mock and browser approval; no chart undo |
| Excel guidance | `excel_guide`, `pair://docs/excel`, portable skill and RPC reference | MCP discovery/calls and installed skill smoke |
| Remembered connections | Private local profile + native thread/session ID; TaskPane saved list and optional reconnect-on-open | Profile permissions, duplicate-consumer exclusion, actual Codex resume and browser reload |
| Assets and usable UI | Icons, manifest, setup/privacy/support pages; direct one-time URL setup | Asset HTTP checks, keyboard approval dialog and 320 px browser screenshots |
| Evals / traces | Public `subscribeBridge`; metadata-first JSONL CLI, opt-in full payloads | Independent subscribers, RPC outcome/duration, chat routing IDs; no provider reasoning/token traces |
| Relay confidentiality | Existing authenticated E2EE channel, preserved through local MCP | Encrypted round trip, plaintext/takeover rejection, no key in traces |

## Checks completed

- Initial focused Vitest: **42 tests passed**, plus one opt-in native smoke skipped in normal runs.
- Hosting refinement: **14 focused tests passed**, including automatic OpenCode startup, private API
  authentication setup, cancellation/cleanup, saved profiles and encrypted bridge reconnect.
- Real OpenCode startup now runs through the adapter's managed subprocess: unauthenticated HTTP
  rejected, MCP loaded, process restarted on a fresh port and the same session resumed. No model
  call was made for this refinement. Scoped CLI build, TypeScript and Biome checks passed.
- Opt-in native smoke passed with **Codex 0.153.4**: actual `excel_guide` MCP call and completed model
  turn, then App Server restart and resume of the same thread. Synthetic thread archived afterward.
- **OpenCode 1.18.29**: real local server, MCP registration and session creation passed. A real model
  turn failed with **`APIError, HTTP 404`** from its configured provider. The adapter surfaces this
  failure and does not replay the request. Its HTTP behavior is also covered with controlled tests.
- **Claude Code 2.1.236** is installed and documents Channels. The channel implementation was tested
  as a real child process over MCP stdio: initialization, incoming event and `pair_reply`. An actual
  interactive Claude session with channel consent/model execution was **not** completed.
- Browser: pair, encrypted RPC/chat, approve write and chart, verify chart list, hide unsupported
  chart undo, manual/automatic reconnect, reset approval waiver, forget, and all help/icon/manifest
  assets. No page errors or horizontal overflow at 320 px.
- Scoped TypeScript and Biome checks; complete distributable build and manifest consistency check.
- Ten packages packed and installed into external npm consumers: ESM/CommonJS/types, binaries,
  included skill/crypto assets and a workbook approval round trip. No npm publication performed.
- Four skill installer targets tested in isolated directories; repository isolation checks passed.
- Dependency audit after adding the MCP SDK: **zero reported vulnerabilities**.

Reproduce from the repository root:

```sh
pnpm install --frozen-lockfile
pnpm build
pnpm exec vitest run packages/protocol/src packages/addin-core/src packages/excel/src/office-adapter.test.ts packages/pair-cli/src apps/pair-addin/src/terminals.test.ts apps/pair-server/src/encrypted.test.ts
pnpm test:browser
pnpm test:harness
pnpm manifest:check
pnpm smoke:consumer
# Starts installed harnesses; creates/archives a synthetic Codex thread, no model call:
PAIR_NATIVE_SMOKE=1 pnpm exec vitest run packages/pair-cli/src/native-smoke.test.ts
# Also makes a real model call with the user's configured Codex account:
PAIR_NATIVE_SMOKE=1 PAIR_NATIVE_TURN=codex pnpm exec vitest run packages/pair-cli/src/native-smoke.test.ts
```

`PAIR_NATIVE_TURN=opencode` tests its model; `1` tests both. These are deliberately opt-in and are
not run by CI. Native model tests require the relevant provider/account to work. Browser screenshots
are generated under ignored `artifacts/pair-review/`; they show synthetic data, not real Excel.

## Before publishing

1. **Run in Microsoft Excel.** Excel was not found in `/Applications` on this machine. Sideload the
   manifest and verify read/write/formula preservation, chart creation, approval rejection, reload,
   disconnect during a write and reconnection. Repeat on every host you claim to support. The
   Office.js mock cannot establish host compatibility; Windows pipe ACLs remain unvalidated.
2. **Complete the two harness checks.** Fix the OpenCode provider's 404 and run a real workbook task.
   Start Claude through the launcher, accept channel consent and verify a round trip. Keep Claude
   marked preview: custom Channels still require the development flag and may be blocked by policy.
3. **Set publication identity and hosting.** This checkout currently has no Git remote. Confirm
   ownership of the retained `@ai-cdl/*` npm scope, set repository metadata and configure the operator's
   HTTPS domain through `PAIR_DOMAIN` (see [self-hosting](../deploy/pair/README.md)),
   and apply the included Changeset before packing the chosen release version. Nothing was pushed,
   published, deployed or submitted to Microsoft during this work.
4. **Finish Marketplace material.** Publisher/support/legal details and genuine Excel screenshots
   remain necessary. See [asset inventory](marketplace-assets.md) and [publishing](publishing.md).

For a release in the next few days, keep this feature scope. Defer WebRTC, chart editing/undo,
formatting/table creation in RPC 0.2, cross-machine native harness control and durable message
queues. The current E2EE WebSocket transport already meets relay payload confidentiality, provided
the client code is trusted. It does not provide forward secrecy or protect against malicious
replacement of TaskPane JavaScript; see [the security model](secure-pairing.md).
