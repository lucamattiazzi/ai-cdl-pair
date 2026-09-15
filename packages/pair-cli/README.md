# @ai-cdl/pair-cli

`ai-cdl-pair-agent` connects Codex, OpenCode or Claude Code to the Excel Pair TaskPane. Requires
Node.js 22.12+ and a signed-in harness. The included skill, bridge and crypto are self-contained.

```sh
npm install -g @ai-cdl/pair-cli@beta
ai-cdl-pair-agent pair --name desk
ai-cdl-pair-agent codex --name desk
# Or: opencode --name desk
# Or: claude --name desk (Channels research preview; local consent required)
ai-cdl-pair-agent list
ai-cdl-pair-agent trace --name desk --out run.jsonl
```

Pair asks for the private TaskPane URL once; subsequent starts use the saved profile and native
conversation, including the selected hosted or self-hosted Pair instance. End users do not start
a Pair relay. OpenCode's local harness API starts automatically on an authenticated loopback port
and stops with the adapter. Optional `--server` attaches an API you already manage instead.
Claude must stay open. The Codex adapter
starts App Server, not an arbitrary existing TUI. Native Codex command/file approvals are declined;
Excel writes/charts retain their TaskPane approval flow.

Public exports `bridgeCommand` and `subscribeBridge` allow local tool/eval integrations. Trace
payloads are excluded by default; `includeContent: true` opts in. Never commit full traces or saved
profiles. A timeout never causes mutation replay. Slow trace listeners cannot block the workbook.

`ai-cdl-pair relay --port 4310` remains the legacy **plaintext loopback relay**; the native adapters
require encrypted `/connect` profiles from the hosted Pair relay instead.
