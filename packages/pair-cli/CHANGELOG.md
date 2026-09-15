# @ai-cdl/pair-cli

## 0.2.0-beta.0

### Minor Changes

- 5c48cc5: Add native Codex, OpenCode and Claude Code channel adapters with remembered conversation bindings,
  shared Excel MCP tools/documentation, and local metadata-first trace subscriptions. Add bounded,
  approved chart creation/listing and expose operation undo availability to the TaskPane. Bundle
  the portable bridge in the CLI distribution. Claude Channels remains a research preview.

  Start an authenticated local OpenCode subprocess automatically, while keeping the hosted Pair
  service configurable for independent self-hosting. End users only launch the harness adapter.

### Patch Changes

- 5c48cc5: Add authenticated end-to-end encrypted socket helpers and stable Pair identities, keeping encrypted
  framing separate from workbook RPC. Support correctly sized Office manifest icons and same-origin
  support pages. The Pair application now uses remembered terminal connections and disables its
  legacy plaintext hosted routes by default.

  Update ws to 8.21.0 and the js-yaml lockfile resolution to patched releases.

- Updated dependencies [5c48cc5]
  - @ai-cdl/protocol@0.2.0-beta.0
