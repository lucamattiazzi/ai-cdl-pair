# @ai-cdl/transport

## 0.2.0-beta.0

### Minor Changes

- 5c48cc5: Add authenticated end-to-end encrypted socket helpers and stable Pair identities, keeping encrypted
  framing separate from workbook RPC. Support correctly sized Office manifest icons and same-origin
  support pages. The Pair application now uses remembered terminal connections and disables its
  legacy plaintext hosted routes by default.

  Update ws to 8.21.0 and the js-yaml lockfile resolution to patched releases.

### Patch Changes

- Updated dependencies [5c48cc5]
  - @ai-cdl/protocol@0.2.0-beta.0
