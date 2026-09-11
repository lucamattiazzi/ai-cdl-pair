# Publishing Pair

The Excel add-in and npm libraries are separate artifacts. This checkout prepares both; it has
not been submitted to AppSource, uploaded to a registry, or assigned a remote Git repository.

## Libraries

`packages/*` contains ten public packages with ESM/CommonJS entry points, TypeScript declarations,
and Apache-2.0 licenses. `@ai-cdl/pair` is the client/session library; `@ai-cdl/pair-cli` provides
the local relay plus `ai-cdl-pair-agent`, its bundled encrypted bridge, MCP tools and native
harness adapters. See [v1 readiness](v1-readiness.md) before release. The root workspace and the two applications remain private npm packages.

Before release, choose the release owner for the retained `@ai-cdl/*` names, confirm npm scope
access, and set repository metadata to the actual new public repository. Versions remain `0.1.1`
from the source revision; check registry state before choosing the release version.

```sh
pnpm build
pnpm smoke:packages
pnpm smoke:consumer
pnpm changeset
pnpm changeset version
pnpm build
pnpm smoke:consumer
```

Review the generated version changes and release notes. Changesets groups the ten libraries into
one fixed release set. Run `pnpm exec changeset publish` only when ready to publish that release.
There is deliberately no workflow that publishes on push.

To inspect a single package without publishing:

```sh
pnpm --filter @ai-cdl/pair pack --pack-destination ../../artifacts/npm
```

## Add-in and relay

The app builds to `apps/pair-addin/dist`; the bundled relay builds to `apps/pair-server/dist`.
Choose the real public HTTPS origin, then generate and build:

```sh
PAIR_PUBLIC_ORIGIN=https://pair.example.com pnpm manifest:production
pnpm pair:build
```

Replace `pair.example.com` with the deployment domain. The generated manifest is copied into the
add-in build. See [Docker/Caddy deployment](../deploy/pair/README.md).

The development manifest remains useful for local sideloading; restore it with
`pnpm manifest:generate` after preparing production artifacts.

Before an AppSource submission, validate the manifest and complete the real Excel/harness checks
in [the testing guide](pair-testing.md). Support/setup/data-handling pages and sized icons are included; see the [asset inventory](marketplace-assets.md).
Production hosting, publisher contact/legal details, real Excel screenshots and Microsoft review
remain separate publication work.
