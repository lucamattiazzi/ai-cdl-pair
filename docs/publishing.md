# Publishing Pair

The Excel add-in and npm libraries are separate artifacts. The source repository is
[ai-cdl-pair](https://github.com/lucamattiazzi/ai-cdl-pair). The first independent release is
`0.2.0-beta.0` on the npm `beta` channel; check registry availability when executing the release.
Microsoft Marketplace submission and production hosting are separate steps.

## Libraries

`packages/*` contains ten public packages with ESM/CommonJS entry points, TypeScript declarations,
and Apache-2.0 licenses. `@ai-cdl/pair` is the client/session library; `@ai-cdl/pair-cli` provides
the local relay plus `ai-cdl-pair-agent`, its bundled encrypted bridge, MCP tools and native
harness adapters. See [v1 readiness](v1-readiness.md) before release. The root workspace and the two applications remain private npm packages.

The `lucamattiazzi` account owns the `@ai-cdl` npm organization (verified 2026-09-15).
Repository metadata points to this checkout. The shared packages already published as `0.1.1`
keep their stable `latest` tags; prereleases must use `beta`. Do not independently release the
same package/version from the original AI-CDL repository.

```sh
pnpm build
pnpm smoke:packages
pnpm smoke:consumer
# Already in beta prerelease mode; add a Changeset for the next change.
pnpm changeset
pnpm changeset version
pnpm build
pnpm smoke:consumer
```

Review the generated version changes and release notes. Changesets groups the ten libraries into
one fixed release set. Run `pnpm release:beta` when ready to publish a beta release. This sets the `beta` tag explicitly
for every package, including packages with no previous npm release; Changesets prerelease
publishing would otherwise put those new packages on `latest`. Complete npm two-factor
authentication in your terminal when prompted. To inspect the upload without publishing, run
`pnpm release:beta --dry-run`.
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
