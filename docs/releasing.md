# Releasing & Deployment

One tracked flow for publishing the npm packages and deploying the add-in.
For daily local development & testing see
[local-development.md](./local-development.md).

## Component map: what ships where

| Component | Package name | Ships via | Current |
| ----------- | -------------- | ----------- | --------- |
| Shared wire protocol | `@dieulc/pi-office-protocol` | npm | 0.1.0 |
| Pi bridge extension | `@dieulc/pi-office-bridge` | npm | 0.1.0 |
| CORS proxy | `pi-for-office-proxy` | npm | 0.2.6-pre |
| Python bridge | `pi-for-office-python-bridge` | npm | 0.1.2 |
| Tmux bridge | `pi-for-office-tmux-bridge` | npm | 0.1.2 |
| **Add-in** | *(not npm-published)* | GitHub Pages | `dist/` of `packages/add-in` |

The add-in is a static web app: it deploys to GitHub Pages from `main` and is
sideloaded into Office via a manifest. It is never published to npm.

## Version policy

- Each npm package versions **independently**; bump only what changed.
- Scoped packages (`@dieulc/*`) must publish with `--access public` (their
  `publishConfig.access` already says `public`).
- Depends on the shared protocol → bump/publish **protocol first**, then
  `@dieulc/pi-office-bridge` (it depends on `@dieulc/pi-office-protocol`).
- A `-pre` suffix means prerelease: it does **not** serve as the `latest`
  dist-tag by default. For a real user-facing release, publish a non-prelude
  version.
- The add-in version is tracked separately (see
  `npm run check:version-sync` in `packages/add-in`).

## The release script

```bash
npm run release:status     # read-only: versions + npm auth + published state
npm run release:check      # typecheck + bridge tests + pack --dry-run (validates prepack sync)
npm run release:publish    # bump + publish in dependency order, with checks & confirmation
```

All commands delegate to `scripts/release.mjs`:

- `status` — prints your npm identity and a table of local vs published
  versions for all 5 packages. Safe to run any time.
- `check` — runs the root typecheck, the bridge-extension test suite, then
  `npm pack --dry-run` for every publishable package (this also exercises each
  package's `prepack` sync, which copies server scripts from
  `packages/add-in/scripts/` into `pkg/<name>/scripts/` before packing).
- `publish` — optional `--bump patch|minor|major`, then publishes in
  dependency order, **skipping versions already published**, and verifies each
  with `npm view` afterwards.

Safety guardrails built into `publish`:

- refuses when not logged in (`npm whoami`),
- refuses on a dirty working tree unless `--allow-dirty`,
- runs `check` first unless `--skip-checks`,
- asks for confirmation unless `--yes`,
- never publishes a version that already exists.

### publish options

```bash
npm run release:publish                     # everything, asks for confirmation
npm run release:publish -- --bump patch     # bump all publishables then publish
npm run release:publish -- --only @dieulc/pi-office-bridge   # one package
npm run release:publish -- --only pi-for-office-python-bridge --bump minor
npm run release:publish -- --tag next       # custom dist-tag
npm run release:publish -- --otp 123456     # 2FA code (or pass it when prompted)
```

## Manual fallback (if you ever want to skip the script)

Publish order matters:

```bash
# 1. protocol first (scoped → public)
cd packages/protocol
npm version patch --no-git-tag-version
npm publish --access public

# 2. bridge extension (scoped → public)
cd ../../bridge-extension
npm version patch --no-git-tag-version
npm publish --access public

# 3. the three pkg packages (prepack syncs their scripts automatically)
cd ../add-in/pkg/proxy
npm version patch --no-git-tag-version
npm publish

cd ../python-bridge
npm version patch --no-git-tag-version
npm publish

cd ../tmux-bridge
npm version patch --no-git-tag-version
npm publish
```

## Deploying the add-in (GitHub Pages)

The add-in deploys **automatically** from `.github/workflows/deploy-pages.yml`:

- **Trigger:** push to `main` (or manual workflow dispatch).
- **What it does:** builds `packages/add-in` with Vite
  (`VITE_BASE_PATH=/pi-for-office/`), regenerates `manifest.prod.xml` with
  `ADDIN_BASE_URL=https://dieuluucanh.github.io/pi-for-office`, publishes
  `dist/` to GitHub Pages.
- **Result:** the hosted add-in is available at
  `https://dieuluucanh.github.io/pi-for-office/src/taskpane.html`, and the
  production manifest lives at
  `https://dieuluucanh.github.io/pi-for-office/manifest.prod.xml`.

So a typical add-in release is: merge to `main` → workflow runs → verify the
hosted URL + manifest.

Details for fork/custom-domain/vercel setups live in:

- `packages/add-in/docs/deploy-github-pages.md`
- `packages/add-in/docs/deploy-vercel.md`
- `packages/add-in/docs/central-proxy.md` (org-hosted CORS proxy)

## Post-publish verification checklist

Run this every time you publish:

```bash
# 1. The npm packages resolve:
npm view @dieulc/pi-office-bridge version
npm view pi-for-office-python-bridge version
npm view pi-for-office-tmux-bridge version

# 2. The npx one-liners the add-in's setup cards show actually work:
npx pi-for-office-python-bridge@latest     # → https://localhost:3340  (Ctrl+C to stop)
npx pi-for-office-tmux-bridge@latest       # → https://localhost:3341  (Ctrl+C to stop)

# 3. Bridge extension installs & serves health:
pi install npm:@dieulc/pi-office-bridge
curl http://127.0.0.1:38617/health        # with a Pi process running the extension

# 4. Hosted add-in is current:
curl -fsSL https://dieuluucanh.github.io/pi-for-office/manifest.prod.xml | head
```

If a clean machine can't resolve `@dieulc/pi-office-protocol` (raw TS) or
`@earendil-works/pi-ai` after `pi install npm:@dieulc/pi-office-bridge`,
see the Risks section in the plan that introduced this flow — the contingency
is bundling the protocol source into the bridge package.

## Rolling back a bad publish

npm versions are immutable — you cannot unpublish a version others may have
installed. Use deprecation instead:

```bash
npm deprecate @dieulc/pi-office-bridge@0.1.0 "broken — use 0.1.1"
```

Then publish a fixed patch version.

## Release checklist (quick)

- [ ] `npm run release:status` — confirm who you are + current versions
- [ ] `npm run release:check` — typecheck, bridge tests, pack dry-runs all green
- [ ] Merge add-in changes to `main` (Pages deploy) if the add-in changed
- [ ] `npm run release:publish -- --bump patch` (or `--only <pkg>` for a subset)
- [ ] Run the post-publish verification checklist above
- [ ] Update release notes (`packages/add-in/docs/release-notes/`) if add-in-facing

## Related

- [local-development.md](./local-development.md)
- `packages/bridge-extension/README.md`
- `packages/add-in/docs/install.md` — sideloading the hosted add-in
