# Deploy hosted build on GitHub Pages (free, no Vercel account needed)

Pi for Office's taskpane is a static site built by Vite (`packages/add-in/dist/`).

GitHub Pages hosts it for free at the project-site subpath:

```text
https://dieuluucanh.github.io/pi-for-office/
```

## How it works

- `.github/workflows/deploy-pages.yml` (repo root) builds `packages/add-in`
  with `VITE_BASE_PATH=/pi-for-office/` so every asset URL resolves under the
  subpath, then publishes the `dist/` folder to Pages.
- The workflow regenerates `manifest.prod.xml` from the dev `manifest.xml`
  with `ADDIN_BASE_URL=https://dieuluucanh.github.io/pi-for-office` before
  building, so the published `/manifest.prod.xml` always matches the deployed
  taskpane.
- `public/.nojekyll` stops Jekyll from interfering with the static files.

## One-time setup (repo owner, ~1 minute)

1. Create a GitHub repo from this project (it is public already:
   `dieuluucanh/pi-for-office`).
2. Push `main`.
3. In the GitHub web UI: repo → **Settings → Pages**.
4. Under **Build and deployment**, set **Source** to **GitHub Actions**.
5. (Optional) Under **Custom domain**, leave empty — the default
   `dieuluucanh.github.io/pi-for-office` URL is used by the manifest.

After that, every push to `main` (and every manual **Run workflow** run)
rebuilds and redeploys automatically.

## If you fork or rename the repo

The URL changes with the owner/repo name. Update all of these together:

- `ADDIN_BASE_URL` in `.github/workflows/deploy-pages.yml`
- the `VITE_BASE_PATH` env value in the same workflow
- regenerate the local manifest:
  `cd packages/add-in && ADDIN_BASE_URL="https://<owner>.github.io/<repo>" npm run manifest:prod`
- the download links in `README.md`, `docs/install.md`, and `public/index.html`

## Side-loading the hosted add-in

Users don't need a dev server. They download the manifest once:

```text
https://dieuluucanh.github.io/pi-for-office/manifest.prod.xml
```

then sideload it into Excel / Word / PowerPoint (Insert → My Add-ins → Upload
My Add-in…, or the shared-folder catalog). The add-in loads the taskpane from
the GitHub Pages URL, so updates deploy automatically.

## Local production-preview alternative

No hosting at all for a quick check: `npm run serve:dist` serves the built
`dist/` over HTTPS on localhost using the dev manifest (`manifest.xml`).
