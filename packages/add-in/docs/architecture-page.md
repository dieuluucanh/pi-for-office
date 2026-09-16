# Maintaining the architecture page

The public architecture page lives at
[`packages/add-in/public/architecture/index.html`](../public/architecture/index.html)
and is served from the GitHub Pages project subpath:

- <https://dieuluucanh.github.io/pi-for-office/architecture/>

It is hand-written HTML (inline CSS + a small TOC/reveal script) and is copied
into `dist/` by Vite like any other file in `public/`. It is not generated and
has no build step — but it **is** guarded against drift.

## The fact contract

Every number or port the page asserts is wrapped in a marker:

```html
<span data-fact="tools.core">17</span> core tools
```

Rules:

- The `data-fact` key must be unique per fact, but the **same key may appear
  several times** (for example in the Overview pillars and again in the owning
  section). All occurrences must show the same value.
- The element's inner text is the value. Markers hold only the value itself —
  units and words live outside the span (`<span data-fact="budget.toolOutputKb">8 KB</span>`
  is fine because the guard derives `8 KB` too).
- Never print a fact in prose without a marker. If a number is worth writing
  down, it is worth guarding.

## The guard

`npm run check:architecture-facts` runs
[`tests/architecture-page-facts.test.ts`](../tests/architecture-page-facts.test.ts),
which:

1. derives every expected value from the code (constants imported or parsed
   from their source of truth — `src/tools/names.ts`, `src/extensions/permissions.ts`,
   `src/ui/overlay-ids.ts`, `src/config/local-service-defaults.ts`,
   `src/workbook/recovery/constants.ts`, `src/compaction/defaults.ts`, the test
   directory, `package.json`, …);
2. extracts every `data-fact` marker from the page;
3. fails on a missing marker, an unknown marker, or a value that disagrees
   with the code;
4. rejects retired terminology (`pi-web-ui`, stale version strings, removed
   credits, `href="/…"` root-absolute links);
5. checks that every same-page `#anchor` resolves and that the nav links stay
   relative (`../`, `../#install`, `../#connect`, `./`);
6. checks that the documented keyboard-shortcut labels still match the
   matcher names in `src/taskpane/keyboard-shortcuts/matchers.ts`.

It is part of `npm run check`, so it runs in the pre-commit hook and in the
[Checks workflow](../../../.github/workflows/checks.yml).

## Adding or changing a fact

1. Change the code.
2. Run `npm run check:architecture-facts` and read the failure — it prints the
   expected value.
3. If the fact is new, add both the marker on the page and an entry in
   `deriveExpectedFacts()` in the guard.
4. Re-run until green.

If a *source shape* changes (a constant is renamed, a list moves files) the
guard throws "could not derive …" instead of silently passing. Move the
expectation to the new source of truth rather than deleting the check.

## Conventions for the page

- **No version strings.** The page documents current `main`; a version in the
  hero would be stale within a release. `package.json` and
  `src/app/metadata.ts` own the version.
- **No source line counts.** They drift on every edit and say nothing useful.
- **No npm package versions** in prose; the release script and npm own those.
- Keep the visual system: `.sec-head` sections, `.card` / `.tool-table` /
  `.spec-list` / `.chip` components, and the sticky TOC. New sections need a
  `#anchor` registered in the TOC (the guard enforces that anchors resolve).
- The footer mirrors `public/index.html` (the landing page) verbatim.

## Deploying a page change

The GitHub Pages workflow (`.github/workflows/deploy-pages.yml`) publishes the
built `dist/` on every push to `main`, so a page edit ships with the next
deploy — no separate step. Links in the page must stay relative or use the
project subpath; `href="/…"` would escape `/pi-for-office/` and 404.
