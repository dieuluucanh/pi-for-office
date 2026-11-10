# UI layout, dialogs and notifications

> Status: implemented 2026-11-10. This file records the rules the overlay,
> dialog and notice CSS depends on. Changing any of them needs a matching change
> here and in the tests listed at the end.

## Why this exists

Two user-visible defects drove the rules:

1. **A toast rendered one glyph per line.** `toasts.css` positioned the single
   toast with `position: fixed; left: 50%; transform: translateX(-50%)` and
   `width: auto` (shrink-to-fit). For an abspos box with `left: 50%` the used
   width can only grow to the distance from `left` to the container's right
   edge — half the pane (~183px in a ~360px taskpane). The content was
   `inline-flex`, the action button `white-space: nowrap` (133px), and the
   message `overflow-wrap: anywhere`; the message was therefore squeezed to a
   12px column and broke at every character, growing the toast to 752px tall.
2. **Dialogs ignored the pane.** Cards were capped at fixed px widths
   (`min(380/520/720px, 92vw)`) and a `vh` ceiling, the resume/settings bodies
   nested scroll regions, and a `@media (max-height: 520px)` override in
   `components/welcome.css` was dead because `components.css` is imported
   *before* `overlays.css`, so the later, equal-specificity
   `max-height: min(80vh, 720px)` always won.

## Geometry tokens (`src/ui/theme/tokens.css`)

| Token | Value | Meaning |
|---|---|---|
| `--pi-dialog-gutter` | `8px` | overlay padding; the card never touches the pane edge |
| `--pi-dialog-gap` | `10px` | inner stack gap used by dialog bodies |
| `--pi-dialog-w-s` | `480px` | confirm / text-input cards |
| `--pi-dialog-w-m` | `720px` | resume history, files workspace |
| `--pi-dialog-w-l` | `960px` | settings shell and its pages |
| `--pi-dialog-w-prompt` | `420px` | welcome/login, model selector, API key, provider prompt |
| `--pi-toast-offset` | `56px` | toast host top anchor (below the tab strip) |
| `--pi-toast-max` | `480px` | readable toast measure on wide panes |
| `--pi-z-overlay` / `--pi-z-popover` / `--pi-z-toast` | `200` / `220` / `300` | toasts must sit above dialogs (a toast can be raised *by* a dialog action) |

Cards are **fluid with a cap**, never fixed:

```css
.pi-welcome-overlay { padding: var(--pi-dialog-gutter); }
.pi-overlay-card      { width: min(100%, var(--pi-dialog-w-m)); max-height: 100%; }
.pi-overlay-card--s   { width: min(100%, var(--pi-dialog-w-s)); }
.pi-overlay-card--l   { width: min(100%, var(--pi-dialog-w-l)); }
```

Plain px and `vh` only — no `dvh`/`svh`, `:has()`, or `env()`. The add-in still
runs in older WebView2 and WPS CEF builds.

### Cascade order gotcha

`src/ui/theme.css` imports in this order:
`tokens → preflight → base → components → overlays → content-overrides →
dialogs → queue → legacy-webview`.

Consequences:

- `overlays/*.css` beats `components/*.css` at equal specificity. Never rely on
  a media query in `components/` to override a plain rule in `overlays/` — that
  is exactly how the short-viewport override became dead code.
- `dialogs.css` loads last and therefore wins ties against everything above it.
- Put the *base* size on `.pi-welcome-overlay` / `.pi-overlay-card*` and treat
  per-dialog files as refinements, not competing definitions.

### One scroll region per dialog

Header and footer are `flex-shrink: 0`; exactly **one** descendant scrolls
(`.pi-overlay-body` for dialogs, `.pi-set-shell__body` for the settings shell).
Nested scrollers (`overflow-y: auto` inside a scroller) halve the usable height
and hide the sticky footer, so lists inside a dialog body use
`overflow: visible`.

## Toasts

Implementation: `src/ui/toast.ts` (DOM/timers) + `src/ui/toast-model.ts` (pure
stacking policy) + `src/ui/theme/components/toasts.css`.

- Markup: a fixed `.pi-toast-host` spanning `left: 8px; right: 8px` with
  `pointer-events: none`, containing `.pi-toast` flex items with
  `pointer-events: auto`. **Never** reintroduce a fixed/centred toast or a
  `white-space: nowrap` action button.
- Message: `flex: 1 1 12ch; min-width: 0` — the action wraps underneath below
  420px instead of eating the text column.
- Every toast has its own `×` (`.pi-toast__close`).
- At most **3** toasts. Eviction order: oldest plain info → oldest error →
  oldest sticky. A repeated message reuses its toast and resets the timer.
- Only the newest action toast keeps its button (the previous one is demoted to
  a plain toast), preserving the single-undo invariant behind the Cmd/Ctrl+Z
  arbitration in `taskpane/keyboard-shortcuts.ts`.
- Timers pause on `mouseenter` and `focusin`, resume on `mouseleave`/`focusout`.
- Durations: info 2s, error 6s, action 7s; `duration: 0` means "until
  dismissed". Extension toasts (`ctx.ui.toast`) go through the same module.

## Dismiss matrix

Every transient surface is dismissable. The guard test below enforces that a
notice module cannot ship without one.

| Surface | Module | Control | Lifetime |
|---|---|---|---|
| toast (info/error/action) | `src/ui/toast.ts` | `×` per toast (`createDismissControl`) | auto 2s / 6s / 7s, paused on hover/focus |
| runtime error banner | `src/ui/error-banner.ts` | `×` (`createDismissControl`) | until dismissed, the next error, or a successful turn |
| proxy helper banner | `src/ui/proxy-banner.ts` | `×` (`createDismissControl`) | pane session; re-arms when the proxy is detected, so a later outage notifies again |
| status-bar popovers | `src/taskpane/status-popovers.ts` | header `×` (`createDismissControl`) | until closed; Escape and outside-click unchanged |
| disclosure bar | `src/ui/disclosure-bar.ts` | `×` + "Got it" | persisted acknowledgement (`pi.onboarding.disclosure.acknowledged`) |
| bridge setup card | `src/ui/bridge-setup-card.ts` | "Dismiss" (`__dismiss`) | removed from its tool card |
| web-search setup card | `src/ui/web-search-setup-card.ts` | "Dismiss" (`__dismiss`) | removed from its tool card |
| overlay dialogs / settings shell | `src/ui/overlay-dialog.ts`, `src/ui/settings-shell.ts` | header `×` (`createDismissControl`) + Escape + backdrop | until closed |
| thinking-level flash bar | `src/taskpane/status-bar.ts` | none — decorative, `pointer-events: none` | 2 animation frames |
| fatal bootstrap error | `src/ui/loading.ts` `renderError` (via `bootstrap.ts`) | none — intentionally permanent | page reload |

"Per pane session" means in-memory only: reloading the taskpane clears the
dismissal. Notification dismissals are never persisted except the disclosure
acknowledgement.

## Tests that pin these invariants

- `tests/toast-model.test.ts` — stacking, eviction, sticky demotion, undo target.
- `tests/toast.test.ts` — host/close/pause/timer DOM behaviour.
- `tests/ui-dismiss-control.test.ts` — one dismiss control for every surface.
- `tests/error-banner.test.ts`, `tests/proxy-banner-dismiss.test.ts`,
  `tests/disclosure-bar-dismiss.test.ts`, `tests/status-popovers.test.ts` —
  per-surface dismissal.
- `tests/overlay-layout.test.ts` — CSS invariants (no centred toast, fluid card
  widths, single scroll region, no dead short-viewport override).
- `tests/notification-dismiss-guard.test.ts` — no notice module without a close
  control.
- `tests/extension-toast-delegation.test.ts` — one toast implementation.

## Verifying visually

`src/ui-gallery.ts` sections `toasts`, `dialogs`, `notices` and `layout-probe`
(see `AGENTS.md`). The probe prints measured card/toast boxes for the current
viewport as JSON on the console, which is how the numbers below were taken.

Measured after the fix (Chromium, 2026-11-10; the taskpane-shaped pane is
330 × 460 px, i.e. slightly narrower than the 360 px assumption above):

| check | pane | before | after |
|---|---|---|---|
| long message + action toast | 330 wide | 183 × 752px, message 12px | 330 × 82px, message 120px |
| same toast, very narrow | 220 wide | one glyph per line | 220 × 96px, message 160px, action wrapped below |
| toast on a wide pane | 754 wide | 377 × 68px | 480 × 54px, message 270px (capped by `--pi-toast-max`) |
| `--s` card | 330 × 460 | 331 × 368 (80% h) | 312 × 442 (95% w, 96% h) |
| `--m` card (resume) | 330 × 460 | 331 × 368 (80% h) | 312 × 442 (95% w, 96% h), one scroller, footer visible |
| `--l` card (settings) | 330 × 460 | 331 × 377 (82% h) | 312 × 442 (95% w, 96% h) |
| `--m` / `--l` on a wide pane | 754 × 480 | 520 / 720 wide, 80–82% h | 720 / 728 wide, 97% h |

Reproduce with `npm run build`, serve `dist/` and open
`src/ui-gallery.html?` at the target pane size, or run `./scripts/ui-verify.sh
layout-probe` against the dev server. The measurements above were taken with the
built CSS (`register-components-*.css`) at those exact viewport sizes.
