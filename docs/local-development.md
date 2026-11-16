# Local Development

How to run and test pi-for-office locally — the add-in, the Pi bridge
extension, and the Python/tmux bridge servers. For **releasing** (npm publish +
deployment) see [releasing.md](./releasing.md).

## Component map

| Component | What it is | Runs where | Port (default) |
| ----------- | ----------- | ------------ | ---------------- |
| `packages/add-in` | Office.js task-pane add-in | your browser / Office sideload | HTTPS `3141` (Vite dev) |
| `packages/bridge-extension` | `@dieulc/pi-office-bridge` — WebSocket bridge, lets Pi drive Excel/Word/PowerPoint | inside a local Pi process | `38617` (loopback) |
| Python bridge | `pi-for-office-python-bridge` — run Python / LibreOffice locally | separate Node process | `3340` |
| Tmux bridge | `pi-for-office-tmux-bridge` — tmux terminal access | separate Node process | `3341` |
| CORS proxy | `pi-for-office-proxy` — OAuth/CORS helper | separate Node process | `3003` |

## Dev vs production modes

Every local service uses **one loopback port**, and dev and production builds
share it. The build mode no longer changes a port — it only changes the setup
commands the UI suggests (repo scripts for a Vite build, the published
one-liners for a production build). That means a Vite **dev add-in can attach
to the same Pi session and helpers you use for real work**; you no longer need
a second copy of every service.

| Service | Port | Started by |
| --- | --- | --- |
| Pi bridge WS | `38617` | the installed bridge extension, or `npm run bridge:dev` (repo source) |
| Python bridge | `3340` | `npx pi-for-office-python-bridge` or `npm run python:bridge:https` |
| Tmux bridge | `3341` | `npx pi-for-office-tmux-bridge` or `npm run tmux:bridge:https` |
| CORS proxy | `3003` | `npx pi-for-office-proxy` or `npm run proxy:https` |

**One process owns each port.** Start the repo script only after stopping (or
`/office close`-ing) the published service on that port; a second listener
fails with `EADDRINUSE` (the bridge's message names the port).

**Legacy URLs.** Older dev builds stored dev-port URLs. On boot they are
repaired once, in every build, by exact match:

| Stored value | Repaired to |
| --- | --- |
| `ws://127.0.0.1:38618` | `ws://127.0.0.1:38617` |
| `https://localhost:3004` (and `http://10.0.2.2:3004` for the WPS harness) | `https://localhost:3003` (and `http://10.0.2.2:3003`) |
| `https://localhost:3350` (also the Python approval grant) | `https://localhost:3340` |
| `https://localhost:3351` | `https://localhost:3341` |

Custom URLs (any other host/port) and the org `VITE_PI_DEFAULT_PROXY_URL`
override are never rewritten.

**Start the full dev stack** (see the per-service sections below):

```bash
# Terminal 1 — repo-source Pi bridge (same port as the installed extension)
npm run bridge:dev   # repo root; then /office connect in that Pi process

# Terminal 2 — Python bridge (repo source)
cd packages/add-in && npm run python:bridge:https

# Terminal 3 — tmux bridge (repo source)
cd packages/add-in && npm run tmux:bridge:https

# Terminal 4 — CORS proxy (only if testing OAuth/provider features)
cd packages/add-in && ALLOW_ALL_TARGET_HOSTS=1 npm run proxy:https

# Terminal 5 — add-in dev server
cd packages/add-in && npm run dev
```

Then sideload the add-in and enable the Local Pi agent toggle. The Bridge URL
row defaults to `ws://127.0.0.1:38617` for both the dev and production builds,
so the dev add-in can attach to your real Pi session (or to the repo-source
bridge from Terminal 1).

## Prerequisites

- Node.js ≥ 20
- `mkcert` (for the HTTPS dev server + bridge certs):
  - macOS: `brew install mkcert`
  - Windows: `choco install mkcert` or `scoop install mkcert`
  - Linux: `apt install mkcert` (or build from source)
- Pi CLI installed and logged in (`pi --version`, `pi api` etc.)
- Office desktop app (Excel / Word / PowerPoint) for sideload testing

One-time setup from the repo root:

```bash
npm install

# TLS certificates for the HTTPS dev server + bridge servers
mkcert -install
mkcert localhost
mv localhost-key.pem key.pem && mv localhost.pem cert.pem
```

## 1. Add-in dev server

```bash
cd packages/add-in
npm run dev
# → https://localhost:3141/src/taskpane.html
```

Sideload into Office (Excel/Word/PowerPoint):

- **Insert → My Add-ins → Upload My Add-in** → pick
  `packages/add-in/manifest.xml`.
- The task pane loads from `https://localhost:3141`.

Useful scripts (see `packages/add-in/package.json`):

| Script | Purpose |
| -------- | --------- |
| `npm run dev` | Vite dev server, HTTPS on port 3141 |
| `npm run dev:portless` | Dev through <https://pi-excel.localhost> (no mkcert/port) |
| `npm run serve:dist` | Serve a production `dist/` build over HTTPS |
| `npm run manifest:dev` | Regenerate the dev manifest |
| `npm run build` | Production build to `dist/` |
| `npm run typecheck` | TypeScript typecheck |
| `npm run test:context` / `test:security` / `test:manifest` | Targeted test suites |
| `npm test` | Full add-in test suite |

## 2. Pi bridge extension (Local Pi agent)

Two parts must talk: the **extension** runs inside a Pi process; the **add-in
card** enables the pane↔Pi connection.

### 2a. Run the extension

Two ways to run the bridge, both on the same port `38617`:

- **Installed extension (normal path):** `pi install npm:@dieulc/pi-office-bridge`
  and restart Pi. Every add-in build — dev or hosted — defaults to `38617`, so
  your real Pi session serves the dev pane too.
- **Repo source (developer path):** run a dedicated Pi process with the local
  extension. `--no-extensions` ignores settings, so the globally installed
  bridge never double-loads or races for the port:

```bash
npm run bridge:dev   # repo root → pi --no-extensions -e ./packages/bridge-extension/src/index.ts
```

The bridge stays **off until you run `/office connect`** in that Pi process
(manual lifecycle since bridge 0.6). If another session already owns `38617`,
`/office close` it there first — the bind error names the port. After connect
you should see:

```
Office bridge listening on ws://127.0.0.1:38617
```

Keep that Pi process running in the background.

### 2b. Enable it in the add-in

1. With the add-in open in Excel/Word/PowerPoint, go to
   **Settings → Connections → Local Pi agent (advanced)**.
2. Flip the **Enable local Pi agent** toggle **on**. The card's status flips to
   **Connecting…** then **Connected** — live, no taskpane reload needed.
3. Toggling **off** disconnects immediately.

The toggle persists `pi-bridge.enabled` in add-in settings, so the bridge
reconnects automatically on the next taskpane load.

The connection **self-heals**: if the Pi socket drops, the card shows
**Reconnecting (attempt N)…** with exponential backoff (1 s → 2 s → … →
30 s cap, reset after 30 s of stability) and reconnects automatically when Pi
returns — immediately on network regain / tab focus. A ping/pong watchdog
force-disconnects half-open sockets (e.g. after hibernation), so the badge
never stays **Connected** on a dead socket.

> Before this feature, `pi-bridge.enabled` was never written by the UI — the
> bridge could never actually start. The toggle (and the `/health` endpoint)
> are the fix; keep them in mind when touching bridge setup.

### 2c. Verify

```bash
# The bridge server answers HTTP /health on its port (bridge ≥ 0.2.0):
curl http://127.0.0.1:38617/health   # whichever bridge owns the port
# → { "ok": true, "service": "pi-office-bridge", "serverVersion": "0.x.y",
#     "capabilities": ["http-health", …], "catalogVersion": 1, "panes": [ … ] }
```

`panes` lists attached panes with their host (`excel` / `word` /
`powerpoint`), so you can confirm host detection right away. `serverVersion`
and `capabilities` tell clients the bridge is current.

The card's **Test connection** button runs this probe and classifies the
result: OK, “older bridge — /health unavailable”, timeout, or
“browser blocked local network access” — never a bare failure. If it reports
an older bridge, update it and restart Pi:

```bash
pi install npm:@dieulc/pi-office-bridge@latest
```

> **Local Network Access caveat:** on **Office on the web** with Chrome 142+
the browser may block cross-origin fetch to loopback. Desktop Office
(WebView2) is unaffected. The probe reports this as “browser blocked local
network access” — the live WebSocket status is still authoritative.

In the add-in, ask the agent to use an `office_*` tool (e.g. “read the current
sheet” in Excel, “summarize this document” in Word). The result should return
real document content — not “Unknown op”.

### 2e. Custom URL / port

The bridge card has a **Bridge URL** row (default `ws://127.0.0.1:38617` for
every build). To run the bridge on another port, start Pi with
`--office-bridge-port <port>` or `PI_OFFICE_BRIDGE_PORT=<port>` and set the
card URL to match — **both** the WebSocket client and the `/health` probe use
it. The URL must stay loopback (`127.0.0.1` / `localhost`) for security.

### 2d. Run the extension tests

```bash
cd packages/bridge-extension
npm test
# smoke test (WebSocket + /health) then interop test (real client ↔ real server)
```

## 3. Python bridge

Both the published one-liner (`npx pi-for-office-python-bridge`) and the repo
script bind the shared port `3340` (see *Dev vs production modes*). From
`packages/add-in`:

```bash
# Stub mode (simulated responses; nothing installed locally) — HTTP:
npm run python:bridge
# HTTPS:
npm run python:bridge:https

# Real mode (needs python3 + LibreOffice):
PYTHON_BRIDGE_MODE=real npm run python:bridge:https
```

Stop any other listener on `3340` first; the repo script does not choose a
random port.

Verify:

```bash
curl -k https://localhost:3340/health
curl -k -X POST https://localhost:3340/v1/python-run \
  -H "Content-Type: application/json" \
  -d '{"code": "result = {\"sum\": 1+1}", "timeout_ms": 5000}'
```

The add-in auto-probes `https://localhost:3340/health` at session start; if the
bridge is unreachable, a setup card offers `npx pi-for-office-python-bridge`.
The `python_run` tool falls back to in-browser Pyodide when the bridge is off.

## 4. Tmux bridge

Both the published one-liner (`npx pi-for-office-tmux-bridge`) and the repo
script bind the shared port `3341` (see *Dev vs production modes*). From
`packages/add-in`:

```bash
# Stub mode (all platforms) — HTTP:
npm run tmux:bridge
# HTTPS:
npm run tmux:bridge:https

# Real tmux mode (macOS/Linux only; tmux must be installed):
TMUX_BRIDGE_MODE=tmux npm run tmux:bridge:https
```

Stop any other listener on `3341` first; the repo script does not choose a
random port.

Verify:

```bash
curl -k https://localhost:3341/health
curl -k -X POST https://localhost:3341/v1/tmux \
  -H "Content-Type: application/json" \
  -d '{"action": "list_sessions"}'
```

No fallback — the tmux tool needs the bridge server.

## Full dev stack

Already covered above — see **Start the full dev stack** in the
[*Dev vs production modes*](#dev-vs-production-modes) section. Every helper
uses one port (`38617` / `3340` / `3341` / `3003`), shared by the repo scripts
and the published `npx` one-liners; run only one process per port.

## Troubleshooting

| Symptom | Likely cause | Fix |
| --------- | -------------- | ----- |
| Agent says it has no `office_excel_modify_structure` (can't add sheets) | Installed bridge < 0.3.0 — its old catalog has no structure/modify tools, and it ignores the pane's `hello.ops` | `pi install npm:@dieulc/pi-office-bridge@latest`, restart Pi; verify with `/office-tools` (29 tools incl. `office_excel_modify_structure`) |
| “Test connection” says the bridge is an older version (/health unavailable) | Installed bridge < 0.2.0 (no HTTP surface) | `pi install npm:@dieulc/pi-office-bridge@latest`, restart Pi, probe again |
| Probe times out (“No answer from /health within 3s”) | Bridge not running, or Pi restarting | Start the bridge (`npm run bridge:dev`, then `/office connect`) or keep Pi running; `curl http://127.0.0.1:38617/health` |
| Agent says a tool is missing (`office_*`) | The bridge owning `38617` is older than the pane (unknown ops are dropped) | Update the installed bridge (`pi install npm:@dieulc/pi-office-bridge@latest`) + restart, or run `npm run bridge:dev`; `/office` shows each pane's ops and `opsIgnoredCount` |
| Probe says “The browser blocked local network access” | Office on the web + Chrome 142+ Local Network Access | Use desktop Office (WebView2), or allow localhost access in the browser |
| Badge stuck on “Reconnecting (attempt N)…” | Pi process died / socket dropped | Start Pi again — the card reconnects automatically (backoff + watchdog) |
| Bridge on a different port than the card expects | Pi started with `--office-bridge-port` / `PI_OFFICE_BRIDGE_PORT` | Set the card's Bridge URL row to that port; the default is `38617` for every build |
| Toggle stays “Error” with connection refused | Pi process not running the extension, or port conflict | Check the port (`EADDRINUSE`); close the owning session (`/office close`) or pass `--office-bridge-port` and match the card URL |
| “Unknown op …” | Client registry empty (older add-in build) | Rebuild the add-in; this was fixed by wiring `ALL_BRIDGE_OPS` |
| Word/PowerPoint registered as Excel | Host detection disabled | Rebuild the add-in; host detection runs from the bridge manager |
| Python/tmux probe fails with TLS warning | Self-signed bridge certs | `mkcert -install` once, or accept the cert in the taskpane webview |
| Port 3141 busy | Another dev server | Pass `--port` to Vite, e.g. `npm run dev -- --port 3142` |

## Related

- [releasing.md](./releasing.md) — publish + deploy flow
- `packages/bridge-extension/README.md` — bridge extension docs
- `packages/add-in/docs/install.md` — end-user sideload instructions
