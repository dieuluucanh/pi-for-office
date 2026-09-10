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

From a source checkout (workspace-linked deps resolve automatically):

```bash
cd packages/bridge-extension
npm install        # first time
pi -e ./src/index.ts
```

You should see a Pi notification:

```
Office bridge listening on ws://127.0.0.1:38617
```

Keep that Pi process running in the background. (For an installed, published
extension instead: `pi install npm:@dieulc/pi-office-bridge`.)

### 2b. Enable it in the add-in

1. With the add-in open in Excel/Word/PowerPoint, go to
   **Settings → Connections → Local Pi agent (advanced)**.
2. Flip the **Enable local Pi agent** toggle **on**. The card's status flips to
   **Connecting…** then **Connected** — live, no taskpane reload needed.
3. Toggling **off** disconnects immediately.

The toggle persists `pi-bridge.enabled` in add-in settings, so the bridge
reconnects automatically on the next taskpane load.

> Before this feature, `pi-bridge.enabled` was never written by the UI — the
> bridge could never actually start. The toggle (and the `/health` endpoint)
> are the fix; keep them in mind when touching bridge setup.

### 2c. Verify

```bash
# The bridge server now answers HTTP /health on its port:
curl http://127.0.0.1:38617/health
# → { "ok": true, "service": "pi-office-bridge", "panes": [ … ] }
```

`panes` lists attached panes with their host (`excel` / `word` /
`powerpoint`), so you can confirm host detection right away.

In the add-in, ask the agent to use an `office_*` tool (e.g. “read the current
sheet” in Excel, “summarize this document” in Word). The result should return
real document content — not “Unknown op”.

### 2d. Run the extension tests

```bash
cd packages/bridge-extension
npm test
# smoke test (WebSocket + /health) then interop test (real client ↔ real server)
```

## 3. Python bridge

```bash
cd packages/add-in

# Stub mode (simulated responses; nothing installed locally) — HTTP:
npm run python:bridge
# HTTPS:
npm run python:bridge:https

# Real mode (needs python3 + LibreOffice):
PYTHON_BRIDGE_MODE=real npm run python:bridge:https
```

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

```bash
cd packages/add-in

# Stub mode (all platforms) — HTTP:
npm run tmux:bridge
# HTTPS:
npm run tmux:bridge:https

# Real tmux mode (macOS/Linux only; tmux must be installed):
TMUX_BRIDGE_MODE=tmux npm run tmux:bridge:https
```

Verify:

```bash
curl -k https://localhost:3341/health
curl -k -X POST https://localhost:3341/v1/tmux \
  -H "Content-Type: application/json" \
  -d '{"action": "list_sessions"}'
```

No fallback — the tmux tool needs the bridge server.

## Full dev stack

```bash
# Terminal 1 — Pi bridge extension
cd packages/bridge-extension && pi -e ./src/index.ts

# Terminal 2 — Python bridge
cd packages/add-in && npm run python:bridge:https

# Terminal 3 — tmux bridge
cd packages/add-in && npm run tmux:bridge:https

# Terminal 4 — CORS proxy (only if testing OAuth/provider features)
cd packages/add-in && npm run proxy:https

# Terminal 5 — add-in dev server
cd packages/add-in && npm run dev
```

Then sideload the add-in and enable the Local Pi agent toggle.

## Troubleshooting

| Symptom | Likely cause | Fix |
| --------- | -------------- | ----- |
| “Could not reach the bridge.” in the Test connection probe | Bridge extension not running, or wrong port | Start `pi -e ./src/index.ts`; check `curl http://127.0.0.1:38617/health` |
| Toggle stays “Error” with connection refused | Pi process not running the extension, or port conflict | Check the port (`EADDRINUSE`); override with `PI_OFFICE_BRIDGE_PORT` on both sides |
| “Unknown op …” | Client registry empty (older add-in build) | Rebuild the add-in; this was fixed by wiring `ALL_BRIDGE_OPS` |
| Word/PowerPoint registered as Excel | Host detection disabled | Rebuild the add-in; host detection runs from the bridge manager |
| Python/tmux probe fails with TLS warning | Self-signed bridge certs | `mkcert -install` once, or accept the cert in the taskpane webview |
| Port 3141 busy | Another dev server | Pass `--port` to Vite, e.g. `npm run dev -- --port 3142` |

## Related

- [releasing.md](./releasing.md) — publish + deploy flow
- `packages/bridge-extension/README.md` — bridge extension docs
- `packages/add-in/docs/install.md` — end-user sideload instructions
