# @dieulc/pi-office-bridge

Native Pi extension that lets a local Pi process drive **Excel**, **Word**, and
**PowerPoint** through the [pi-for-office](../add-in/README.md) task-pane
add-in.

Pure extension — it only uses Pi's public extension API, so **Pi core is never
touched** and Pi can be updated freely.

## How it works

```
Excel / Word / PowerPoint  (pi-for-office task pane)
        │  WebSocket  ws://127.0.0.1:38617
        ▼
local Pi process  (this extension)
   • registers office_<host>_<op> tools
   • proxies Office.js calls back to the pane
   • injects pane prompts into the Pi session
```

Two flows:

1. **Tool proxy (Pi → pane):** the Pi agent calls `office_excel_read_range`,
   `office_word_insert_text`, …; the extension forwards a `tool_call` to the
   attached pane; the pane runs the Office.js op and answers with a
   `tool_result`. The LLM then sees the document content **and** has Pi's full
   system tools (bash, git, files).

2. **Pane-driven chat (pane → Pi):** the user types in the add-in sidebar; the
   pane forwards a `user_message`; the extension injects it into the Pi session
   and streams the assistant's final reply back to the pane.

## Install

```bash
pi install npm:@dieulc/pi-office-bridge
```

**Update to the latest** (the add-in's `/health` probe and version display
require ≥ 0.2.0):

```bash
pi install npm:@dieulc/pi-office-bridge@latest
```

Then restart Pi. Note that `pi install` pins the version it fetched into
`~/.pi/agent/npm/package.json`, so Pi will **not** auto-upgrade — re-run the
command above to get the newest bridge.

Or from the monorepo (development):

```bash
cd packages/bridge-extension
npm install
# then load it in pi for a quick test:
pi -e ./src/index.ts
```

## Enable in the add-in

1. Install/run the bridge so a Pi process with this extension is listening
   (see above). Keep that Pi process running in the background.
2. Open pi-for-office in Excel / Word / PowerPoint.
3. Go to **Settings → Connections → Local Pi agent (advanced)** and flip the
   **Enable local Pi agent** toggle on. The card shows the live connection
   state (Connecting… → Connected); no taskpane reload is needed.
4. Verify with the card's **Test connection** button, or from a terminal:

```bash
curl http://127.0.0.1:38617/health
# → { "ok": true, "service": "pi-office-bridge", "serverVersion": "0.2.0",
#     "capabilities": ["http-health"], "panes": [ … ] }
```

`/health` lists the attached pane(s) and their host app (excel / word /
powerpoint), so it doubles as a quick host-detection check. The add-in's
probe classifies the response (current / older bridge / timeout / browser
blocked) instead of reporting a bare failure — see
[`docs/local-development.md`](../../docs/local-development.md).

## Version & capabilities

Every `welcome` frame (and `GET /health`) advertises additive server
metadata:

- `serverVersion` — this package's version (e.g. `"0.2.0"`).
- `capabilities` — `["http-health"]` means the HTTP `/health` surface is
  served.

Clients (the add-in card) treat an absent `capabilities` as “legacy bridge
(< 0.2.0)” and tell the user to update instead of claiming the bridge is
down. The connection state is surfaced live to the Pi TUI as soon as a pane
attaches or detaches (`onPanesChanged`).

## Commands

| Command | Description |
|---------|-------------|
| `/office` | Show bridge status: port + attached apps (Excel/Word/PowerPoint) |
| `/office-tools` | List every `office_*` tool registered |

## Configuration

- **Port** — flag `--office-bridge-port <port>` or env `PI_OFFICE_BRIDGE_PORT`
  (default `38617`). The add-in connects to the same default; change both if you
  override it (the add-in's bridge card has a **Bridge URL** row that both the
  WebSocket client and the probe use). If the port is already taken by another
  Pi process, the extension reports `EADDRINUSE` with the override hint.
- **Allowed origins** — env `PI_OFFICE_BRIDGE_ALLOWED_ORIGINS` (comma-separated)
  extends the browser origins allowed to read `GET /health`. Defaults cover the
  dev Vite server (`https://localhost:3141`) and the hosted GitHub Pages add-in
  (`https://dieuluucanh.github.io`). The pane's WebSocket connection is
  loopback-only and is not restricted by this list.

## Office tools

The extension registers a `office_<host>_<op>` tool per op in the shared
catalog. The catalog is the single source of truth: it lives in
`@dieulc/pi-office-protocol` (`office-catalog.ts`) and BOTH the Pi extension
and the add-in derive from it — the Pi side here, and the pane's bridge op
registry (`packages/add-in/src/bridge/`) there. Op ids are namespaced by host:

| Host | Ops |
| --- | --- |
| Excel | `get_overview`, `read_range`, `write_cells`, `fill_formula`, `search_workbook`, `modify_structure`, `format_cells`, `conditional_format`, `charts`, `trace_dependencies`, `explain_formula`, `view_settings`, `comments`, `workbook_history` |
| Word | `get_overview`, `read_document`, `insert_text`, `replace_text`, `format_range`, `insert_blocks`, `insert_table`, `insert_page_break`, `insert_image`, `insert_hyperlink` |
| PowerPoint | `get_overview`, `read_slide`, `add_slide`, `add_text_box`, `format_slide` |

### Active-tool reconciliation

The office tools are registered at `session_start` from the catalog, but only
**the ops the currently attached pane advertises are kept active** in the Pi
session (`pi.setActiveTools()`). Opening an app activates that host's tools;
closing it deactivates them; everything else stays untouched. Panes that don't
advertise an `ops` list (legacy 0.2.x clients) are given only the v1 op set.
This keeps the agent's prompt small and focused on the app actually open.

### Capability handshake

Panes send `hello.ops` + `hello.catalogVersion` with the op ids they can
execute. The server validates them against its own catalog (entries that don't
belong to the pane's host, or that the server doesn't know, are dropped and
counted). `callOfficeTool` then rejects any op the pane did not advertise with
an actionable message, so a mismatched add-in/bridge pair fails loudly instead
of silently.

- `/office` shows each attached pane's host, op count, catalog version, and any
  ignored-op count.
- `/office-tools` lists every registered tool + catalog version.
- `GET /health` exposes `catalogVersion` and per-pane `ops` / `catalogVersion`.
- `before_agent_start` appends a pane-context block (attached host, "the
  office_* tools edit the live document; formatting is fully supported; never
  emit HTML for Word") so the agent uses the tools directly.

The pane-side executors are thin delegates to the same local tool factories the
browser-only path uses (see `packages/add-in/src/bridge/`), and a parity test
(`packages/add-in/tests/bridge-catalog-parity.test.ts`) fails CI if the pane
registry ever drifts from the shared catalog.

## Development

```bash
npm run typecheck     # typecheck against @earendil-works/pi-coding-agent 0.85.x
npm run build         # emit dist/ (for the node smoke tests)
npm test              # smoke test + end-to-end interop test (real client ↔ real server)
```

`tests/pane-interop.mjs` wires the **real** add-in `PaneBridgeClient` to the
**real** bridge server through the shared `@dieulc/pi-office-protocol` package —
the strongest proof the two halves agree on the wire format.

## Protocol

The wire protocol is shared in `@dieulc/pi-office-protocol`
(`packages/protocol`). Bump `BRIDGE_PROTOCOL_VERSION` on breaking changes.

## License

MIT
