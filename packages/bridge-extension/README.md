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
# → { "ok": true, "service": "pi-office-bridge", "panes": [ … ] }
```

`/health` lists the attached pane(s) and their host app (excel / word /
powerpoint), so it doubles as a quick host-detection check.

## Commands

| Command | Description |
|---------|-------------|
| `/office` | Show bridge status: port + attached apps (Excel/Word/PowerPoint) |
| `/office-tools` | List every `office_*` tool registered |

## Configuration

- **Port** — flag `--office-bridge-port <port>` or env `PI_OFFICE_BRIDGE_PORT`
  (default `38617`). The add-in connects to the same default; change both if you
  override it.
- **Allowed origins** — env `PI_OFFICE_BRIDGE_ALLOWED_ORIGINS` (comma-separated)
  extends the browser origins allowed to read `GET /health`. Defaults cover the
  dev Vite server (`https://localhost:3141`) and the hosted GitHub Pages add-in
  (`https://dieuluucanh.github.io`). The pane's WebSocket connection is
  loopback-only and is not restricted by this list.

## Office tools

The extension registers a `office_<host>_<op>` tool per op in the shared
catalog. The catalog lives in
[`src/office-tools.ts`](./src/office-tools.ts); op ids are namespaced by host:

| Host | Ops |
|------|-----|
| Excel | `get_overview`, `read_range`, `write_cells`, `fill_formula`, `search_workbook` |
| Word | `get_overview`, `read_document`, `insert_text`, `replace_text` |
| PowerPoint | `get_overview`, `read_slide`, `add_slide`, `add_text_box` |

The pane-side executors are the counterpart contract — see
`packages/add-in/src/bridge/` (same repo). **When adding an op, update both
sides** (see "Bridge contract" in the add-in README).

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
