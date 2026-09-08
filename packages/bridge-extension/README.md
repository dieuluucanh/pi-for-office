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

## Commands

| Command | Description |
|---------|-------------|
| `/office` | Show bridge status: port + attached apps (Excel/Word/PowerPoint) |
| `/office-tools` | List every `office_*` tool registered |

## Configuration

- **Port** — flag `--office-bridge-port <port>` or env `PI_OFFICE_BRIDGE_PORT`
  (default `38617`). The add-in connects to the same default; change both if you
  override it.

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
