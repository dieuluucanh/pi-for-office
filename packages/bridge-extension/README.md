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

**The server is off until you ask for it (bridge 0.6+).** Opening a Pi session no
longer claims the bridge port. Run `/office connect` in the session you want to
use, or click **Connect this session** in Pi Desktop's Documents rail. Every new
session starts off; `/office close` releases the port for another session.

Two flows:

1. **Tool proxy (Pi → pane):** the Pi agent calls `office_excel_read_range`,
   `office_word_insert_text`, …; the extension forwards a `tool_call` to the
   attached pane; the pane runs the Office.js op and answers with a
   `tool_result`. The LLM then sees the document content **and** has Pi's full
   system tools (bash, git, files).

2. **Pane-driven chat (pane → Pi):** the user types in the add-in sidebar; the
   pane forwards a `user_message`; the extension injects it into the Pi session
   and streams the assistant's final reply back to the pane.

## Multiple open documents (protocol v2)

The bridge serves **every open document of every attached host app**, not just
the most recently connected one. In `hello` each pane reports a
privacy-preserving document identity (`document.id` = a local hash,
`document.name` = display name, `document.source`), and the server routes each
tool call to one specific document:

- Every `office_*` tool takes an optional **`document`** argument: a document id
  (from `office_list_documents`), a unique document name, or `"active"`.
- An omitted / `"active"` target resolves to the **focused** document, then to
  the only attached document of that host. If several documents are open and
  none is focused, the call **fails with the candidate list** instead of
  editing a guess.
- Panes report focus/visibility with `pane_state` frames; typing in a pane also
  counts as focus, and Excel additionally uses `Workbook.onActivated`.
- `office_list_documents` lists the attached documents (id, name, host, focused,
  windows); the system prompt carries the same list before every turn.
- Two windows of the **same file** share one document id: the call routes to the
  focused window, and errors if neither is focused.
- In the Pi TUI, type **`#`** while composing a prompt to pick a document from a
  dropdown (focused first). Enter inserts `#<name>` — or `#<document id>` when
  two open documents share a name. The server strips the leading `#` when
  resolving the target, so the mention is directly usable.
- Agent replies are matched to the pane prompt that caused them and are sent
  **only to that pane** (never broadcast), and the injected prompt carries an
  `[office:<host> "<name>" id=…]` tag so a shared conversation stays traceable.

**Protocol v2 is breaking:** update the add-in and the bridge together
(protocol 0.4.0 + bridge 0.4.0). A v1 pane is rejected at the handshake with
`protocol_mismatch`.

## Install

```bash
pi install npm:@dieulc/pi-office-bridge
```

**Update to the latest** (the full 29-tool catalog incl. structure/format tools
needs ≥ 0.3.0; multi-document targeting, `office_list_documents` and `#` mention
autocomplete need ≥ 0.4.0 (protocol v2, which requires the matching add-in
build); the `/health` probe and version display need ≥ 0.2.0):

```bash
pi install npm:@dieulc/pi-office-bridge@latest
```

Then restart Pi. Note that `pi install` pins the version it fetched into
`~/.pi/agent/npm/package.json`, so Pi will **not** auto-upgrade — re-run the
command above to get the newest bridge.

Or from the monorepo (development) as the **dev bridge** on port `38618`
(`--no-extensions` ignores settings, so the globally installed prod bridge
never double-loads in the same process):

```bash
npm run bridge:dev   # repo root → pi --no-extensions -e ./packages/bridge-extension/src/index.ts --office-bridge-port 38618
```

> **Troubleshooting: the agent says it has no `office_excel_modify_structure`**
> (or any other newer tool). A bridge **< 0.3.0** ships a small hardcoded
> catalog (12 tools) and ignores the pane's `hello.ops`, so newer ops silently
> never register. Fix: `pi install npm:@dieulc/pi-office-bridge@latest`, restart
> Pi, then verify with `/office-tools` (29 tools) and `/office` (pane ops +
> catalog version).

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

- `serverVersion` — this package's version (e.g. `"0.6.0"`).
- `capabilities` — `http-health`, `agent-deltas`, `client-abort` and
  `manual-lifecycle`. `manual-lifecycle` means the bridge stays off until
  `/office connect`; a client that needs to start or stop it must check this
  capability (or the version), never assume the 0.6 commands exist.
- `GET /health` also carries `owner` — `{ pid, sessionId, client, cwd }` — so
  Pi Desktop can map the live bridge to a tab or report a terminal owner. It is
  null/absent on older bridges.

Clients (the add-in card) treat an absent `capabilities` as “legacy bridge
(< 0.2.0)” and tell the user to update instead of claiming the bridge is
down. The connection state is surfaced live to the Pi TUI as soon as a pane
attaches or detaches (`onPanesChanged`).

## Commands

| Command | Description |
|---------|-------------|
| `/office` | Status (same as `/office status`) |
| `/office status` | Show state, port, owner and attached apps |
| `/office connect [port]` | Start the bridge; the port defaults to the configured port |
| `/office close` | Stop the bridge and disconnect the add-in |
| `/office restart [port]` | Stop, then start (optionally on another port) |
| `/office help` | Show the command list |
| `/office-tools` | List every `office_*` tool and whether it is active |

Aliases: `on`/`start` = `connect`; `stop`/`off`/`disconnect` = `close`;
`info` = `status`.

## Lifecycle (bridge 0.6+)

```text
/office connect → listening → /office close → off
```

- Only one Pi session can own the port. A second `/office connect` fails with a
  message naming the port and telling you to close the owner (or connect on a
  different port).
- The status line shows the bound port and the attached apps while on; it is
  cleared when the bridge is off.
- Office tools are registered at `session_start` but **inactive while off**. The
  discovery tool `office_list_documents` becomes active as soon as the bridge is
  on; catalogue tools activate only for the ops the attached pane advertises.
- `session_shutdown` always stops the bridge. `/new`, `/resume`, `/reload` and a
  Desktop host restart therefore start off; reconnect explicitly. Keeping a
  server across Pi's extension re-instantiation would orphan it.
- The bridge never auto-connects just because a session opened. The opt-in
  `--office-bridge-autostart` / `PI_OFFICE_BRIDGE_AUTOSTART=1` restores the old
  behaviour for scripts and power users.

## Configuration

- **Port** — flag `--office-bridge-port <port>` or env `PI_OFFICE_BRIDGE_PORT`
  (default `38617`). `/office connect <port>` overrides both for one run. The
  add-in connects to the same default; change both if you override it (the
  add-in's bridge card has a **Bridge URL** row that both the WebSocket client
  and the probe use). If the port is already taken by another Pi process, the
  extension reports `EADDRINUSE` with the owner/close hint.
- **Autostart** — off by default. `--office-bridge-autostart` or
  `PI_OFFICE_BRIDGE_AUTOSTART=1` starts the bridge in `session_start` (the old
  pre-0.6 behaviour). There is no persisted GUI setting.
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

Plus the host-agnostic discovery tool **`office_list_documents`** (id, name,
host, focused, windows). It is active only while the bridge is on.

### Active-tool reconciliation

The office tools are registered at `session_start` from the catalog, but the
active set follows the lifecycle and the attached panes
(`pi.setActiveTools()`):

| State | Catalogue tools | `office_list_documents` |
| --- | --- | --- |
| bridge off | inactive | inactive |
| bridge on, no pane | inactive | active |
| bridge on, pane attached | exactly the pane's advertised ops | active |

Opening an app activates that host's tools; closing it deactivates them;
closing the bridge deactivates the whole family. Everything else stays
untouched. Panes that don't advertise an `ops` list (legacy 0.2.x clients) are
given only the v1 op set.

### Capability handshake

Panes send `hello.ops` + `hello.catalogVersion` with the op ids they can
execute. The server validates them against its own catalog (entries that don't
belong to the pane's host, or that the server doesn't know, are dropped and
counted). `callOfficeTool` then rejects any op the pane did not advertise with
an actionable message, so a mismatched add-in/bridge pair fails loudly instead
of silently. That op list — together with `opsIgnoredCount` — is the real
enforcement: `catalogVersion` is informational only (reported in `/health` and
the `/office` summary), so bumping it never silently changes what a pane may
execute.

- `/office` shows each attached pane's host, op count, catalog version, and any
  ignored-op count.
- `/office-tools` lists every registered tool + catalog version.
- `GET /health` exposes `catalogVersion` and per-pane `ops` / `catalogVersion`.
- `before_agent_start` appends a pane-context block (attached host, "the
  office_* tools edit the live document", "never emit HTML for Word") so the
  agent uses the tools directly. It deliberately does **not** claim formatting is
  universally supported: each tool's schema lists exactly the properties it
  accepts, and the block instructs the agent to say a property is not available
  rather than assert support — consulting `office_word_get_capabilities` when a
  request depends on a Word feature it is unsure about.

Tools never silently degrade: when a requested feature needs an Office.js
requirement set the host lacks, the call returns a message naming the capability,
the set + version, the platform, and the remedy. Two CI guards keep the surface
honest — every advertised formatting property must have a real assignment in the
implementation that applies it, and every declared capability must either be
consumed by a caller or be explicitly probe-only. The panes and this extension
both read the *same* catalog parameter objects (no per-side schema copies), so a
new parameter cannot drift between the two agents.

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
(`packages/protocol`). It is currently at **v3** (streamed deltas + pane abort);
bump `BRIDGE_PROTOCOL_VERSION` on breaking changes. The lifecycle commands and
the `/health` owner are additive and do not change the wire version. See
[`packages/protocol/README.md`](../protocol/README.md) for the changelog.

## License

MIT
