# @dieulc/pi-office-protocol

Shared WebSocket wire protocol between the pi-for-office task-pane add-in
(`packages/add-in/src/bridge/`) and the Pi bridge extension
(`packages/bridge-extension/`).

Holding the protocol in one package keeps the two halves of the bridge in
lockstep — a breaking change is one commit, and the end-to-end interop test
(`packages/bridge-extension/tests/pane-interop.mjs`) proves both sides speak the
same format at runtime.

## Messages

- **Client → Server:** `hello`, `ping`, `tool_result`, `user_message`, `status`, `pane_state`
- **Server → Client:** `welcome`, `pong`, `tool_call`, `agent_message` (delta/final), `tool_activity`, `error`

## Protocol v2 (breaking)

v2 adds **document identity and targeting** so the bridge can serve several open
documents of the same host app:

- `hello.document: DocumentIdentity` — **required**. `{ id, name, source }`
  where `id` is a stable local hash (`url_sha256:…` / `wps_path_sha256:…`) or
  `null` for unsaved documents, and `source` is one of `"document.url"`,
  `"wps.activeWorkbook.fullName"`, `"unknown"`. Raw URLs/paths are never sent.
- `pane_state` (client → server) — best-effort `{ focused?, visible?, activeAt? }`
  so `document: "active"` can resolve to the document the user is looking at.
- `tool_call.documentId?: string` — the document the call was routed to (echoed
  for logging and pane-side validation).
- `agent_message` final gains `promptId?: string` — correlates a reply with the
  pane prompt that triggered it.
- `OFFICE_TARGET_PARAM = "document"` — the tool-argument key that selects a
  target; the server strips it before forwarding the call to a pane.

A v1 client is rejected at the handshake with `protocol_mismatch`; update both
halves together (protocol 0.4.0, bridge 0.4.0).

## Version & capabilities

`BRIDGE_PROTOCOL_VERSION` is bumped only on **breaking** message-shape changes.
Additive fields keep both halves compatible:

- `welcome.serverVersion?: string` — the bridge extension's package version.
- `welcome.capabilities?: BridgeCapability[]` — currently `"http-health"`,
  meaning the server answers `GET /health` with JSON metadata.

Both are **optional**. An absent `capabilities` array (or absent
`serverVersion`) means the peer is a legacy 0.1.0 bridge: clients must degrade
gracefully (the live WebSocket status is still authoritative) instead of
assuming the server is unreachable.

See `src/protocol.ts` for the typed shapes.
