# @dieulc/pi-office-protocol

Shared WebSocket wire protocol between the pi-for-office task-pane add-in
(`packages/add-in/src/bridge/`) and the Pi bridge extension
(`packages/bridge-extension/`).

Holding the protocol in one package keeps the two halves of the bridge in
lockstep — a breaking change is one commit, and the end-to-end interop test
(`packages/bridge-extension/tests/pane-interop.mjs`) proves both sides speak the
same format at runtime.

## Messages

- **Client → Server:** `hello`, `ping`, `tool_result`, `user_message`, `status`, `pane_state`, `abort`
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

## Protocol v3 (additive)

v3 makes a pane prompt feel live and interruptible. It is **additive**: a v2
pane is still accepted and simply receives final replies only.

- `agent_message` **delta** frames are now actually produced: while the agent
  answers a pane prompt the server streams
  `{ type: "agent_message", kind: "delta", text, promptId?, messageId? }`,
  coalesced to stay under `MAX_TEXT_CHARS`. Deltas are **display-only**: the
  concatenation is not guaranteed to equal the final text, so the pane replaces
  its live buffer with `kind: "final"` when it arrives.
- `abort` (client → server) — `{ type: "abort", promptId? }` stops the run
  answering that prompt (or whatever run is in flight for this pane when
  `promptId` is omitted). A v2 server ignores the frame, so a Stop button must
  stay usable if the reply keeps arriving.
- `MIN_SUPPORTED_PROTOCOL_VERSION = 2` and `isProtocolVersionSupported()` — the
  server accepts the whole `[MIN_SUPPORTED_PROTOCOL_VERSION, BRIDGE_PROTOCOL_VERSION]`
  window instead of demanding an exact match.
- Deltas are only sent to panes whose `hello.protocolVersion >= 3`; the server
  never sends a delta to a v2 pane (which would drop it).

Update both halves together on a breaking change; an additive change may ship
protocol-first (protocol 0.5.1, bridge 0.6.0).

## Version & capabilities

`BRIDGE_PROTOCOL_VERSION` is bumped only on **breaking** message-shape changes.
Additive fields keep both halves compatible:

- `welcome.serverVersion?: string` — the bridge extension's package version.
- `welcome.capabilities?: BridgeCapability[]` — currently `"http-health"`
  (the server answers `GET /health` with JSON metadata), `"agent-deltas"`
  (the server streams `agent_message` deltas), `"client-abort"` (the server
  honours the `abort` frame), and `"manual-lifecycle"` (the bridge stays off
  until `/office connect`; a client must not assume the 0.6 commands exist on a
  bridge without this capability). Prefer a capability check over a version
  comparison so a peer's features are discoverable at runtime.

Both are **optional**. An absent `capabilities` array (or absent
`serverVersion`) means the peer is a legacy 0.1.0 bridge: clients must degrade
gracefully (the live WebSocket status is still authoritative) instead of
assuming the server is unreachable.

See `src/protocol.ts` for the typed shapes.
