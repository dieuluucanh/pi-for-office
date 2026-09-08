# @dieulc/pi-office-protocol

Shared WebSocket wire protocol between the pi-for-office task-pane add-in
(`packages/add-in/src/bridge/`) and the Pi bridge extension
(`packages/bridge-extension/`).

Holding the protocol in one package keeps the two halves of the bridge in
lockstep — a breaking change is one commit, and the end-to-end interop test
(`packages/bridge-extension/tests/pane-interop.mjs`) proves both sides speak the
same format at runtime.

## Messages

- **Client → Server:** `hello`, `ping`, `tool_result`, `user_message`, `status`
- **Server → Client:** `welcome`, `pong`, `tool_call`, `agent_message` (delta/final), `tool_activity`, `error`

See `src/protocol.ts` for the typed shapes. Bump `BRIDGE_PROTOCOL_VERSION` on
breaking changes.
