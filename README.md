# pi-for-office

Pi-powered AI agent for Microsoft Office — Excel, Word, PowerPoint, and Power BI.
Native-Pi, open-source, and **zero changes to Pi core**.

> **Status:** Monorepo scaffolded. `packages/add-in` is a fork of
> [pi-for-excel](https://github.com/tmustier/pi-for-excel) (MIT) being generalized
> from Excel-only to a unified multi-host Office add-in.

## Packages

| Package | Description |
|---------|-------------|
| [`packages/add-in`](packages/add-in/) | Unified Office.js task-pane add-in (Excel + Word + PowerPoint). Forked from pi-for-excel. |
| [`packages/bridge-extension`](packages/bridge-extension/) | Native Pi extension — WebSocket bridge + Office tool proxy for the add-in (`pi install npm:@dieulc/pi-office-bridge`). |
| `packages/powerbi-visual` | Power BI custom visual (Phase 3, not yet scaffolded). |

## Architecture (planned)

```
Office app (task pane add-in)  ──WebSocket──►  local Pi process (bridge extension)
       │                                              │
   Office.js tools                         agent loop + system tools (bash, git, files)
       │                                              │
       └────────────── LLM providers (BYOK, or Pi's own auth)
```

- **Standalone mode** — the add-in runs a browser agent directly (like pi-for-excel).
- **Bridge mode** — when a local Pi process is running with the bridge extension
  installed, the add-in auto-detects it and gains full system-tool capabilities.

## Repo layout

```
pi-for-office/
├── packages/
│   ├── add-in/               ← Office.js add-in (Vite + Lit + pi-agent-core)
│   └── bridge-extension/     ← Pi extension (server side of the bridge)
├── package.json              ← npm workspaces root
└── README.md
```

## Development

```bash
npm install                 # install all workspaces
npm run typecheck           # typecheck every package
npm run build               # build every package
```

See each package's README for app-specific instructions.

## License

MIT. `packages/add-in` retains upstream attribution to
[pi-for-excel](https://github.com/tmustier/pi-for-excel) (© tmustier, MIT).
