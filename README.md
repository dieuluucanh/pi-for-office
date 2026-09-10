# pi-for-office

Pi-powered AI agent for Microsoft Office — Excel, Word, PowerPoint, and Power BI.
Native-Pi, open-source, and **zero changes to Pi core**.

> **Status:** Working build. `packages/add-in` is a fork of
> [pi-for-excel](https://github.com/tmustier/pi-for-excel) (MIT) being generalized
> from Excel-only to a unified multi-host Office add-in.

## Packages

| Package | Description |
| --------- | ------------- |
| [`packages/add-in`](packages/add-in/) | Unified Office.js task-pane add-in (Excel + Word + PowerPoint). Forked from [pi-for-excel](https://github.com/tmustier/pi-for-excel). |
| [`packages/bridge-extension`](packages/bridge-extension/) | Native Pi extension — WebSocket bridge + Office tool proxy for the add-in (`pi install npm:@dieulc/pi-office-bridge`). |
| `packages/powerbi-visual` | Power BI custom visual (Phase 3, not yet scaffolded). |

## Architecture

```
Office app (task pane add-in)
        │
        ├── Browser-only (default) ──► LLM providers directly (BYOK / API keys)
        │        │                        no local process required
        │        └────────────► optional local CORS proxy (npx pi-for-office-proxy)
        │                             only needed for OAuth logins (Anthropic /
        │                             OpenAI ChatGPT / Google) & advanced features
        │
        └── Bridge mode (advanced) ── WebSocket ──► local Pi process (bridge extension)
                                                       agent loop + system tools
                                                       (bash, git, files)
```

Two operational modes, one app:

- **Browser-only mode (default)** — the add-in runs a browser agent directly and
  talks to LLM providers with Bring-Your-Own-Key API keys. No local proxy, no Pi
  process, no Node.js — this is the mode most users stay in.
- **Proxy mode (opt-in)** — enable in `/settings → Proxy` and run
  `npx pi-for-office-proxy` only when you need OAuth-based provider logins
  (Anthropic subscription, OpenAI ChatGPT, Google Code Assist/Antigravity,
  GitHub Copilot) that are CORS-blocked inside Office webviews.
- **Bridge mode (advanced)** — when a local Pi process is running with the
  bridge extension installed, the add-in auto-detects it and gains full
  system-tool capabilities (bash, git, files).

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

## Docs for development & releases

- [docs/local-development.md](docs/local-development.md) — run & test everything
  locally (add-in, Pi bridge extension, Python/tmux bridges, proxy), including
  how to enable the **Local Pi agent** in Settings → Connections.
- [docs/releasing.md](docs/releasing.md) — one tracked flow for publishing the
  npm packages (`npm run release:status` / `release:check` / `release:publish`)
  and deploying the add-in to GitHub Pages.

## License

MIT. `packages/add-in` retains upstream attribution to
[pi-for-excel](https://github.com/tmustier/pi-for-excel) (© tmustier, MIT).
