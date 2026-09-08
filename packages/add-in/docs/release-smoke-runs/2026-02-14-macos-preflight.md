# Smoke Run — macOS preflight

- Date: 2026-02-14
- Commit: `48336f158c64d967f51b4ad7123182092566c5f9`
- Environment: local repo checkout on macOS (CLI validation)
- Checklist source: `docs/release-smoke-test-checklist.md`

## Commands executed

1. `npm run check`
2. `npm run build`
3. `npm run test:models`
4. `npm run test:context`
5. `npm run test:security`

All commands passed.

## Checklist status snapshot

| ID | Status | Notes |
|---|---|---|
| PRE-1 | Pass | Full preflight suite passed on commit above. |
| C-1 | Blocked | Requires live Excel workbook/session interaction. |
| C-2 | Blocked | Requires taskpane tab + restore UX interaction in Excel host. |
| C-3 | Blocked | Requires in-host mutate + restore checkpoint flow. |
| C-4 | Blocked | Requires formatting application in live workbook. |
| C-5 | Blocked | Requires extension install + widget interaction in Excel taskpane. |
| P-1 | Blocked | Requires live workbook data and agent-driven edits. |
| P-2 | Blocked | Requires configured provider + optional PDF bridge in live host. |
| P-3 | Blocked | Requires live workbook updates + extension install execution. |
| P-4 | Blocked | Requires tmux bridge + external coding-agent runtime integration. |
| I-1 | Blocked | Manual sideload/install verification in desktop Excel required. |
| I-2 | Blocked | Windows-specific manual pass required. |
| I-3 | Blocked | Requires interactive provider login + prompt in taskpane. |
| I-4 | Blocked | Requires running proxy + in-app `/login` verification. |
| H-1 | Blocked | Requires interactive network/provider failure injection in host. |
| H-2 | Blocked | Requires large workbook loaded in Excel host. |
| H-3 | Pass | Security checks passed (`npm run test:security`). |
| H-4 | Blocked | Requires explicit corruption/quota simulation in host environment. |

## Next run targets

1. macOS in-host run for C-1..C-5 + I-1/I-3/I-4
2. Windows install/login pass for I-2/I-3/I-4
3. Error-path matrix execution (H-1/H-4) with captured screenshots/log snippets
