# Rollout plan (install → onboarding → hackability)

This doc captures the recommended order of work to make **Pi for Excel** easy to install for non-technical users *and* remain hackable/OSS-friendly.

## Principles

- **Non-technical users should not need**: git, Node, mkcert, a local dev server, or terminal commands.
- The add-in UI is a **static website** loaded by Excel from a manifest URL.
- Keep the project **OSS + self-hostable**: advanced users can run from source and/or host the static build themselves.
- Prefer an update story where **web assets update automatically** (hosted), and manifest updates are rare.

## Recommended order

### Step 1 — Distribution: hosted build + production manifest (Issue #16)

Goal: a user can install Pi for Excel by downloading a manifest and using Excel UI (no local server).

Deliverables:
- Host `dist/` at a stable HTTPS URL (Vercel recommended).
- Provide `manifest.prod.xml` that points to that host (`SourceLocation` + icons).
- Add a non-technical install doc (macOS + Windows) with screenshots.
- Update README to split:
  - **Install (recommended)**
  - **Developer setup**

**Update story (best-effort automatic):**
- If the manifest points to a stable URL (e.g. `https://<host>/src/taskpane.html`), then deploying a new build to the same host updates the add-in **without users reinstalling**.
- Users may need to close/reopen the taskpane or restart Excel if Office/webview caches aggressively.
- If we must change the manifest (rare), users reinstall the updated manifest.

### Step 2 — First-run onboarding success (Issue #11 + small onboarding improvements)

Goal: after install, users can connect a provider and get a first response quickly.

Deliverables:
- Better welcome copy + examples.
- Clear provider recommendations:
  - API key flows that “just work”
  - OAuth/subscription flows that may require the local HTTPS proxy helper
- Optionally validate keys with a tiny test request and show a friendly error.

### Step 3 — Hackability in hosted build: Extensions Manager (Issue #13)

Goal: users who are not engineers (but use Claude/LLM workflows) can install extensions *inside* the hosted add-in.

Deliverables:
- `/extensions` manager UI:
  - list installed extensions
  - enable/disable/uninstall
  - show last load error
  - reload
- Install methods:
  - **Paste code** (recommended): store in IndexedDB; load via Blob URL + `import()`
  - Optional: install from URL
- Copyable “ask Claude to write an extension” prompt template.

### Step 4 — Power-user lane: local models (Ollama / vLLM / LM Studio)

Goal: advanced users can connect local model servers easily.

Deliverables:
- Expose custom providers UI (Providers & Models).
- Ensure our provider filtering + API key resolution works for custom providers.
- Make Office webview + localhost HTTP workable (likely via the local HTTPS proxy).

### Step 5 — Packaged helper app (optional)

Goal: make the local HTTPS proxy (and future local features) non-technical:
- installer for macOS/Windows
- runs in background
- provides `https://localhost:<port>`

This improves OAuth/subscription flows + local model connectivity, but is higher effort.

## Tracking issues

- #16 Distribution: non-technical install (hosted build + prod manifest)
- #11 UX: revise welcome copy and example prompts
- #13 Extensions API: design & build-out
