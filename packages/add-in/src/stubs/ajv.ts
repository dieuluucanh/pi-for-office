/**
 * Stub for Ajv in the Office Add-in environment.
 *
 * Ajv v8 uses `new Function()` internally when compiling JSON schemas.
 * The Office webview enforces a strict CSP (`script-src 'self' ...`) that
 * blocks `unsafe-eval`, so Ajv.compile() always throws at runtime.
 *
 * pi-ai's validation.js already has a fallback path: when `ajv` is null it
 * skips schema validation and trusts the LLM output directly. But the guard
 * only checks for Chrome extensions, not Office Add-ins.
 *
 * This stub ensures `new Ajv()` construction fails gracefully so the
 * existing fallback in pi-ai kicks in. We throw during construction so the
 * try/catch in validation.js sets `ajv = null`.
 */

class AjvStub {
  constructor() {
    throw new Error("Ajv disabled: Office Add-in CSP does not allow unsafe-eval");
  }
}

export default AjvStub;
