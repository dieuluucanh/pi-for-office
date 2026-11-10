/**
 * Persistent per-tool approval allowlist.
 *
 * Tools named in the allowlist skip their runtime approval prompt entirely:
 * both the direct-JS gate (`execute_office_js` / `execute_wps_js`) and the
 * Safe-mode workbook-mutation gate consult this list. Storage value is a
 * JSON array of tool names, mirroring the durable approval precedent set by
 * `python.bridge.approved.url`.
 *
 * The allowlist is intentionally per-tool and blanket: allowing a tool
 * suppresses prompts for that tool in every execution mode and for both pure
 * host-API code and ambient-authority-flagged code. Grants are surfaced and
 * revocable in Settings → Behavior → Always-allowed tools.
 */

export const ALWAYS_ALLOW_TOOLS_SETTING_KEY = "approvals.alwaysAllowTools.v1";

/** Upper bound on persisted entries to keep the setting small. */
export const MAX_ALWAYS_ALLOWED_TOOLS = 100;

export interface AlwaysAllowStore {
  get: (key: string) => Promise<DynamicValue>;
  set: (key: string, value: DynamicValue) => Promise<void>;
}

function normalizeToolName(value: DynamicValue): string | null {
  if (typeof value !== "string") {
    return null;
  }

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

/**
 * Normalize an unknown stored value to a deduplicated list of tool names.
 * Invalid entries are dropped; order is preserved.
 */
export function normalizeAlwaysAllowedTools(value: DynamicValue): string[] {
  if (!Array.isArray(value)) {
    return [];
  }

  const seen = new Set<string>();
  const tools: string[] = [];

  for (const entry of value) {
    const tool = normalizeToolName(entry);
    if (tool === null || seen.has(tool)) {
      continue;
    }

    seen.add(tool);
    tools.push(tool);

    if (tools.length >= MAX_ALWAYS_ALLOWED_TOOLS) {
      break;
    }
  }

  return tools;
}

export async function getStoredAlwaysAllowedTools(
  store: AlwaysAllowStore,
): Promise<string[]> {
  return normalizeAlwaysAllowedTools(
    await store.get(ALWAYS_ALLOW_TOOLS_SETTING_KEY),
  );
}

export async function setStoredAlwaysAllowedTools(
  store: AlwaysAllowStore,
  tools: string[],
): Promise<string[]> {
  const normalized = normalizeAlwaysAllowedTools(tools);
  await store.set(ALWAYS_ALLOW_TOOLS_SETTING_KEY, normalized);
  return normalized;
}

export interface AlwaysAllowController {
  /** Snapshot of the allowlisted tool names (copy; safe to mutate). */
  getTools: () => string[];
  hasTool: (toolName: string) => boolean;
  /** Add a tool. No-op when already present. Throws when the list is full. */
  addTool: (toolName: string) => Promise<void>;
  /** Remove a tool. No-op when absent. */
  removeTool: (toolName: string) => Promise<void>;
  /** Clear the allowlist. */
  clear: () => Promise<void>;
}

/**
 * Runtime controller for the persisted allowlist. Loads once at boot and
 * caches in memory; gates read synchronously, so Settings-page changes take
 * effect on the next tool call without a runtime refresh.
 */
export async function createAlwaysAllowController(options: {
  settings: AlwaysAllowStore;
}): Promise<AlwaysAllowController> {
  let tools = await getStoredAlwaysAllowedTools(options.settings);

  const persist = async (next: string[]): Promise<void> => {
    const previous = tools;
    tools = next;

    try {
      await setStoredAlwaysAllowedTools(options.settings, next);
    } catch (error) {
      // Never claim a grant that was not durably persisted.
      tools = previous;
      throw error;
    }
  };

  return {
    getTools: () => [...tools],
    hasTool: (toolName: string) => {
      const normalized = normalizeToolName(toolName);
      return normalized !== null && tools.includes(normalized);
    },
    addTool: async (toolName: string) => {
      const normalized = normalizeToolName(toolName);
      if (normalized === null || tools.includes(normalized)) {
        return;
      }

      if (tools.length >= MAX_ALWAYS_ALLOWED_TOOLS) {
        throw new Error(
          `Always-allow list is full (${MAX_ALWAYS_ALLOWED_TOOLS} tools).`,
        );
      }

      await persist([...tools, normalized]);
    },
    removeTool: async (toolName: string) => {
      const normalized = normalizeToolName(toolName);
      if (normalized === null || !tools.includes(normalized)) {
        return;
      }

      await persist(tools.filter((tool) => tool !== normalized));
    },
    clear: async () => {
      if (tools.length === 0) {
        return;
      }

      await persist([]);
    },
  };
}
