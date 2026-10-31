/**
 * `/office` command grammar.
 *
 * One parser for the TUI, Pi Desktop's `commands.run`, and the tests. The
 * command is intentionally small: status, connect, close, restart, help. An
 * invalid argument is an error command rather than a thrown parse exception so
 * the handler can render the same help line everywhere.
 */

export type OfficeCommand =
  | { kind: "status" }
  | { kind: "connect"; port: number | null }
  | { kind: "close" }
  | { kind: "restart"; port: number | null }
  | { kind: "help" }
  | { kind: "error"; message: string };

/** One autocomplete suggestion, shaped like Pi's `AutocompleteItem`. */
export interface OfficeCommandCompletion {
  value: string;
  label: string;
  description?: string;
}

const COMMANDS: readonly {
  value: string;
  label: string;
  description: string;
}[] = [
  {
    value: "status",
    label: "status",
    description: "Show the bridge state, port and attached Office apps",
  },
  {
    value: "connect",
    label: "connect",
    description: "Start the bridge (optional port: /office connect 38617)",
  },
  {
    value: "close",
    label: "close",
    description: "Stop the bridge and disconnect the Office add-in",
  },
  {
    value: "restart",
    label: "restart",
    description: "Stop and start the bridge (optional port)",
  },
  {
    value: "help",
    label: "help",
    description: "Show the /office command list",
  },
] as const;

const ALIASES: Readonly<Record<string, string>> = {
  on: "connect",
  start: "connect",
  stop: "close",
  off: "close",
  disconnect: "close",
  info: "status",
};

function parsePort(raw: string): number | null {
  if (!/^\d+$/u.test(raw)) return null;
  const value = Number.parseInt(raw, 10);
  if (!Number.isInteger(value) || value <= 0 || value > 65535) return null;
  return value;
}

/** Parse the argument string after `/office`. Empty means `status`. */
export function parseOfficeCommand(args: string): OfficeCommand {
  const tokens = args.trim().split(/\s+/u).filter((token) => token.length > 0);
  const raw = tokens[0]?.toLowerCase();
  if (raw === undefined) return { kind: "status" };

  const name = ALIASES[raw] ?? raw;
  if (name === "status") {
    if (tokens.length > 1) {
      return { kind: "error", message: "Usage: /office status" };
    }
    return { kind: "status" };
  }
  if (name === "close") {
    if (tokens.length > 1) {
      return { kind: "error", message: "Usage: /office close" };
    }
    return { kind: "close" };
  }
  if (name === "help") {
    return { kind: "help" };
  }
  if (name === "connect" || name === "restart") {
    if (tokens.length > 2) {
      return {
        kind: "error",
        message: `Usage: /office ${name} [port]`,
      };
    }
    const portToken = tokens[1];
    if (portToken === undefined) return { kind: name, port: null };
    const port = parsePort(portToken);
    if (port === null) {
      return {
        kind: "error",
        message: `Invalid port "${portToken}": expected an integer from 1 to 65535.`,
      };
    }
    return { kind: name, port };
  }

  return {
    kind: "error",
    message: `Unknown subcommand "${tokens[0] ?? ""}". ${officeCommandHelp()}`,
  };
}

/** The command help block, shared by `/office help` and parse errors. */
export function officeCommandHelp(): string {
  return [
    "Office bridge commands:",
    "  /office status                 show state, port and attached apps",
    "  /office connect [port]         start the bridge (default 38617)",
    "  /office close                  stop the bridge",
    "  /office restart [port]         stop, then start",
    "  /office help                   this list",
    "",
    "The bridge is off until you connect. Only one Pi session can own the port.",
  ].join("\n");
}

/** Argument completions for `/office <prefix>`. */
export function officeCommandCompletions(
  prefix: string,
): OfficeCommandCompletion[] | null {
  const needle = prefix.trim().toLowerCase();
  const matches = COMMANDS.filter((entry) => entry.value.startsWith(needle));
  return matches.length > 0 ? matches : null;
}
