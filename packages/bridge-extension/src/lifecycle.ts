/**
 * Bridge lifecycle controller — the one owner of the bridge server's start/stop
 * state.
 *
 * The extension used to start the WebSocket server unconditionally on
 * `session_start`, so every Pi process claimed the port. This controller keeps
 * the server off until the user asks for it, makes `connect`/`close` idempotent,
 * coalesces concurrent starts, and guarantees a failed start never leaves a
 * half-open server behind.
 *
 * It is deliberately free of Pi imports: the extension owns the handlers and
 * passes a `createServer` factory, so this module is unit-testable without a
 * live session.
 */

/** The minimal server surface the controller drives. */
export interface BridgeServerLike {
  readonly isRunning: boolean;
  readonly actualPort: number | null;
  attachedPanes(): readonly unknown[];
  start(): Promise<void>;
  stop(): Promise<void>;
}

export type BridgeLifecycleState = "off" | "starting" | "on" | "closing";

/** A point-in-time view for the status line, commands and tests. */
export interface BridgeSnapshot {
  state: BridgeLifecycleState;
  /** The port this connection asked for (or the resolved default when off). */
  port: number;
  /** The bound port while on, otherwise null. */
  actualPort: number | null;
  /** Attached panes while on, otherwise 0. */
  panes: number;
  /** Why the last start/stop failed, or null. */
  lastError: string | null;
  /** Epoch ms of the successful start, or null when off. */
  startedAt: number | null;
}

export interface BridgeControllerDeps<S extends BridgeServerLike> {
  /** Build the server for one port. Called only when a new connection starts. */
  createServer(port: number): S;
  /** The port used when `connect()` is called without one. */
  defaultPort: number;
  /** Injectable clock (tests use a fixed one). */
  now?: () => number;
  /** Called after every transition. Must not throw. */
  onChange?: (snapshot: BridgeSnapshot) => void;
}

export interface BridgeController<S extends BridgeServerLike> {
  snapshot(): BridgeSnapshot;
  /** The live server, or null when off. */
  server(): S | null;
  connect(port?: number | null): Promise<BridgeSnapshot>;
  close(): Promise<BridgeSnapshot>;
  restart(port?: number | null): Promise<BridgeSnapshot>;
}

/** A start/stop failure with a user-actionable message. */
export class BridgeLifecycleError extends Error {
  readonly code: string | null;

  constructor(message: string, code: string | null = null) {
    super(message);
    this.name = "BridgeLifecycleError";
    this.code = code;
  }
}

function errorCode(error: unknown): string | null {
  const code = (error as NodeJS.ErrnoException | null)?.code;
  return typeof code === "string" && code.length > 0 ? code : null;
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/**
 * Map a bind failure to the sentence the command surface shows. The port-owner
 * hint is the whole point: a silent `EADDRINUSE` is what made the old
 * auto-start confusing.
 */
export function describeBridgeStartError(error: unknown, port: number): string {
  const code = errorCode(error);
  const detail = errorMessage(error);
  if (code === "EADDRINUSE") {
    return (
      `Office bridge could not start: port ${port} is already in use. ` +
      `Another Pi session or program owns it. Run /office status there and close it, ` +
      `or connect this session on another port with /office connect <port>.`
    );
  }
  if (code === "EACCES") {
    return `Office bridge could not start on :${port}: permission denied (${detail}).`;
  }
  return `Office bridge could not start on :${port}: ${detail}`;
}

/** Normalize a command-supplied port; null means "use the default". */
export function normalizeBridgePort(value: number | null | undefined): number | null {
  if (value === null || value === undefined) return null;
  if (!Number.isInteger(value) || value <= 0 || value > 65535) {
    throw new BridgeLifecycleError(
      `Invalid port ${String(value)}: expected an integer from 1 to 65535.`,
      "bad_port",
    );
  }
  return value;
}

export function createBridgeController<S extends BridgeServerLike>(
  deps: BridgeControllerDeps<S>,
): BridgeController<S> {
  const now = deps.now ?? ((): number => Date.now());

  let state: BridgeLifecycleState = "off";
  let port = deps.defaultPort;
  let actualPort: number | null = null;
  let lastError: string | null = null;
  let startedAt: number | null = null;
  let server: S | null = null;
  let inFlight: Promise<BridgeSnapshot> | null = null;

  const snapshot = (): BridgeSnapshot => ({
    state,
    port,
    actualPort,
    panes: state === "on" && server ? server.attachedPanes().length : 0,
    lastError,
    startedAt,
  });

  const emit = (): BridgeSnapshot => {
    const next = snapshot();
    try {
      deps.onChange?.(next);
    } catch {
      // A status-line listener must never break a lifecycle transition.
    }
    return next;
  };

  const requireDefaultPort = (): number => {
    const resolved = normalizeBridgePort(deps.defaultPort) ?? deps.defaultPort;
    if (!Number.isInteger(resolved) || resolved <= 0 || resolved > 65535) {
      throw new BridgeLifecycleError(
        `Invalid default bridge port ${String(deps.defaultPort)}.`,
        "bad_port",
      );
    }
    return resolved;
  };

  const start = async (requested: number | null): Promise<BridgeSnapshot> => {
    const target = requested ?? requireDefaultPort();
    if (state === "on") {
      if (target === port) return snapshot();
      throw new BridgeLifecycleError(
        `Office bridge is already listening on :${actualPort ?? port}. ` +
          `Run /office restart ${target} to move it.`,
        "already_running",
      );
    }
    // `connect` sequences a concurrent start; `restart` only reaches here
    // after `stop` has completed. Keep this function single-flight by contract.
    port = target;
    lastError = null;
    state = "starting";
    const created = deps.createServer(target);
    server = created;
    emit();

    try {
      await created.start();
      actualPort = created.actualPort;
      startedAt = now();
      state = "on";
      return emit();
    } catch (error) {
      const message = describeBridgeStartError(error, target);
      // A failed listen can leave the server object's `wss`/`httpServer` set;
      // never keep it, and never claim `on`.
      server = null;
      actualPort = null;
      startedAt = null;
      state = "off";
      lastError = message;
      emit();
      throw new BridgeLifecycleError(message, errorCode(error));
    }
  };

  const stop = async (): Promise<BridgeSnapshot> => {
    if (state === "off" && server === null) return snapshot();
    if (state === "starting") {
      // Let the in-flight start settle; if it failed, there is nothing to stop.
      await inFlight?.catch(() => undefined);
    }
    if (server === null) {
      state = "off";
      actualPort = null;
      startedAt = null;
      return emit();
    }

    state = "closing";
    emit();
    const active = server;
    server = null;
    try {
      await active.stop();
      actualPort = null;
      startedAt = null;
      state = "off";
      return emit();
    } catch (error) {
      const message = `Office bridge did not close cleanly: ${errorMessage(error)}`;
      actualPort = null;
      startedAt = null;
      state = "off";
      lastError = message;
      emit();
      throw new BridgeLifecycleError(message, errorCode(error));
    }
  };

  return {
    snapshot,
    server: () => server,
    connect(portValue = null) {
      const requested = normalizeBridgePort(portValue);
      if (state === "starting") {
        return (inFlight ?? Promise.resolve(snapshot())).then((current) => {
          if (current.state === "on" && (requested === null || current.port === requested)) {
            return current;
          }
          if (current.state === "on") {
            throw new BridgeLifecycleError(
              `Office bridge is already listening on :${current.actualPort ?? current.port}. ` +
                `Run /office restart ${requested} to move it.`,
              "already_running",
            );
          }
          return start(requested);
        });
      }
      if (state === "on") {
        if (requested === null || requested === port) {
          return Promise.resolve(snapshot());
        }
        return Promise.reject(
          new BridgeLifecycleError(
            `Office bridge is already listening on :${actualPort ?? port}. ` +
              `Run /office restart ${requested} to move it.`,
            "already_running",
          ),
        );
      }
      if (state === "closing") {
        return Promise.reject(
          new BridgeLifecycleError(
            "Office bridge is still closing; try again once it has stopped.",
            "closing",
          ),
        );
      }
      inFlight = start(requested).finally(() => {
        inFlight = null;
      });
      return inFlight;
    },
    close() {
      if (state === "closing") return Promise.resolve(snapshot());
      inFlight = stop().finally(() => {
        inFlight = null;
      });
      return inFlight;
    },
    async restart(portValue = null) {
      const requested = normalizeBridgePort(portValue);
      await stop().catch(() => undefined);
      return start(requested);
    },
  };
}
