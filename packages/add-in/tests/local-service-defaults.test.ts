/**
 * Local-service defaults — dev/prod port profiles and the one-time dev
 * migration of stored prod-default URLs.
 *
 * Under Node there is no Vite `import.meta.env`, so the guarded mode read
 * fails closed to the production profile: every default must equal today's
 * prod value, and the dev profile must only be reachable through the pure
 * `isDev`-taking resolvers.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import {
  IS_DEV_BUILD,
  DEV_LOCAL_PROXY_URL,
  DEV_PI_BRIDGE_URL,
  DEV_PYTHON_BRIDGE_URL,
  DEV_TMUX_BRIDGE_URL,
  PROD_LOCAL_PROXY_URL,
  PROD_PI_BRIDGE_URL,
  PROD_PYTHON_BRIDGE_URL,
  PROD_TMUX_BRIDGE_URL,
  defaultBridgePort,
  resolveLocalProxyUrl,
  resolvePiBridgeUrl,
  resolveProxySetupCommand,
  resolvePythonBridgeSetupCommand,
  resolvePythonBridgeUrl,
  resolveTmuxBridgeSetupCommand,
  resolveTmuxBridgeUrl,
} from "../src/config/local-service-defaults.ts";
import {
  applyDevServiceDefaultMigrations,
  planDevServiceDefaultMigrations,
} from "../src/config/service-default-migration.ts";
import { DEFAULT_PI_BRIDGE_URL } from "../src/bridge/pi-bridge-manager.ts";
import {
  DEFAULT_PYTHON_BRIDGE_URL,
  DEFAULT_TMUX_BRIDGE_URL,
} from "../src/tools/experimental-tool-gates/types.ts";
import { DEFAULT_LOCAL_PROXY_URL } from "../src/auth/proxy-validation.ts";

void test("Node/test runtime resolves to the production profile", () => {
  assert.equal(IS_DEV_BUILD, false);
  assert.equal(defaultBridgePort(), 38617);
  assert.equal(
    DEFAULT_PI_BRIDGE_URL,
    PROD_PI_BRIDGE_URL,
    "bridge default must stay prod",
  );
  assert.equal(DEFAULT_PYTHON_BRIDGE_URL, PROD_PYTHON_BRIDGE_URL);
  assert.equal(DEFAULT_TMUX_BRIDGE_URL, PROD_TMUX_BRIDGE_URL);
  assert.equal(DEFAULT_LOCAL_PROXY_URL, PROD_LOCAL_PROXY_URL);
});

void test("dev/prod URL resolvers map to the port profile", () => {
  assert.equal(resolvePiBridgeUrl(false), "ws://127.0.0.1:38617");
  assert.equal(resolvePiBridgeUrl(true), "ws://127.0.0.1:38618");
  assert.equal(resolveLocalProxyUrl(false), "https://localhost:3003");
  assert.equal(resolveLocalProxyUrl(true), "https://localhost:3004");
  assert.equal(resolvePythonBridgeUrl(false), "https://localhost:3340");
  assert.equal(resolvePythonBridgeUrl(true), "https://localhost:3350");
  assert.equal(resolveTmuxBridgeUrl(false), "https://localhost:3341");
  assert.equal(resolveTmuxBridgeUrl(true), "https://localhost:3351");
});

void test("setup commands are dev-aware", () => {
  assert.equal(
    resolvePythonBridgeSetupCommand(false),
    "npx pi-for-office-python-bridge",
  );
  assert.equal(
    resolvePythonBridgeSetupCommand(true),
    "npm run python:bridge:dev:https",
  );
  assert.equal(
    resolveTmuxBridgeSetupCommand(false),
    "npx pi-for-office-tmux-bridge",
  );
  assert.equal(
    resolveTmuxBridgeSetupCommand(true),
    "npm run tmux:bridge:dev:https",
  );
  assert.equal(resolveProxySetupCommand(false), "npx pi-for-office-proxy");
  assert.equal(resolveProxySetupCommand(true), "npm run proxy:dev");
});

void test("dev migration planner rewrites only exact prod defaults", () => {
  const stored = {
    "pi-bridge.url": PROD_PI_BRIDGE_URL,
    "proxy.url": PROD_LOCAL_PROXY_URL,
    "python.bridge.url": PROD_PYTHON_BRIDGE_URL,
    "tmux.bridge.url": PROD_TMUX_BRIDGE_URL,
  };
  assert.deepEqual(planDevServiceDefaultMigrations(stored, true), [
    { settingKey: "pi-bridge.url", value: DEV_PI_BRIDGE_URL },
    { settingKey: "proxy.url", value: DEV_LOCAL_PROXY_URL },
    { settingKey: "python.bridge.url", value: DEV_PYTHON_BRIDGE_URL },
    { settingKey: "tmux.bridge.url", value: DEV_TMUX_BRIDGE_URL },
  ]);
});

void test("dev migration planner is a no-op without a dev build", () => {
  assert.deepEqual(
    planDevServiceDefaultMigrations(
      { "pi-bridge.url": PROD_PI_BRIDGE_URL },
      false,
    ),
    [],
  );
});

void test("dev migration planner never rewrites custom URLs", () => {
  assert.deepEqual(
    planDevServiceDefaultMigrations(
      {
        "pi-bridge.url": "ws://127.0.0.1:39999",
        "proxy.url": "https://my-proxy.example.com",
        "python.bridge.url": "https://localhost:3340/",
        "tmux.bridge.url": "https://example.com:1234",
      },
      true,
    ),
    [],
  );
});

void test("dev migration applier writes through the settings store", async () => {
  const writes: Record<string, string> = {};
  const settings = {
    get: async (key: string): Promise<string | null> => writes[key] ?? null,
    set: async (key: string, value: string): Promise<void> => {
      writes[key] = value;
    },
  };

  // Prime with the prod defaults, then migrate in a dev build.
  writes["pi-bridge.url"] = PROD_PI_BRIDGE_URL;
  writes["proxy.url"] = PROD_LOCAL_PROXY_URL;
  writes["python.bridge.url"] = PROD_PYTHON_BRIDGE_URL;
  writes["tmux.bridge.url"] = PROD_TMUX_BRIDGE_URL;

  await applyDevServiceDefaultMigrations(settings, true);

  assert.equal(writes["pi-bridge.url"], DEV_PI_BRIDGE_URL);
  assert.equal(writes["proxy.url"], DEV_LOCAL_PROXY_URL);
  assert.equal(writes["python.bridge.url"], DEV_PYTHON_BRIDGE_URL);
  assert.equal(writes["tmux.bridge.url"], DEV_TMUX_BRIDGE_URL);
});

void test("dev migration applier is a no-op in a prod build", async () => {
  const writes: Record<string, string> = {
    "pi-bridge.url": PROD_PI_BRIDGE_URL,
    "proxy.url": PROD_LOCAL_PROXY_URL,
  };
  const settings = {
    get: async (key: string): Promise<string | null> => writes[key] ?? null,
    set: async (key: string, value: string): Promise<void> => {
      writes[key] = value;
    },
  };

  await applyDevServiceDefaultMigrations(settings, false);

  assert.equal(writes["pi-bridge.url"], PROD_PI_BRIDGE_URL);
  assert.equal(writes["proxy.url"], PROD_LOCAL_PROXY_URL);
});
