/**
 * Local-service ports — one shared dev/prod profile and the one-time repair of
 * stored legacy dev-port URLs.
 *
 * Under Node there is no Vite `import.meta.env`, so the guarded mode read
 * fails closed to the production profile. Ports are the same in both builds,
 * though; only the setup-command copy is mode-aware.
 */

import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";

import { BRIDGE_DEFAULT_PORT } from "@dieulc/pi-office-protocol";

import {
  BRIDGE_PORT,
  IS_DEV_BUILD,
  LOCAL_PROXY_URL,
  PI_BRIDGE_URL,
  PROXY_PORT,
  PYTHON_BRIDGE_URL,
  PYTHON_PORT,
  TMUX_BRIDGE_URL,
  TMUX_PORT,
  resolveProxySetupCommand,
  resolvePythonBridgeSetupCommand,
  resolveTmuxBridgeSetupCommand,
} from "../src/config/local-service-defaults.ts";
import {
  applyLegacyServiceUrlRepairs,
  planLegacyServiceUrlRepairs,
} from "../src/config/service-url-repair.ts";
import { DEFAULT_PI_BRIDGE_URL } from "../src/bridge/pi-bridge-manager.ts";
import {
  DEFAULT_PYTHON_BRIDGE_URL,
  DEFAULT_TMUX_BRIDGE_URL,
} from "../src/tools/experimental-tool-gates/types.ts";
import {
  DEFAULT_LOCAL_PROXY_URL,
  WPS_DEV_HOST_GATEWAY_PROXY_URL,
} from "../src/auth/proxy-validation.ts";

void test("Node/test runtime resolves the shared port profile", () => {
  assert.equal(IS_DEV_BUILD, false);
  assert.equal(BRIDGE_PORT, 38617);
  assert.equal(PROXY_PORT, 3003);
  assert.equal(PYTHON_PORT, 3340);
  assert.equal(TMUX_PORT, 3341);

  assert.equal(PI_BRIDGE_URL, "ws://127.0.0.1:38617");
  assert.equal(LOCAL_PROXY_URL, "https://localhost:3003");
  assert.equal(PYTHON_BRIDGE_URL, "https://localhost:3340");
  assert.equal(TMUX_BRIDGE_URL, "https://localhost:3341");

  assert.equal(DEFAULT_PI_BRIDGE_URL, PI_BRIDGE_URL);
  assert.equal(DEFAULT_LOCAL_PROXY_URL, LOCAL_PROXY_URL);
  assert.equal(DEFAULT_PYTHON_BRIDGE_URL, PYTHON_BRIDGE_URL);
  assert.equal(DEFAULT_TMUX_BRIDGE_URL, TMUX_BRIDGE_URL);
  assert.equal(
    WPS_DEV_HOST_GATEWAY_PROXY_URL,
    `http://10.0.2.2:${PROXY_PORT}`,
  );
});

void test("setup commands are dev-aware but bind the same ports", () => {
  assert.equal(
    resolvePythonBridgeSetupCommand(false),
    "npx pi-for-office-python-bridge",
  );
  assert.equal(
    resolvePythonBridgeSetupCommand(true),
    "npm run python:bridge:https",
  );
  assert.equal(
    resolveTmuxBridgeSetupCommand(false),
    "npx pi-for-office-tmux-bridge",
  );
  assert.equal(
    resolveTmuxBridgeSetupCommand(true),
    "npm run tmux:bridge:https",
  );
  assert.equal(resolveProxySetupCommand(false), "npx pi-for-office-proxy");
  assert.equal(resolveProxySetupCommand(true), "npm run proxy:https");
});

const LEGACY_STORED = {
  "pi-bridge.url": "ws://127.0.0.1:38618",
  "proxy.url": "https://localhost:3004",
  "python.bridge.url": "https://localhost:3350",
  "python.bridge.approved.url": "https://localhost:3350",
  "tmux.bridge.url": "https://localhost:3351",
} as const;

void test("legacy repair planner rewrites every old dev default", () => {
  assert.deepEqual(planLegacyServiceUrlRepairs(LEGACY_STORED), [
    { settingKey: "pi-bridge.url", value: PI_BRIDGE_URL },
    { settingKey: "proxy.url", value: LOCAL_PROXY_URL },
    { settingKey: "python.bridge.url", value: PYTHON_BRIDGE_URL },
    {
      settingKey: "python.bridge.approved.url",
      value: PYTHON_BRIDGE_URL,
    },
    { settingKey: "tmux.bridge.url", value: TMUX_BRIDGE_URL },
  ]);
});

void test("legacy repair planner handles the WPS host-gateway proxy", () => {
  assert.deepEqual(
    planLegacyServiceUrlRepairs({ "proxy.url": "http://10.0.2.2:3004" }),
    [{ settingKey: "proxy.url", value: `http://10.0.2.2:${PROXY_PORT}` }],
  );
});

void test("legacy repair planner never rewrites custom or current URLs", () => {
  assert.deepEqual(
    planLegacyServiceUrlRepairs({
      "pi-bridge.url": "ws://127.0.0.1:39999",
      "proxy.url": "https://my-proxy.example.com",
      "python.bridge.url": PYTHON_BRIDGE_URL,
      "python.bridge.approved.url": "https://localhost:3999",
      "tmux.bridge.url": "https://example.com:1234",
    }),
    [],
  );

  // Exact-match policy: a non-canonical legacy host is not rewritten, because
  // the old builds always wrote the canonical 127.0.0.1 URL.
  assert.deepEqual(
    planLegacyServiceUrlRepairs({ "pi-bridge.url": "ws://localhost:38618" }),
    [],
  );
});

void test("legacy repair planner is idempotent", () => {
  assert.deepEqual(
    planLegacyServiceUrlRepairs({
      "pi-bridge.url": PI_BRIDGE_URL,
      "proxy.url": LOCAL_PROXY_URL,
      "python.bridge.url": PYTHON_BRIDGE_URL,
      "python.bridge.approved.url": PYTHON_BRIDGE_URL,
      "tmux.bridge.url": TMUX_BRIDGE_URL,
    }),
    [],
  );
});

void test("legacy repair applier writes through the settings store", async () => {
  const writes: Record<string, string> = { ...LEGACY_STORED };
  const settings = {
    get: (key: string): Promise<string | null> =>
      Promise.resolve(writes[key] ?? null),
    set: (key: string, value: string): Promise<void> => {
      writes[key] = value;
      return Promise.resolve();
    },
  };

  await applyLegacyServiceUrlRepairs(settings);

  assert.equal(writes["pi-bridge.url"], PI_BRIDGE_URL);
  assert.equal(writes["proxy.url"], LOCAL_PROXY_URL);
  assert.equal(writes["python.bridge.url"], PYTHON_BRIDGE_URL);
  assert.equal(writes["python.bridge.approved.url"], PYTHON_BRIDGE_URL);
  assert.equal(writes["tmux.bridge.url"], TMUX_BRIDGE_URL);
});

void test("legacy repair applier leaves current values alone", async () => {
  const writes: Record<string, string> = {
    "pi-bridge.url": PI_BRIDGE_URL,
    "proxy.url": LOCAL_PROXY_URL,
    "python.bridge.url": PYTHON_BRIDGE_URL,
    "python.bridge.approved.url": PYTHON_BRIDGE_URL,
    "tmux.bridge.url": TMUX_BRIDGE_URL,
  };
  const seen: string[] = [];
  const settings = {
    get: (key: string): Promise<string | null> => {
      seen.push(key);
      return Promise.resolve(writes[key] ?? null);
    },
    set: (key: string, value: string): Promise<void> => {
      writes[key] = value;
      return Promise.resolve();
    },
  };

  await applyLegacyServiceUrlRepairs(settings);

  assert.equal(seen.length, 5, "each distinct setting key is read once");
  assert.deepEqual(writes, {
    "pi-bridge.url": PI_BRIDGE_URL,
    "proxy.url": LOCAL_PROXY_URL,
    "python.bridge.url": PYTHON_BRIDGE_URL,
    "python.bridge.approved.url": PYTHON_BRIDGE_URL,
    "tmux.bridge.url": TMUX_BRIDGE_URL,
  });
});

void test("the add-in bridge port matches the bridge server default", () => {
  assert.equal(BRIDGE_PORT, BRIDGE_DEFAULT_PORT);
});

void test("the port constants match the helper server scripts", async () => {
  const readDefaultPort = async (file: string): Promise<number> => {
    const source = await readFile(
      new URL(`../scripts/${file}`, import.meta.url),
      "utf8",
    );
    const match = /const DEFAULT_PORT = (\d+);/u.exec(source);
    assert.ok(match, `${file} must declare DEFAULT_PORT`);
    return Number.parseInt(match[1] as string, 10);
  };

  assert.equal(await readDefaultPort("cors-proxy-server.mjs"), PROXY_PORT);
  assert.equal(await readDefaultPort("python-bridge-server.mjs"), PYTHON_PORT);
  assert.equal(await readDefaultPort("tmux-bridge-server.mjs"), TMUX_PORT);
});
