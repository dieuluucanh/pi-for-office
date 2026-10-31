/** Unit tests for the `/office` command grammar (bridge 0.6+). */

import assert from "node:assert/strict";
import { test } from "node:test";

import {
  officeCommandCompletions,
  officeCommandHelp,
  parseOfficeCommand,
} from "../dist/commands.js";

test("empty args show status, preserving the old /office behaviour", () => {
  assert.deepEqual(parseOfficeCommand(""), { kind: "status" });
  assert.deepEqual(parseOfficeCommand("   "), { kind: "status" });
  assert.deepEqual(parseOfficeCommand("status"), { kind: "status" });
  assert.deepEqual(parseOfficeCommand("info"), { kind: "status" });
});

test("connect accepts no port, an explicit port and aliases", () => {
  assert.deepEqual(parseOfficeCommand("connect"), { kind: "connect", port: null });
  assert.deepEqual(parseOfficeCommand("connect 38617"), { kind: "connect", port: 38617 });
  assert.deepEqual(parseOfficeCommand("on 38620"), { kind: "connect", port: 38620 });
  assert.deepEqual(parseOfficeCommand("start"), { kind: "connect", port: null });
  assert.deepEqual(parseOfficeCommand("  CONNECT   38617  "), {
    kind: "connect",
    port: 38617,
  });
});

test("close understands the stop/off/disconnect aliases", () => {
  for (const command of ["close", "stop", "off", "disconnect"]) {
    assert.deepEqual(parseOfficeCommand(command), { kind: "close" }, command);
  }
});

test("restart and help are parsed", () => {
  assert.deepEqual(parseOfficeCommand("restart"), { kind: "restart", port: null });
  assert.deepEqual(parseOfficeCommand("restart 38620"), { kind: "restart", port: 38620 });
  assert.deepEqual(parseOfficeCommand("help"), { kind: "help" });
});

test("invalid ports are error commands with a range hint", () => {
  for (const port of ["0", "65536", "-1", "1.5", "abc"]) {
    const result = parseOfficeCommand(`connect ${port}`);
    assert.equal(result.kind, "error", port);
    assert.match(result.message, /1 to 65535/, port);
  }
});

test("extra arguments and unknown subcommands are errors", () => {
  assert.equal(parseOfficeCommand("status extra").kind, "error");
  assert.equal(parseOfficeCommand("close extra").kind, "error");
  assert.equal(parseOfficeCommand("connect 38617 extra").kind, "error");

  const unknown = parseOfficeCommand("pause");
  assert.equal(unknown.kind, "error");
  assert.match(unknown.message, /Unknown subcommand "pause"/);
  assert.match(unknown.message, /\/office connect/);
});

test("help text names every command and the one-owner rule", () => {
  const help = officeCommandHelp();
  for (const name of ["status", "connect", "close", "restart", "help"]) {
    assert.match(help, new RegExp(`/office ${name}`), name);
  }
  assert.match(help, /Only one Pi session/);
});

test("completions filter by prefix and return null when nothing matches", () => {
  const all = officeCommandCompletions("");
  assert.ok(all);
  assert.deepEqual(
    all.map((entry) => entry.value),
    ["status", "connect", "close", "restart", "help"],
  );
  const c = officeCommandCompletions("c");
  assert.ok(c);
  assert.deepEqual(
    c.map((entry) => entry.value),
    ["connect", "close"],
  );
  assert.equal(officeCommandCompletions("zzz"), null);
});
