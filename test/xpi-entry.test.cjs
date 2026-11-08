"use strict";

const assert = require("node:assert/strict");
const { readFileSync } = require("node:fs");
const path = require("node:path");
const { test } = require("node:test");
const vm = require("node:vm");

test("packaged experiment publishes a private endpoint and requires its session token", async () => {
  const core = await import("../extension/mcp_server/v5.mjs");
  const code = readFileSync(path.join(__dirname, "../extension/mcp_server/api.js"), "utf8");
  let endpoint;
  let published;
  let stopped = false;
  const calls = [];
  class HttpServer {
    identity = { primaryPort: 19101 };
    registerPathHandler(route, handler) {
      assert.equal(route, "/");
      endpoint = handler;
    }
    start(port) { assert.equal(port, -1); }
    stop() { stopped = true; }
  }
  const stream = {
    init(_file, _flags, mode) { assert.equal(mode, 0o600); },
    write(value, length) { published = JSON.parse(value); return length; },
    QueryInterface() { return this; },
    finish() {},
  };
  class XHR {
    channel = { QueryInterface: () => this.channel };
    open(method, url) {
      assert.equal(method, "POST");
      assert.equal(url, "http://127.0.0.1:47654/rpc");
    }
    setRequestHeader(name, value) {
      if (name === "Authorization") assert.equal(value, "Bearer tbsync");
    }
    send(payload) {
      calls.push(JSON.parse(payload));
      this.status = 200;
      this.responseText = JSON.stringify({ ok: true, result: { helperVersion: 2, unrestricted: false,
        target: { exists: true, accountId: "fixture", accountName: "Test" }, resources: {} } });
      this.onload();
    }
  }
  const dir = { append() {}, exists: () => true, isDirectory: () => true, isSymlink: () => false,
    permissions: 0o700, clone: () => ({ append() {}, exists: () => false }) };
  const sandbox = {
    ExtensionCommon: { ExtensionAPI: class {} },
    Ci: { nsIFile: {} },
    Cc: {
      "@mozilla.org/network/protocol;1?name=resource": { getService: () => ({ setSubstitution() {} }) },
      "@mozilla.org/security/random-generator;1": { createInstance: () => ({
        generateRandomBytes: size => Buffer.alloc(size, 1),
      }) },
      "@mozilla.org/network/safe-file-output-stream;1": { createInstance: () => stream },
    },
    Components: { utils: { importGlobalProperties() {} } },
    Services: { dirsvc: { get: () => dir }, appinfo: { OS: "Linux", processID: 1 } },
    XMLHttpRequest: XHR,
    ChromeUtils: { importESModule(uri) {
      const name = uri.split("/").at(-1);
      switch (name) {
        case "httpd.sys.mjs": return { HttpServer };
        case "NetUtil.sys.mjs": return { NetUtil: { readInputStreamToString: input => input.text } };
        case "Ical.sys.mjs": return { default: {} };
        case "event-ical.mjs": return { createEventCodec: () => ({}) };
        case "v5.mjs": return core;
        case "Timer.sys.mjs": return { setTimeout };
        default: assert.fail(`Unexpected import ${uri}`);
      }
    } },
  };
  vm.createContext(sandbox);
  vm.runInContext(code, sandbox);
  const instance = new sandbox.tbsyncMcpServer();
  const api = instance.getAPI({ extension: { rootURI: "fixture" } }).tbsyncMcpServer;
  const started = await api.start();
  assert.equal(started.success, true);
  assert.equal(published.port, 19101);
  assert.match(published.token, /^[0-9a-f]{64}$/);

  function request(message, token) {
    const text = JSON.stringify(message);
    return new Promise(resolve => {
      let code, body;
      endpoint({
        method: "POST",
        hasHeader: () => token !== undefined,
        getHeader: () => `Bearer ${token}`,
        bodyInputStream: { available: () => text.length, text },
      }, {
        processAsync() {},
        setStatusLine(_version, status) { code = status; },
        setHeader() {},
        write(value) { body = JSON.parse(value); },
        finish() { resolve({ code, body }); },
      });
    });
  }

  const list = { jsonrpc: "2.0", id: 1, method: "tools/list" };
  assert.equal((await request(list)).code, 401);
  assert.equal((await request(list, "incorrect")).code, 401);
  const authorized = await request(list, published.token);
  assert.equal(authorized.code, 200);
  assert.ok(authorized.body.result.tools.some(tool => tool.name === "createCalendarEvent"));
  const accounts = await request({ ...list, method: "tools/call", params: { name: "tbsyncListAccounts" } }, published.token);
  assert.deepEqual(JSON.parse(accounts.body.result.content[0].text), [{ accountID: "fixture", accountname: "Test" }]);
  assert.deepEqual(calls, [{ cmd: "status", args: {}, timeoutMs: 5000 }]);
  instance.onShutdown();
  assert.equal(stopped, true);
});
