/* global ExtensionCommon, ChromeUtils, Components, Services, Cc, Ci, TextEncoder, XMLHttpRequest */
"use strict";

const RESOURCE_NAME = "thunderbird-tbsync-mcp-v5";
const RESOURCE_ROOT = `resource://${RESOURCE_NAME}/`;
const HELPER_URL = "http://127.0.0.1:47654/rpc";
const MAX_REQUEST_BYTES = 256 * 1024;
const MAX_HELPER_BYTES = 2 * 1024 * 1024;
const resourceProtocol = Cc["@mozilla.org/network/protocol;1?name=resource"]
  .getService(Ci.nsISubstitutingProtocolHandler);

function callHelper(cmd, args = {}, timeoutMs = 5000) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest({ mozAnon: true, mozSystem: true });
    xhr.open("POST", HELPER_URL);
    xhr.channel.QueryInterface(Ci.nsIHttpChannel).redirectionLimit = 0;
    xhr.timeout = timeoutMs + 1000;
    xhr.setRequestHeader("Authorization", "Bearer tbsync");
    xhr.setRequestHeader("Content-Type", "application/json");
    xhr.onprogress = event => {
      if (event.loaded > MAX_HELPER_BYTES) {
        reject(new Error("TbSync helper response exceeded 2 MiB."));
        xhr.abort();
      }
    };
    xhr.onerror = () => reject(new Error("TbSync beta helper is unavailable. Check the Bridge tab."));
    xhr.ontimeout = () => reject(new Error(`TbSync helper timed out during ${cmd}.`));
    xhr.onload = () => {
      if (xhr.status !== 200) {
        reject(new Error(`TbSync helper returned HTTP ${xhr.status} for ${cmd}.`));
        return;
      }
      try {
        const body = JSON.parse(xhr.responseText);
        if (body?.ok !== true || !Object.hasOwn(body, "result")) {
          const code = typeof body?.errorCode === "string" && /^E:[A-Z_]+$/.test(body.errorCode)
            ? ` (${body.errorCode})` : "";
          reject(new Error(`TbSync refused ${cmd}${code}. Check the Bridge tab and Event Log.`));
          return;
        }
        resolve(body.result);
      } catch {
        reject(new Error("TbSync helper returned invalid JSON."));
      }
    };
    xhr.send(JSON.stringify({ cmd, args, timeoutMs }));
  });
}

function authToken() {
  const rng = Cc["@mozilla.org/security/random-generator;1"].createInstance(Ci.nsIRandomGenerator);
  return Array.from(rng.generateRandomBytes(32), byte => byte.toString(16).padStart(2, "0")).join("");
}

function matchesToken(header, expected) {
  if (typeof header !== "string" || header.length !== expected.length) return false;
  let diff = 0;
  for (let i = 0; i < expected.length; i++) diff |= header.charCodeAt(i) ^ expected.charCodeAt(i);
  return diff === 0;
}

function publishConnection(port, token) {
  const directory = Services.dirsvc.get("TmpD", Ci.nsIFile);
  directory.append("thunderbird-tbsync-mcp");
  if (!directory.exists()) directory.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
  if (!directory.isDirectory() || directory.isSymlink() ||
      (Services.appinfo.OS !== "WINNT" && (directory.permissions & 0o077))) {
    throw new Error("The TbSync MCP temp directory must be private and cannot be a symlink.");
  }
  const file = directory.clone();
  file.append("connection.json");
  if (file.exists() && (!file.isFile() || file.isSymlink())) {
    throw new Error("The TbSync MCP connection file is not a regular file.");
  }
  const stream = Cc["@mozilla.org/network/safe-file-output-stream;1"].createInstance(Ci.nsIFileOutputStream);
  stream.init(file, 0x02 | 0x08 | 0x20, 0o600, 0);
  const value = JSON.stringify({ port, token, pid: Services.appinfo.processID });
  try {
    if (stream.write(value, value.length) !== value.length) throw new Error("Incomplete connection file write.");
    stream.QueryInterface(Ci.nsISafeOutputStream).finish();
  } catch (error) {
    stream.close();
    throw error;
  }
}

function finish(res, status, statusText, body) {
  res.setStatusLine("1.1", status, statusText);
  res.setHeader("Content-Type", "application/json; charset=utf-8", false);
  res.setHeader("Cache-Control", "no-store", false);
  if (body !== undefined) {
    // httpd's string writer narrows UTF-16 code units to bytes. The helper
    // already supplies Unicode strings, so encode the complete JSON once.
    const bytes = new TextEncoder().encode(JSON.stringify(body));
    const output = Cc["@mozilla.org/binaryoutputstream;1"].createInstance(Ci.nsIBinaryOutputStream);
    output.setOutputStream(res.bodyOutputStream);
    output.writeByteArray(bytes);
  }
  res.finish();
}

function registerEndpoint(server, token, handle, NetUtil) {
  server.registerPathHandler("/", (req, res) => {
    res.processAsync();
    if (req.method !== "POST") {
      finish(res, 405, "Method Not Allowed");
      return;
    }
    const header = req.hasHeader("Authorization") ? req.getHeader("Authorization") : "";
    if (!matchesToken(header, `Bearer ${token}`)) {
      finish(res, 401, "Unauthorized");
      return;
    }

    const stream = req.bodyInputStream;
    const length = stream.available();
    if (length > MAX_REQUEST_BYTES) {
      finish(res, 413, "Request Too Large");
      return;
    }
    let message;
    try {
      message = JSON.parse(NetUtil.readInputStreamToString(stream, length, { charset: "UTF-8" }));
    } catch {
      finish(res, 400, "Bad Request", { jsonrpc: "2.0", id: null, error: { code: -32700, message: "Invalid JSON." } });
      return;
    }
    (async () => {
      let result;
      try {
        result = await handle(message);
      } catch {
        result = { jsonrpc: "2.0", id: message?.id ?? null,
          error: { code: -32000, message: "TbSync MCP request failed." } };
      }
      if (result === null) finish(res, 204, "No Content");
      else finish(res, 200, "OK", result);
    })();
  });
}

var tbsyncMcpServer = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    resourceProtocol.setSubstitution(RESOURCE_NAME, context.extension.rootURI);
    return {
      tbsyncMcpServer: {
        start: async () => {
          if (this.server) return { success: true, port: this.server.identity.primaryPort };
          try {
            Components.utils.importGlobalProperties(["TextEncoder", "XMLHttpRequest"]);
            const { HttpServer } = ChromeUtils.importESModule(`${RESOURCE_ROOT}httpd.sys.mjs`);
            const { NetUtil } = ChromeUtils.importESModule("resource://gre/modules/NetUtil.sys.mjs");
            const { default: ICAL } = ChromeUtils.importESModule("resource:///modules/calendar/Ical.sys.mjs");
            const { createEventCodec } = ChromeUtils.importESModule(`${RESOURCE_ROOT}mcp_server/event-ical.mjs`);
            const { createV5Handler } = ChromeUtils.importESModule(`${RESOURCE_ROOT}mcp_server/v5.mjs`);
            const token = authToken();
            const { setTimeout } = ChromeUtils.importESModule("resource://gre/modules/Timer.sys.mjs");
            const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
            const handle = createV5Handler(callHelper, { eventCodec: createEventCodec(ICAL), wait });
            const server = new HttpServer();
            registerEndpoint(server, token, handle, NetUtil);
            server.start(-1);
            try {
              publishConnection(server.identity.primaryPort, token);
            } catch (error) {
              await server.stop();
              throw error;
            }
            this.server = server;
            return { success: true, port: server.identity.primaryPort };
          } catch (error) {
            return { success: false, error: error.message };
          }
        },
      },
    };
  }

  onShutdown() {
    if (this.server) this.server.stop();
    this.server = null;
    resourceProtocol.setSubstitution(RESOURCE_NAME, null);
  }
};
