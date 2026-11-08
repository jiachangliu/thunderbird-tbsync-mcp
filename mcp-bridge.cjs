#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const http = require("node:http");
const os = require("node:os");
const path = require("node:path");
const readline = require("node:readline");

const MAX_RESPONSE_BYTES = 4 * 1024 * 1024;
const REQUEST_TIMEOUT_MS = 240000;
const backend = process.env.TBSYNC_MCP_BACKEND ?? "extension";
if (backend !== "extension" && backend !== "v5") {
  process.stderr.write("TBSYNC_MCP_BACKEND must be extension or v5 (direct read-only mode).\n");
  process.exit(1);
}

function readConnection() {
  const file = process.env.TBSYNC_MCP_CONNECTION_FILE ??
    path.join(os.tmpdir(), "thunderbird-tbsync-mcp", "connection.json");
  try {
    const directory = fs.lstatSync(path.dirname(file));
    const uid = process.getuid?.();
    if (!directory.isDirectory() || (uid !== undefined && (directory.uid !== uid || (directory.mode & 0o077)))) {
      throw new Error();
    }
    const fd = fs.openSync(file, fs.constants.O_RDONLY | fs.constants.O_NOFOLLOW | fs.constants.O_NONBLOCK);
    try {
      const stat = fs.fstatSync(fd);
      if (!stat.isFile() || stat.size > 4096 || (uid !== undefined && (stat.uid !== uid || (stat.mode & 0o077)))) {
        throw new Error();
      }
      const { port, token } = JSON.parse(fs.readFileSync(fd, "utf8"));
      if (!Number.isInteger(port) || port < 1 || port > 65535 ||
          typeof token !== "string" || !/^[0-9a-f]{64}$/.test(token)) {
        throw new Error();
      }
      return { port, token };
    } finally {
      fs.closeSync(fd);
    }
  } catch {
    throw new Error("The authenticated TbSync MCP connection file is missing or invalid. Check the v5 extension and TMPDIR.");
  }
}

async function postJson(payload) {
  const { port, token } = readConnection();
  const data = Buffer.from(JSON.stringify(payload), "utf8");
  return new Promise((resolve, reject) => {
    const req = http.request({
      hostname: "127.0.0.1", port, path: "/", method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        "Content-Length": data.length,
      },
    }, res => {
      const chunks = [];
      let bytes = 0;
      res.on("data", chunk => {
        bytes += chunk.length;
        if (bytes > MAX_RESPONSE_BYTES) {
          reject(new Error("TbSync MCP response exceeded 4 MiB."));
          req.destroy();
          return;
        }
        chunks.push(chunk);
      });
      res.on("error", reject);
      res.on("end", () => {
        if (res.statusCode === 204) {
          resolve(null);
        } else if (res.statusCode !== 200) {
          reject(new Error(`TbSync MCP extension returned HTTP ${res.statusCode}.`));
        } else {
          try {
            resolve(JSON.parse(Buffer.concat(chunks).toString("utf8")));
          } catch {
            reject(new Error("Invalid JSON from the TbSync MCP extension."));
          }
        }
      });
    });
    const deadline = setTimeout(() => req.destroy(new Error(
      "TbSync MCP request timed out. Check the calendar and Outlook before retrying a write.")), REQUEST_TIMEOUT_MS);
    req.once("close", () => clearTimeout(deadline));
    req.on("error", reject);
    req.end(data);
  });
}

const dispatch = backend === "v5"
  ? require("./tbsync-v5.cjs").createV5Handler(
    process.env.TBSYNC_V5_PORT === undefined ? undefined : Number(process.env.TBSYNC_V5_PORT))
  : postJson;

const lines = readline.createInterface({ input: process.stdin, crlfDelay: Infinity });
const reply = message => process.stdout.write(JSON.stringify(message) + "\n");

lines.on("line", async line => {
  if (!line.trim()) return;
  let message;
  try {
    message = JSON.parse(line);
  } catch {
    reply({ jsonrpc: "2.0", id: null, error: { code: -32700, message: "Invalid JSON." } });
    return;
  }
  try {
    const response = await dispatch(message);
    if (response !== null) reply(response);
  } catch (error) {
    if (message && typeof message === "object" && !Object.hasOwn(message, "id")) return;
    reply({ jsonrpc: "2.0", id: message?.id ?? null, error: { code: -32000, message: error.message } });
  }
});
