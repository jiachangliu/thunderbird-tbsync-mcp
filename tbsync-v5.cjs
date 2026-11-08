"use strict";

const http = require("node:http");
const HELPER_TIMEOUT_MS = 5000;
const MAX_RESPONSE_BYTES = 2 * 1024 * 1024;

function callHelper(port, cmd, args = {}) {
  const payload = Buffer.from(JSON.stringify({ cmd, args, timeoutMs: HELPER_TIMEOUT_MS }));
  return new Promise((resolve, reject) => {
    const request = http.request({
      hostname: "127.0.0.1",
      port,
      path: "/rpc",
      method: "POST",
      // Upstream helper v2 requires this static header to block simple
      // cross-origin browser requests. It is not a local-process secret.
      headers: {
        Authorization: "Bearer tbsync",
        "Content-Type": "application/json",
        "Content-Length": payload.length,
      },
    }, response => {
      const chunks = [];
      let bytes = 0;
      response.on("data", chunk => {
        bytes += chunk.length;
        if (bytes > MAX_RESPONSE_BYTES) {
          request.destroy(new Error("TbSync response exceeded 2 MiB. Try a smaller date range."));
          return;
        }
        chunks.push(chunk);
      });
      response.on("error", reject);
      response.on("end", () => {
        if (response.statusCode !== 200) {
          reject(new Error(`TbSync helper returned HTTP ${response.statusCode} for ${cmd}.`));
          return;
        }
        try {
          const body = JSON.parse(Buffer.concat(chunks).toString("utf8"));
          if (body?.ok !== true || !Object.hasOwn(body, "result")) {
            const code = typeof body?.errorCode === "string" && /^E:[A-Z_]+$/.test(body.errorCode)
              ? ` (${body.errorCode})` : "";
            reject(new Error(`TbSync refused ${cmd}${code}. Check the Bridge tab.`));
            return;
          }
          resolve(body.result);
        } catch {
          reject(new Error("TbSync helper returned invalid JSON."));
        }
      });
    });
    const deadline = setTimeout(() => request.destroy(new Error(`TbSync timed out during ${cmd}.`)),
      HELPER_TIMEOUT_MS + 1000);
    request.once("close", () => clearTimeout(deadline));
    request.on("error", error => {
      reject(error.code === "ECONNREFUSED"
        ? new Error("TbSync v5 helper is unavailable. Check the Bridge tab in the TbSync beta build.")
        : error);
    });
    request.end(payload);
  });
}

function createV5Handler(port = 47654) {
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error("TBSYNC_V5_PORT must be a TCP port from 1 to 65535.");
  }
  const handler = import("./extension/mcp_server/v5.mjs").then(({ createV5Handler }) =>
    createV5Handler((cmd, args) => callHelper(port, cmd, args)));
  return async message => (await handler)(message);
}

module.exports = { createV5Handler };
