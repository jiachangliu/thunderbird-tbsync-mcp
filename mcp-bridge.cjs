#!/usr/bin/env node

/**
 * MCP stdio <-> HTTP bridge for thunderbird-tbsync-mcp.
 *
 * The Thunderbird extension listens on localhost:8766.
 */

const http = require("http");

const PORT = process.env.TBSYNC_MCP_PORT ? Number(process.env.TBSYNC_MCP_PORT) : 8766;

function postJson(payload) {
  return new Promise((resolve, reject) => {
    const data = Buffer.from(JSON.stringify(payload), "utf8");
    const req = http.request(
      {
        hostname: "127.0.0.1",
        port: PORT,
        path: "/",
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Content-Length": data.length,
        },
      },
      (res) => {
        let chunks = "";
        res.setEncoding("utf8");
        res.on("data", (c) => (chunks += c));
        res.on("end", () => {
          try {
            resolve(JSON.parse(chunks));
          } catch (e) {
            reject(new Error(`Invalid JSON from extension: ${e.message}\n${chunks.slice(0, 500)}`));
          }
        });
      }
    );
    req.on("error", reject);
    req.write(data);
    req.end();
  });
}

let buffer = "";
process.stdin.setEncoding("utf8");

process.stdin.on("data", (chunk) => {
  buffer += chunk;
  let idx;
  while ((idx = buffer.indexOf("\n")) !== -1) {
    const line = buffer.slice(0, idx).trim();
    buffer = buffer.slice(idx + 1);
    if (!line) continue;

    let msg;
    try {
      msg = JSON.parse(line);
    } catch (e) {
      process.stdout.write(
        JSON.stringify({
          jsonrpc: "2.0",
          id: null,
          error: { code: -32700, message: `Parse error: ${e.message}` },
        }) + "\n"
      );
      continue;
    }

    postJson(msg)
      .then((resp) => process.stdout.write(JSON.stringify(resp) + "\n"))
      .catch((err) => {
        process.stdout.write(
          JSON.stringify({
            jsonrpc: "2.0",
            id: msg.id ?? null,
            error: { code: -32000, message: err.message },
          }) + "\n"
        );
      });
  }
});
