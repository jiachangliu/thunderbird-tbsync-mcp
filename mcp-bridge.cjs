#!/usr/bin/env node

/**
 * MCP stdio <-> HTTP bridge for thunderbird-tbsync-mcp.
 *
 * Legacy mode uses the Thunderbird extension on localhost:8766.
 * V5 mode uses the scoped TbSync beta native helper.
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

const backend = process.env.TBSYNC_MCP_BACKEND ?? "legacy";
if (backend !== "legacy" && backend !== "v5") {
  process.stderr.write("TBSYNC_MCP_BACKEND must be legacy or v5.\n");
  process.exit(1);
}
const dispatch = backend === "v5"
  ? require("./tbsync-v5.cjs").createV5Handler(
    process.env.TBSYNC_V5_PORT === undefined ? undefined : Number(process.env.TBSYNC_V5_PORT))
  : postJson;

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

    dispatch(msg)
      .then(resp => {
        if (resp !== null) process.stdout.write(JSON.stringify(resp) + "\n");
      })
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
