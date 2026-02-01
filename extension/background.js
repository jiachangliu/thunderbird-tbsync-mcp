/* global browser */

async function init() {
  try {
    const result = await browser.tbsyncMcpServer.start();
    if (result && result.success) {
      console.log("TbSync MCP server started on port", result.port);
    } else {
      console.error("Failed to start TbSync MCP server:", result && result.error);
    }
  } catch (e) {
    console.error("Error starting TbSync MCP server:", e);
  }
}

browser.runtime.onInstalled.addListener(init);
browser.runtime.onStartup.addListener(init);
