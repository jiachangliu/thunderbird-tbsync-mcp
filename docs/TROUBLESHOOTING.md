# Thunderbird TbSync MCP v5 troubleshooting

- **Helper unavailable.** Check that TbSync is the beta build, the matching
  native helper is installed, and the Bridge tab reports a running
  helper. The HTTP listener at 47654 belongs to the helper.
- **Connection file missing.** The owned XPI and the TbSync helper have
  separate endpoints. Check the XPI's startup log for its local port and
  use the XPI's `thunderbird-tbsync-mcp/connection.json` from Thunderbird's
  temp directory. Set `TBSYNC_MCP_CONNECTION_FILE` if your MCP client
  uses another temp directory. Keep this file private.
- **No grant or invalid scope.** Select the EAS account and its bound
  calendar in the TbSync Bridge tab. Leave unrestricted mode off.
- **Pre-sync refusal.** Check the calendar's status in TbSync Manager.
  Authentication errors, disabled resources, an unavailable provider,
  pending changes, or a stale sync result can prevent a write.
- **Post-sync failure or uncertain result.** Inspect the same item in
  Thunderbird and Outlook before retrying. The response reports whether
  the edit was accepted locally or its outcome is uncertain.
