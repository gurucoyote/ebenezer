# MCP Readiness Summary

To bring up an MCP server quickly after the metadata/logging work lands, we need:

1. **Session Manager**: Replace the global `appState` singleton with a session factory and map (session ID → `*app.State`). MCP will create/destroy sessions while the CLI keeps using the singleton factory.
2. **Action Facade**: Build MCP handlers that translate RPC requests into action invocations (using the metadata for validation). Example: `cursor/set` → `actions.Goto`, `cell/edit` → `actions.Edit`.
3. **Discovery Endpoint**: Expose an MCP method (e.g., `actions/list`) that returns the new metadata schema so clients know available commands/args. `actions.Discover()` already serializes the registry into this JSON payload, so the MCP surface only needs to proxy the helper.
4. **Structured Responses**: Ensure `actions.Result.Data` contains machine-readable payloads for MCP-friendly commands (info, search results, etc.). CLI will still print `Result.Message`.
5. **Logging Integration**: Wire MCP server to pass its own logger into `actions.Context`, capturing session IDs and RPC metadata.
6. **Security/Sandboxing**: Leverage metadata (idempotent flag, categories) to enforce policy (e.g., block destructive actions in read-only mode) and tie into existing sandbox settings.

Implementation mirrors `/root/roderik/cmd/mcp.go` by reusing `github.com/mark3labs/mcp-go/mcp` + `server` helpers so logging, tool registration, and transport semantics match the existing automation stack. Phase P1 tools (`workspace_open/close`, `workbook_save`, `cursor_get/set`, `info_get`) are live alongside the full P2 editing set (`cell_edit`, `cell_clear`, `row_insert_above`, `row_insert_below`, `row_delete`, `column_insert_left/right`, `column_delete`) and the P3 utilities (`selection_set`, `selection_clear`, `selection_export`, `clipboard_get`, `clipboard_set`). Selection exports now support CSV/XLSX destinations and emit inline style metadata, while the first wave of P4 styling tools (`style_describe`, `style_apply`, enriched `info_get`) unlock remote formatting workflows; the remaining styling semantics work focuses on workbook annotations and policy hooks.

Once metadata/logging are complete, the MCP MVP is essentially:
- `pkg/mcp/server.go` with session manager + dispatcher.
- JSON schema docs for metadata and action results (sample below).
- Tests covering `actions/list`, `cursor/get/set`, `cell/edit`, `save`, leveraging the new action layer.

### Example `actions/list` payload
```json
{
  "actions": [
    {
      "name": "info",
      "description": "Summarize workbook metadata (sheets, dimensions, cursor, file info)",
      "category": "metadata",
      "args": [
        {"name": "path", "description": "Optional file to inspect", "optional": true},
        {"name": "--details", "description": "Include file size/timestamps", "optional": true}
      ],
      "idempotent": true
    }
  ]
}
```

The same payload is available locally via `./ebenezer actions list`, which is helpful when validating metadata changes during development. Remote agents should call the MCP `actions_list` tool exposed by `./ebenezer mcp serve`; both surfaces are thin wrappers over `actions.Discover()`, guaranteeing they stay in sync as the registry evolves.

Focus now: land the metadata/logging work (M1–M5), then pivot to session management + MCP RPC surface.
