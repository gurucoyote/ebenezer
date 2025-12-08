# MCP Readiness Summary

To bring up an MCP server quickly after the metadata/logging work lands, we need:

1. **Session Manager**: Replace the global `appState` singleton with a session factory and map (session ID → `*app.State`). MCP will create/destroy sessions while the CLI keeps using the singleton factory.
2. **Action Facade**: Build MCP handlers that translate RPC requests into action invocations (using the metadata for validation). Example: `cursor/set` → `actions.Goto`, `cell/edit` → `actions.Edit`.
3. **Discovery Endpoint**: Expose an MCP method (e.g., `actions/list`) that returns the new metadata schema so clients know available commands/args.
4. **Structured Responses**: Ensure `actions.Result.Data` contains machine-readable payloads for MCP-friendly commands (info, search results, etc.). CLI will still print `Result.Message`.
5. **Logging Integration**: Wire MCP server to pass its own logger into `actions.Context`, capturing session IDs and RPC metadata.
6. **Security/Sandboxing**: Leverage metadata (idempotent flag, categories) to enforce policy (e.g., block destructive actions in read-only mode) and tie into existing sandbox settings.

Once metadata/logging are complete, the MCP MVP is essentially:
- `pkg/mcp/server.go` with session manager + dispatcher.
- JSON schema docs for metadata and action results.
- Tests covering `actions/list`, `cursor/get/set`, `cell/edit`, `save`, leveraging the new action layer.

Focus now: land the metadata/logging work (M1–M5), then pivot to session management + MCP RPC surface.
