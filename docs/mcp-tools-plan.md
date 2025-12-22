# MCP Tooling Implementation Plan (v1)

This plan outlines the staged rollout of Ebenezer’s MCP surface so we can serve MCP user stories (US-04, US-09, US-10) without destabilizing the CLI.

## Objectives
1. **Session Safety**: Replace the global singleton with a light session manager so MCP clients can open and isolate workbooks.
2. **Core Editing Surface**: Expose the minimum viable set of workspace/cursor/cell tools required for remote editing (open, inspect, edit, save).
3. **Selection & Styling Awareness**: Extend MCP endpoints with range selection, exports, and style semantics mandated by US-05/US-09/US-10.
4. **Observability**: Emit structured logs/telemetry for every MCP call so agents can audit edits.

## Phased Rollout

| Phase | Scope | Key Tools | Notes | Status |
|-------|-------|-----------|-------|--------|
| P0 – Discovery **(DONE)** | Action metadata exposure | `actions_list` | Already implemented via `actions.Discover()` and mirrored in CLI. | ✅ complete |
| P1 – Session & Cursor **(DONE)** | Session manager, open/save, cursor inspection | `workspace_open`, `workspace_close`, `workbook_save`, `cursor_get`, `cursor_set`, `info_get` | Landed in `pkg/mcp/server.go`; responses expose sheet lists + cursor metadata. | ✅ complete |
| P2 – Cell & Row Ops **(DONE)** | Cell edits, clears, structural row/column operations | `cell_edit`, `cell_clear`, `row_insert_above`, `row_insert_below`, `row_delete`, `column_insert_left/right`, `column_delete` | Backed by Go actions; responses include placeholder `styleStatus` plus cursor metadata. | ✅ complete |
| P3 – Selection & Export **(DONE)** | Range selection, save-range, clipboard access, width helpers | `selection_set`, `selection_clear`, `selection_export`, `clipboard_get/set`, `column_width` | Selection activation/clearing + clipboard APIs ship with CSV/XLSX `selection_export` destinations, serialized range clipboard payloads, and the new `column_width` tool that surfaces show/set/auto heuristics. | ✅ complete |
| P4 – Styling Semantics *(IN PROGRESS)* | Info schema, style describe/apply | `info_get` (extended schema), `style_describe`, `style_apply` | First cut shipped (style tooling + inline style exports); remaining work: richer schema + semantic annotations to fully satisfy US-07/US-09/US-10. | ⚠️ in progress (doc task `ebenezer-6z0`) |
| P5 – Telemetry & Policy | Structured logs, read-only enforcement, auth hooks | `log_subscribe` (or streaming), policy gate | Aligns with US-04’s telemetry & sandbox requirements. | ⚠️ planned (design task `ebenezer-13x`) |

## Deliverables per Phase
- **Design Artifacts**: Update `SPEC.md §13`, `IMPLEMENTATION_PLAN.md`, and this file whenever scope shifts.
- **Code**: MCP server (`pkg/mcp`), session manager, DTOs matching documented schemas.
- **Tests**: Table-driven unit tests for tool handlers + integration harness once multiple tools exist.
- **Docs**: `docs/mcp-readiness-summary.md` + new `docs/mcp-style-schema.md` when style semantics land.

## Immediate Next Steps
1. **Round out P4 styling schema**: document the enriched `info_get` payload (style aggregates, semantics) in `docs/mcp-style-schema.md` so downstream clients can version against it (`ebenezer-6z0`).
2. **Extend styling semantics**: wire workbook annotations (style meanings, conditional formatting) into `info_get` once the schema lands.
3. **Design P5 Telemetry & Policy**: specify the structured logging stream, session read-only enforcement, and auth hooks so implementation can follow immediately after styling ships (`ebenezer-13x`).
