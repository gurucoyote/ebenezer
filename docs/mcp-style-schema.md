# MCP Style Schema

This document spells out the enriched MCP payloads that land in Phase P4 so style-aware clients can version against the full `info_get` result, understand `styleStatus` hints, and interpret range/selection reports before downstream MCP work finishes.

## `info_get` contract

`info_get` is already part of the MCP surface (see `pkg/mcp/server.go`), and it now returns the `actions.WorkbookInfo` object defined in `internal/actions/info.go`. The call accepts `session_id` plus an optional `details` flag; when `details` is `true` the same payload includes the recorded file metadata (`FileSize`, `ModTime`) and the count of styled cells, but every field below is present even when the flag is omitted.

```json
{
  "sessionId": "session-123",
  "info": {
    "path": "/workspace/data/model.xlsx",
    "activeSheet": "Budget",
    "rows": 200,
    "cols": 12,
    "sheetNames": ["Budget", "Notes"],
    "activeCell": "B14",
    "fileSize": 42816,
    "modTime": "2025-12-23T08:09:17Z",
    "styles": {
      "styledCells": 41,
      "fillColors": {"#FFEB9C": 8},
      "fontColors": {"#C00000": 5},
      "numberFormats": {"$#,##0.00": 12},
      "borderUsage": {"bottom": 10},
      "boldCells": 9,
      "italicCells": 2,
      "underlineCells": 1
    }
  }
}
```

### Workbook info fields

| Field | Meaning |
| --- | --- |
| `path` | Absolute or relative path of the workbook currently open in the session. | 
| `activeSheet` | Name of the sheet containing the cursor. |
| `rows` / `cols` | Maximum coordinates tracked for the active sheet (`workbook.MaxCoords`). |
| `sheetNames` | List of workbook sheets, ordered as the user sees them. |
| `activeCell` | Excel-style address of the active cell if known. |
| `fileSize` | Size in bytes (zero when the underlying file is a new, unsaved workbook). |
| `modTime` | Last-modified timestamp (RFC 3339). |
| `styles` | `StyleSummary` aggregate (see below). |

## Style aggregates (`StyleSummary`)

The `styles` object is the `actions.StyleSummary` struct, which tallies the formatting metadata collected from `internal/workbook.CellStyle`. Every styled cell increments `styledCells`, while the maps count how often particular values occur.

| Property | Description |
| --- | --- |
| `styledCells` | Total cells that have at least one fill/font/number/border/flag attribute. |
| `fillColors` | Map from normalized hex color (`#RRGGBB` or `#AARRGGBB` trimmed to the RGB portion) to the number of cells using that fill. |
| `fontColors` | Map of font color hexes. |
| `numberFormats` | Map of Excel number format codes (`$#,##0`, `0%`, `mm-dd`, etc.). |
| `borderUsage` | Map keyed by edge (`top`, `bottom`, `left`, `right`, `vertical`, `horizontal`) counting how often a border exists on that edge across all styled cells. |
| `boldCells`, `italicCells`, `underlineCells` | Total cells that specify those text emphases. |

The summary is recomputed when the workbook is opened and when styles change via `style_apply`. The `style_describe`/`style_apply` tools (see `pkg/mcp/server.go`) expose the same `workbook.CellStyle` shape, so clients can map the legend entries above back to individual cells when needed.

## `styleStatus` semantics

Several MCP write operations (`cell_edit`, `row_insert_*`, `column_insert_*`, `style_apply`, etc.) include a `styleStatus` string as a lightweight assurance that formatting was preserved, inherited, or blocked. Presently the example values are:

- `retained`: The touched cell/row/column kept its pre-existing style or the operation applied without needing a new style.
- `inherited`: The workbook had to reuse a default style (e.g., the row/column took the formatting of a neighboring cell when new cells were inserted).
- `error`: The requested style operation conflicted with workbook invariants (e.g., styles were malformed or the session is read-only).

Future status codes can be added; treat the field as advisory metadata rather than a control signal.

## Range selection & export reports

The MCP selection tools report user-friendly summaries so remote clients can mirror the keyboard-mode UX.

### `selection_set`

`selection_set` returns:

- `range`: The normalized input range (e.g., `A1:C3`).
- `summary`: A compact representation from `internal/app.State.SelectionSummary()` (row range or `A1:C3 (3x3)`).
- `active`: `true` when visual range/row mode is enabled (set to `false` by `selection_clear`).

### `selection_export`

Exports now emit both values and inline style metadata so downstream tooling can round-trip formatting.

```json
{
  "sessionId": "session-123",
  "range": "A1:B2",
  "values": [["42", "=SUM(...)"], ["foo", ""], ...],
  "rows": 2,
  "cols": 2,
  "styles": [
    {
      "row": 1,
      "col": 1,
      "address": "A1",
      "rowOffset": 0,
      "colOffset": 0,
      "style": {"bold": true, "fillColor": "#FFDAB9"},
      "value": "42"
    }
  ],
  "savedPath": "/tmp/export.csv",
  "format": "csv",
  "bytes": 328
}
```

Each entry in `styles` is a `rangeStyleEntry` (see `pkg/mcp/server.go`):

- `row` / `col` / `address`: Absolute coordinates of the styled cell.
- `rowOffset` / `colOffset`: Offset relative to the exported range (useful for re-mapping into existing sheets).
- `value`: Optional formatted value for the cell (mirrors the exported table for that cell). If a cell has a style but no value, this field is omitted.
- `style`: The full `workbook.CellStyle` payload (`fillColor`, `fontColor`, `numberFormat`, `horizontalAlign`, `verticalAlign`, `borders`, `bold`, `italic`, `underline`).

Clients should preserve these styles when re-importing or applying them via `style_apply` so formatting stays in sync with the remote workbook.

## Style legend & next steps

The MCP schema currently documents the formatting attributes that `workbook.CellStyle` exposes (see `internal/workbook/styleinfo.go`). Additional semantics (conditional formatting, named styles, policy flags) will be layered into `info_get` once the relevant backend work finishes. Update this doc when the `styles` object gains new keys or when `selection_export.styles` starts including policy annotations.
