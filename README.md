# Ebenezer (Go Rewrite)

Ebenezer is a headless-yet-interactive spreadsheet editor with a Vim-inspired workflow. This branch contains the Go rewrite that will eventually replace the original Node/JS prototype located in `js/`.

> **Deprecated JS Version:** The `js/` directory still hosts the legacy Node implementation for reference, but it is no longer maintained. New features and fixes land in the Go codebase (`cmd/`, `internal/`).

## Status: Early Demo
- CLI scaffold available via Cobra (`cmd/ebenezer`).
- Keyboard input loop lifted from the Gordon project (`internal/ui/keyboard`).
- Workbook layer backed by in-memory sample data, CSV loading, and `.xlsx` parsing via Excelize (`internal/workbook`).
- Visual selection mode mirrors Vim: `v` for rectangular ranges, `V` for whole rows, ESC to exit, and the status line reports `VISUAL <range>` while active so clipboard ops preserve formatting.
- Keyboard demo supports arrow-key navigation, `g` to jump to a cell, `s`/`ct`/`rt`/`style` to inspect values and formatting, `:` to enter command mode, and `q` to exit.
- `.xlsx` files reopen at the last active Excel cell/sheet when that metadata exists, and `ps` lets you list/switch sheets without reopening the file.

## Build & Run
```bash
./build.sh              # builds ./ebenezer
./ebenezer data.xlsx    # load a workbook and jump straight into keyboard mode
# or fall back to the sample workbook
./ebenezer              # launches keyboard mode with the built-in sample data
```

While in keyboard mode:
- Arrow keys move the cursor; output shows `→ <address> = "<value>"`.
- `g` prompts for a cell address (e.g., `B12`) and jumps there.
- `/` runs a forward search (prompt pre-fills the last query), `?` searches backward, and `n`/`N` repeat in the same/opposite direction respectively—with wrap-around matching against any cell text.
- Use `:search-case sensitive|insensitive|toggle` to control whether `/` and `?` respect casing (default insensitive).
- `s` prints the current cell via the status helper.
- `c` `t` prints the column header (row 1), `r` `t` prints the row header (column 1).
- `i` edits the current cell (inline prompt). `y` yanks, `x` cuts, `p`/`P` pastes (after/before), `d` `c` clears the cell.
- `v` toggles rectangular visual mode, `V` toggles whole-row selections, and ESC cancels visual mode; while active the status line shows `VISUAL A1:B3 (3x2)` style summaries and `y/x/d/p` operate on the highlighted region (row selections yank/delete entire rows, pasting replaces the selection).
- `Y` yanks the current row, `X` cuts it, `D` deletes it, `O/o` insert blank rows above/below.
- `:style` describes the formatting of the current (or specified) cell.
- `:ps` lists sheets; `:ps Sheet2` switches sheets (only for workbooks opened from disk).
- `:` enters command mode—run commands like `status`, `open js/test.xlsx --sheet Sheet1`, `colheader`, `rowheader`, `sample`, `style B3`, `ps`, `ns ReportCopy Budget`, `style-copy A1:B2`, `style-paste C3:D4`, `save`, `saveas report.xlsx`.
- `q` exits keyboard mode.

### Discovery & MCP
- Run `./ebenezer actions list` to inspect every registered action (name, description, args, idempotency) as pretty JSON. This command reads directly from the action registry (`actions.Discover()`) and is the quickest way to validate new metadata before exposing it to MCP or other transports.
- `./ebenezer mcp serve` launches the stdio-based MCP server. It currently exposes `actions_list`, `workspace_open/close`, `cursor_get/set`, `workbook_save`, `info_get`, `cell_edit`, `cell_clear`, `row_insert_above`, `row_insert_below`, `row_delete`, `column_insert_left/right`, `column_delete`, `selection_set`, `selection_clear`, `selection_export`, `clipboard_get`, `clipboard_set`, plus the new `style_describe` and `style_apply` tools for inspecting and mutating formatting. `selection_export` now mirrors both values and inline style metadata (and still supports CSV/XLSX destinations), while `clipboard_set` handles `kind=range` payloads so agents can round-trip rectangular selections. Styling semantics are documented in `docs/mcp-tools-plan.md` and expand automatically via `info_get`'s enriched payload.

## Roadmap Highlights
See `SPEC.md` and `IMPLEMENTATION_PLAN.md` for the full set of milestones (CLI parity, styling preservation, formula evaluation, MCP server). Near-term goals include:
1. Replacing the CSV stub with full Excelize-based workbook I/O while preserving styles.
2. Building the modal command registry and editing commands (`i`, `yy`, `dd`, etc.).
3. Adding targeted unit tests and automation for the Go codebase.

> **Note on `docs/user-stories-status.xlsx`:** This workbook is treated as a throwaway playground for experimentation and visual demos only. `user-stories.txt` remains the single source of truth for backlog tracking, so never rely on the `.xlsx` file for documentation or planning data.

## Contributing
- Use `go test ./...` before commits; add table-driven tests for new packages.
- Keep docs (`SPEC.md`, `IMPLEMENTATION_PLAN.md`, `docs/`) in sync with code changes.
- The JS prototype remains as a historical reference—avoid modifying it unless fixing migration blockers.
