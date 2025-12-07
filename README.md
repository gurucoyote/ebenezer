# Ebenezer (Go Rewrite)

Ebenezer is a headless-yet-interactive spreadsheet editor with a Vim-inspired workflow. This branch contains the Go rewrite that will eventually replace the original Node/JS prototype located in `js/`.

> **Deprecated JS Version:** The `js/` directory still hosts the legacy Node implementation for reference, but it is no longer maintained. New features and fixes land in the Go codebase (`cmd/`, `internal/`).

## Status: Early Demo
- CLI scaffold available via Cobra (`cmd/ebenezer`).
- Keyboard input loop lifted from the Gordon project (`internal/ui/keyboard`).
- Workbook layer backed by in-memory sample data, CSV loading, and `.xlsx` parsing via Excelize (`internal/workbook`).
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
- `s` prints the current cell via the status helper.
- `c` `t` prints the column header (row 1), `r` `t` prints the row header (column 1).
- `i` edits the current cell (inline prompt). `y` yanks, `x` cuts, `p`/`P` pastes (after/before), `d` `c` clears the cell.
- `Y` yanks the current row, `X` cuts it, `D` deletes it, `O/o` insert blank rows above/below.
- `:style` describes the formatting of the current (or specified) cell.
- `:ps` lists sheets; `:ps Sheet2` switches sheets (only for workbooks opened from disk).
- `:` enters command mode—run commands like `status`, `open js/test.xlsx --sheet Sheet1`, `colheader`, `rowheader`, `sample`, `style B3`, `ps`.
- `q` exits keyboard mode.

## Roadmap Highlights
See `SPEC.md` and `IMPLEMENTATION_PLAN.md` for the full set of milestones (CLI parity, styling preservation, formula evaluation, MCP server). Near-term goals include:
1. Replacing the CSV stub with full Excelize-based workbook I/O while preserving styles.
2. Building the modal command registry and editing commands (`i`, `yy`, `dd`, etc.).
3. Adding targeted unit tests and automation for the Go codebase.

## Contributing
- Use `go test ./...` before commits; add table-driven tests for new packages.
- Keep docs (`SPEC.md`, `IMPLEMENTATION_PLAN.md`, `docs/`) in sync with code changes.
- The JS prototype remains as a historical reference—avoid modifying it unless fixing migration blockers.
