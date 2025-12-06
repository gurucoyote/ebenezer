# Ebenezer (Go Rewrite)

Ebenezer is a headless-yet-interactive spreadsheet editor with a Vim-inspired workflow. This branch contains the Go rewrite that will eventually replace the original Node/JS prototype located in `js/`.

> **Deprecated JS Version:** The `js/` directory still hosts the legacy Node implementation for reference, but it is no longer maintained. New features and fixes land in the Go codebase (`cmd/`, `internal/`).

## Status: Early Demo
- CLI scaffold available via Cobra (`cmd/ebenezer`).
- Keyboard input loop lifted from the Gordon project (`internal/ui/keyboard`).
- Demo-friendly workbook layer with in-memory sample data plus CSV loading (`internal/workbook`).
- Keyboard demo supports arrow-key navigation, `s` to print cell status, `:` to enter command mode, and `q` to exit.

## Build & Run
```bash
./build.sh              # builds ./ebenezer
./ebenezer keyboard     # launch demo keyboard loop
```

While in keyboard mode:
- Arrow keys move the cursor; output shows `→ <address> = "<value>"`.
- `s` prints the current cell via the status helper.
- `:` enters command mode—run commands like `status`, `open js/test.csv`, `sample`.
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
