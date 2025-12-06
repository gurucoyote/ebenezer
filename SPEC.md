# Ebenezer Go CLI & MCP Server Spec

## 1. Vision & Scope
- **Goal**: Rebuild Ebenezer as a fast, portable Go CLI that edits `.xlsx`/`.csv` workbooks in a headless yet interactive way, preserving the Vim-style workflow from the Node prototype while improving ergonomics, installability, and reliability.
- **Secondary Goal**: Provide a machine-consumable interface (MCP server) that exposes spreadsheet editing primitives to CLI coding agents (Codex and others) without relying on terminal key bindings.
- **Out of Scope v1**: Full Excel feature parity, graphical UI, concurrent multi-user editing, online storage integrations.

## 2. Target Personas & Workflows
1. **Terminal Power User** editing structured data quickly (e.g., CSV exports) without launching Excel/LibreOffice.
2. **Automation Engineer / Agent** that needs deterministic editing primitives to script spreadsheet tweaks in a CI or CLI agent environment.

Key workflows:
- Open an existing workbook or create a new one when no filename is provided.
- Navigate with modal key bindings, inspect formula results, and edit cell contents quickly.
- Manipulate rows/columns (insert, delete, yank/cut/paste) and switch between sheets.
- Save to new filenames, preserving prior filenames as history suggestions.
- Search within a column and jump to matches.
- Run ad-hoc scripts inside an embedded REPL (optional stretch goal for parity with Node version).
- Future: expose these operations via MCP for remote invocation.

## 3. Functional Requirements (CLI)
### 3.1 Launch & Session State
- Command: `ebenezer-go [--sheet SHEETNAME] [FILE]`, exposed via Cobra root command so subcommands (e.g., `ebenezer-go config`, `ebenezer-go mcp serve`) can reuse shared flags.
- If `FILE` omitted, start with in-memory workbook named `untitled.xlsx` and prompt to save on exit.
- Maintain session state struct (`AppState`) containing workbook handle, active sheet, cursor position (`row`, `col`), yank buffer, filename history, config, and abort controller.
- Detect file format by extension; allow overriding delimiter for CSV via `--delimiter` (default `;`).

### 3.2 Modes & Key Handling
- Default to **Normal mode** in raw terminal input. Buffer sequences within configurable `keyWait` (default 250ms) to match multi-key commands.
- **Insert mode** temporarily disables key-sequence capture to allow freeform text entry, with ESC to abort edits.
- Maintain a command registry mapping sequences to actions and help text; pressing `h`/`?` prints grouped help.

### 3.3 Navigation Commands
- `down|enter|return`, `up`, `left`, `right` move the cursor one cell, clamped at row/column ≥ 1.
- `g` enters insert prompt pre-filled with current cell address; accepts e.g. `B12`.
- `ct`/`rt` read header cell (row 1, current column) or first column of current row.

### 3.4 Edit Commands
- `i`: edit current cell; if input begins with `=`, treat as formula, otherwise literal string/number.
- `y`/`x`: yank/cut cell value into clipboard.
- `Y|yy`/`X|xx`: yank/cut row; `D|dd`: delete row (shift up).
- `yc`/`xc`/`dc`: column yank/cut/delete.
- `O`/`o`: insert blank row above/below; insertion updates cursor to the new row.
- `P`/`p`: paste clipboard before/after current row/column/location, based on clipboard type (cell/row/column).
- `ns`: prompt for new sheet name, validate uniqueness, create sheet, and switch focus.
- `ps`: prompt with sheet list, switch when valid choice made.
- `fi`: column search: prompt for regex/text, list matches, allow arrow navigation and jump.
- `wb`: save workbook; prompt with filename history, support `.xlsx` and `.csv` (sheet-scoped writer).
- `:` (stretch goal) open Go REPL or Lua-like scripting environment (optional for parity).

### 3.5 Formula Evaluation
- Display cell value or resolved formula result when reporting the current cell.
- Support evaluation of Excel-compatible formulas (SUM, AVERAGE, references, ranges) within a single sheet; circular references yield warning and raw formula.
- Provide fallback when a formula cannot be evaluated (log error, display `#ERR`).

### 3.6 Column Search & History
- Maintain last-search string and revisit via `fi` prompt history.
- Provide user-friendly navigation: after populating results, the user selects via arrow keys or enters a coordinate.

- Maintain MRU filename list (max configurable, default 5) in memory; persist via Viper-managed config file (default `$XDG_CONFIG_HOME/ebenezer/config.yaml`) for future sessions.

## 4. Non-Functional Requirements
- **Portability**: builds for macOS, Linux, Windows without CGO.
- **Binary size**: target < 25MB.
- **Performance**: open 50k-row CSV under 3s on modern hardware.
- **Reliability**: ensure yank/cut/paste operations fail fast if workbook is in read-only state.
- **Accessibility**: all feedback via stdout/stderr; avoid reliance on color.

## 5. Architecture Overview
CLI orchestration (Cobra) and configuration (Viper) should be layered so that the same command primitives can later be exposed through the MCP server module without duplicating logic.
```
cmd/ebenezer/main.go        // CLI entrypoint, flag parsing
internal/app/state.go       // AppState struct, lifecycle hooks
internal/workbook/io.go     // Load/save logic using Excelize & encoding/csv
internal/formula/engine.go  // Formula evaluation abstraction
internal/ui/terminal.go     // Terminal raw-mode, key buffer, prompts
internal/commands/*.go      // Individual command handlers registered in map
internal/search/search.go   // Column search utilities
pkg/mcp/server.go           // MCP server implementation (Phase 2)
```
- Use dependency injection so both CLI and MCP server share workbook + command primitives.
- Introduce `Action` interface with `Name() string`, `Description() string`, `Exec(*AppState) error` for reuse.

## 6. External Dependencies (tentative)
- [`github.com/spf13/cobra`](https://github.com/spf13/cobra) to declare the CLI surface (root command plus future MCP/utility subcommands) and manage flag parsing consistently.
- [`github.com/spf13/viper`](https://github.com/spf13/viper) for config discovery, overrides (env vars, flags), and persistence of MRU filenames/delimiter settings.
- [`github.com/AlecAivazis/survey/v2`](https://github.com/AlecAivazis/survey) for all interactive selections/prompts (sheet picker, filename history chooser, column search results) to give consistent UX across terminals.
- [`github.com/xuri/excelize/v2`](https://github.com/xuri/excelize) for workbook I/O and formula metadata.
- [`github.com/rivo/tview`](optional) if richer TUI needed; initial implementation can use `golang.org/x/term` and custom rendering.
- [`github.com/antonmedv/expr`](optional) or a lightweight expression parser if Excelize formula evaluation proves insufficient.

## 7. Data Structures
- `type Clipboard struct { Kind ClipboardKind; Cells [][]CellValue }` where `ClipboardKind ∈ {Cell, Row, Column}`.
- `type CellValue struct { Display string; Formula string; Type ValueType }` to round-trip formulas.
- `type AppState struct { Workbook *excelize.File; ActiveSheet string; Cursor CellRef; Clipboard Clipboard; Filename string; FilenameHistory []string; KeyWait time.Duration; Config Config; }`.
- `type Command struct { Sequence string; Help string; Handler func(*AppState, *UI) error }` stored in a slice for help rendering and in a map for lookup.

## 8. Terminal Interaction & Prompts
- Implement `UI` abstraction with methods: `ReadLine(prompt string, default string, history []string) (string, error)`, `PrintHelp(map[string][]string)`, `ReportCell(CellRef, DisplayValue)`, `DisplayMatches([]Match)`.
- Back all user selections (sheet picker, filename chooser, column search results, config toggles) with `survey` widgets (select, multiselect, confirm) for consistent keyboard support.
- ESC from insert mode cancels active prompt; deliver consistent status line updates (e.g., `A1: 42`).
- Provide status/log output channel so MCP server can consume structured logs instead of raw strings.

## 9. File Handling & Validation
- On load, validate extension and give explicit errors for unsupported types.
- On save, prompt before overwriting existing file (unless `--force`).
- For CSV, support configurable delimiter/quote strategy via CLI flags or config file.
- Implement autosave toggle (default off) by writing to temp file then atomic rename.

## 10. Error Handling & Logging
- Use structured logging (zerolog or stdlib log/slog) with log levels; default to minimal output in CLI but verbose when `--debug` enabled.
- Provide user-facing errors for invalid commands, formulas, or sheet names; never panic on malformed input.

## 11. Testing Strategy
- Unit tests for: column/row index conversion, clipboard operations, save/load, command handlers.
- Integration tests using golden files for CSV/XLSX transformations.
- Terminal interaction tests via pseudo-terminal (pty) harness for key-sequence buffering.
- For MCP server, add protocol conformance tests with mock clients.

## 12. Distribution & Tooling
- Provide `make build` (cross-compiles), `make lint`, `make test`.
- Optional `brew`/`scoop` formulas once stable.

## 13. MCP Server Extension (Phase 2)
### 13.1 Goals
- Allow CLI agents to open workbooks, query metadata, and perform mutations through structured messages instead of terminal keypresses.
- Keep feature parity with human CLI operations but expose them as idempotent RPC-style calls.

### 13.2 Protocol Surface (initial)
| Capability | Request Fields | Response |
|------------|----------------|----------|
| `workspace/open` | `path`, `format`, `sheet` | session token, sheet list |
| `cursor/get` | session token | `row`, `col`, `address`, `value`, `formula`, `computedValue` |
| `cursor/set` | session token, address | updated cursor info |
| `cell/edit` | session, address, `value` or `formula` | confirmation |
| `row/insert`, `row/delete`, `column/insert`, `column/delete` | session, position | updated sheet snapshot (diff) |
| `clipboard/paste` | session, destination, clipboard payload | confirmation |
| `sheet/list`, `sheet/select`, `sheet/create` | session, parameters | updated session |
| `workbook/save` | session, `path`, `format` | success + bytes written |

- Sessions map 1:1 with open workbooks; when invoked from CLI, commands can call MCP handlers internally to ensure single logic path.
- Return structured errors with machine-readable codes (e.g., `ErrSheetExists`, `ErrInvalidAddress`).

### 13.3 Transport & Packaging
- Implement MCP server using stdio transport (same as Codex) for compatibility; optionally expose Unix socket/TCP later.
- Provide feature discovery metadata so agents can list supported commands.
- Document rate limits and concurrency model (single-threaded per session, queue requests).

### 13.4 Security & Permissions
- Deny file operations outside working directory by default; allow opt-in via config.
- Expose read-only mode for agents that must inspect data without mutating.

### 13.5 Telemetry & Observability
- Emit structured events for each MCP call (start, success/failure, duration) to aid debugging when integrated with agents.

## 14. Roadmap
1. **Milestone 1**: Implement CLI core (loading/saving, navigation, editing, yanking/pasting, sheet mgmt, column search, help output).
2. **Milestone 2**: Add formula evaluation abstraction, config persistence, autosave.
3. **Milestone 3**: Harden CSV options, add REPL/script hooks, expand tests, release v0.1 binaries.
4. **Milestone 4**: Build MCP server module, share command handlers, document protocol, integrate with Codex.
5. **Milestone 5**: Performance tuning, package distribution, community feedback loop.

---
This spec provides the baseline requirements and architecture notes needed to start the Go rewrite while keeping future MCP integration in mind. Update it as design decisions are finalized.
