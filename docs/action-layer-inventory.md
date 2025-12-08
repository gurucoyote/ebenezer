# Action Layer Inventory (Phase 0)

This inventory captures the current CLI commands and keyboard bindings so we know exactly what must migrate into the action layer.

## Baseline Test Status (2025-12-07)
- Command: `go test ./...`
- Result: **FAIL** – `testing/internal/testdeps: package testmain: cannot find package` while building `internal/app` and `internal/workbook`.
- Cause: the default Go build cache at `/root/.cache/go-build` is read-only in this environment, which makes the toolchain report misleading “cannot find package” errors.
- Workaround: run tests with a writable cache, e.g. `GOCACHE=/root/ebenezer/.cache go test ./...` (recommended in AGENTS.md).
- Verified: `GOCACHE=/root/ebenezer/.cache go test ./...` now passes for all packages.

## CLI Commands
| Command | File | Description (from `Short`) | Notes / Key Bindings |
|---------|------|----------------------------|----------------------|
| `ebenezer` (root) | `internal/cmd/root.go` | Ebenezer spreadsheet CLI entrypoint | Loads file or sample, then enters keyboard mode |
| `open` | `internal/cmd/open.go` | Open a workbook (.csv or .xlsx) | Accepts optional `--sheet` |
| `keyboard` | `internal/cmd/keyboard.go` | Enter vim-like keyboard control mode | Main TUI loop |
| `status` (`:status` / `s`) | `internal/cmd/status.go` | Print the current cursor/value | Bound to `s` key in keyboard mode |
| `move` | `internal/cmd/move.go` | Move the demo cursor | Arrow keys map to `move` |
| `goto` (`:goto` / `g`) | `internal/cmd/goto.go` | Move the cursor to a cell address | `g` prompt launches command |
| `colheader` (`ct`) | `internal/cmd/goto.go` | Print current column header | Keyboard `c` then `t` |
| `rowheader` (`rt`) | `internal/cmd/goto.go` | Print current row header | Keyboard `r` then `t` |
| `edit` (`i`) | `internal/cmd/edit.go` | Set the current cell's value | Triggered via `i` |
| `clear` | `internal/cmd/edit.go` | Clear the current cell | Keyboard `d` then `c` (delete cell shortcut) |
| `yank` (`y`) | `internal/cmd/edit.go` | Copy the current cell | Keyboard `y` |
| `cut` (`x`) | `internal/cmd/edit.go` | Cut the current cell | Keyboard `x` |
| `paste` (`p`/`P`) | `internal/cmd/edit.go` | Paste clipboard | `p` (after) / `P` (before) |
| `row` (group) | `internal/cmd/row.go` | Row operations wrapper | unused directly |
| `row yank` (`Y`) | `internal/cmd/row.go` | Yank current row | Keyboard `Y` |
| `row cut` (`X`) | `internal/cmd/row.go` | Cut current row | Keyboard `X` |
| `row delete` (`D`) | `internal/cmd/row.go` | Delete current row | Keyboard `D` |
| `row insert-above` (`O`) | `internal/cmd/row.go` | Insert blank row above | Keyboard `O` |
| `row insert-below` (`o`) | `internal/cmd/row.go` | Insert blank row below | Keyboard `o` |
| `sheet select` (`:ps`) | `internal/cmd/sheet.go` | List/switch sheets | `:ps` / command palette |
| `sheet new` (`:ns`) | `internal/cmd/sheet.go` | Create new sheet / clone | `:ns` |
| `sample` (`:sample`) | `internal/cmd/sample.go` | Reload built-in sample workbook | CLI only |
| `style` (`:style`) | `internal/cmd/style.go` | Describe formatting of a cell | CLI command |
| `style-copy` (`:style-copy`) | `internal/cmd/style.go` | Copy formatting | CLI command |
| `style-paste` (`:style-paste`) | `internal/cmd/style.go` | Paste formatting | CLI command |
| `search` (`/` `?`) | `internal/cmd/search.go` | Search current sheet | `/` or `?` prompts in keyboard mode |
| `search-next` (`n`) | `internal/cmd/search.go` | Repeat search forward | Keyboard `n` |
| `search-prev` (`N`) | `internal/cmd/search.go` | Repeat search backward | Keyboard `N` |
| `search-case` (`:search-case`) | `internal/cmd/searchcase.go` | View/change case-sensitivity | CLI command |
| `save` (`:save`) | `internal/cmd/save.go` | Save current workbook | CLI command |
| `save!` | `internal/cmd/save.go` | Force save | CLI command |
| `saveas` (`:saveas`) | `internal/cmd/save.go` | Save to new file | CLI command |
| `write` (`:w`) | `internal/cmd/save.go` | Vim-style save alias | CLI command |
| `write!` (`:w!`) | `internal/cmd/save.go` | Force save alias | CLI command |
| `echo` (`:echo`) | `internal/cmd/echo.go` | Echo text (testing placeholder) | CLI only |

## Keyboard Bindings (Normal Mode)
Source: `internal/cmd/keyboard.go` (Bindings). Keys not listed fall through to command-mode `:` or no-op.

| Key / Sequence | Action Triggered |
|----------------|------------------|
| Arrow Left/Right/Up/Down | `move left/right/up/down` |
| `Esc` | Clears current selection |
| `i` | `edit` (insert) |
| `/` | `search` forward prompt |
| `?` | `search` backward prompt |
| `n` | `search-next` |
| `N` | `search-prev` |
| `v` | Toggle rectangular visual selection |
| `V` | Toggle row visual selection |
| `s` | `status` |
| `g` | `goto` prompt |
| `c` then `t` | `colheader` |
| `r` then `t` | `rowheader` |
| `y` | `yank` cell |
| `Y` | `row yank` |
| `x` | `cut` cell |
| `X` | `row cut` |
| `p` | `paste` after |
| `P` | `paste --before` |
| `d` then `c` | Clear/delete cell |
| `D` | `row delete` |
| `O` | `row insert-above` |
| `o` | `row insert-below` |
| `:` | Enter command mode (Cobra command input) |
| `q`/`Q` | Quit keyboard mode |

This table will guide the action-layer migration and ensure no binding is lost during refactor.
