# Keyboard Loop Module

This package lifts the reusable pieces from `/root/gordon/cmd/keyboard.go` so Ebenezer can reuse the same vim-like keyboard workflow.

## Highlights
- Wraps `github.com/eiannone/keyboard` for raw key capture.
- Provides a `Loop` type with customizable rune/key bindings plus default `:` command-mode and `q/Q` quit keys.
- Command-mode integrates with any dispatcher that satisfies `CommandExecutor` (Cobra root command, future MCP bridge, etc.).

## Next Steps
1. Wire `Loop` into the future `internal/ui/terminal` package once command registry scaffolding exists.
2. Register rune/key bindings that map to spreadsheet actions (navigation, edit, yank/paste) via the shared `Action` callbacks.
3. Extend bindings to support multi-key sequences (e.g. `dd`, `yy`) by layering a small buffer on top of the provided `Context`.

See `loop.go` for inline docs referencing SPEC §3.2.
