# Agents Guide for Ebenezer (Go Rewrite)

This document gives CLI coding agents the context they need to work effectively inside the Go rewrite of Ebenezer.

## Project Layout
- `cmd/ebenezer`: Binary entrypoint; keep `main.go` minimal.
- `internal/app`: Session state, cursor handling, workbook lifecycle.
- `internal/workbook`: Workbook I/O (CSV/XLSX via Excelize), style metadata helpers.
- `internal/cmd`: Cobra commands plus keyboard-mode wiring.
- `internal/ui/keyboard`: Vim-like loop lifted from `gordon` and generalized.
- `docs/`, `SPEC.md`, `IMPLEMENTATION_PLAN.md`: Always update as behavior changes.
- `js/`: Deprecated Node prototype; read-only unless fixing migration blockers.

## Build & Test Workflow
- `./build.sh` compiles the CLI (`./ebenezer`).
- `./ebenezer [--sheet SHEET] [FILE]` opens CSV/XLSX files and drops into keyboard mode; omit `FILE` to use the sample data.
- `go test ./...` (set `GOCACHE=/root/ebenezer/.cache go test ./...` if needed) runs existing unit tests; add table-driven tests beside new code.

## Coding Conventions
- Stick to Go 1.22 modules; run `gofmt`/`goimports` on touched files.
- Public APIs live under `internal/...` until we intentionally expose them; avoid circular dependencies.
- Keep command logic in `internal/cmd` thin—delegate behavior to `internal/app`/`internal/workbook` helpers.
- Add short comments only when behavior is non-obvious (e.g., style-hash comparisons, modal key buffering).

## Testing Expectations
- Provide `_test.go` files for new state/workbook helpers and any deterministic command behavior.
- Prefer table-driven tests; use sample CSV/XLSX fixtures under `js/` when practical.
- Run `go test ./...` before every commit; document any manual keyboard-mode smoke tests in PR notes.

## Git & Commit Guidance
- Use Conventional Commit prefixes (`feat:`, `fix:`, `docs:`, `test:`, etc.).
- Keep commits scoped: update code + docs/tests together; don’t mix unrelated changes.
- Never remove or alter user modifications outside your scope; check `git status` before editing shared files.

## Agent Tips
- When you need styles or formula metadata, prefer the Excelize APIs already in `internal/workbook`—don’t reimplement parsing.
- For new keyboard shortcuts, wire them via `internal/cmd/keyboard.go` so they dispatch through Cobra commands; update docs accordingly.
- If you introduce new dependencies, update `go.mod`/`go.sum` and note the reason in the PR summary.
- Always keep `README.md`, `docs/demo.md`, and the implementation plan in sync with user-facing changes.
