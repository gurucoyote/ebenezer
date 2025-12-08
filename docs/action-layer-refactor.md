# Action Layer Refactor Plan (v2)

This revision focuses on a practical, single-user refactor that extracts command logic into reusable Go actions while keeping future multi-session/telemetry goals achievable. Scope explicitly excludes telemetry plumbing and multi-session orchestration for now; the interface leaves room for them later.

## Desired End State
1. **Action Core**: Every CLI behavior lives in a testable function that accepts a context (`State + IO + optional clock/logger`) and returns a structured `Result` (data + optional message).
2. **Adapters Only**: Cobra commands, keyboard bindings, and future frontends are thin shells that translate user input into action calls and render the result.
3. **Test Coverage**: Each action has table-driven tests; CLI/keyboard smoke tests verify adapters pass through arguments untouched.

## Scope Exclusions (for now)
- No telemetry schema or logging bus (keep placeholders in the action context for future use).
- Single session only; state injection ensures later multi-session work doesn’t require another rewrite.

## Phase Breakdown
### Phase 0 – Inventory & Test Harness (2 days)
- List every command + keyboard binding, the arguments they accept, and the state methods they touch.
- Add or extend minimal fixtures/tests for current behavior (e.g., `internal/app/state_test.go`, CLI smoke test) to catch regressions before refactor.
- Deliverable: `docs/action-layer-inventory.md` with table of commands/bindings.

### Phase 1 – Action Skeleton & Pilot (3 days)
- Create `internal/actions` package with:
  ```go
  type Context struct { State *app.State; Writer io.Writer }
  type Result struct { Message string; Data any }
  type Action interface { Name() string; Exec(ctx Context, args []string) (Result, error) }
  ```
- Implement pilot actions (`Move`, `Goto`, `Status`) and migrate their Cobra handlers to call these actions.
- Wire keyboard bindings for arrow keys/`g` to invoke pilot actions directly. `:` command mode remains on Cobra for Vim-like workflows, but those Cobra handlers must still call actions under the hood.
- Deliverable: passing tests for pilot actions + unchanged CLI UX.

### Phase 2 – Registry & Navigation/Edit Commands (5 days)
- Build a simple registry (`internal/actions/registry.go`) mapping names → action metadata (usage/help).
- Migrate navigation and editing commands (`move`, `goto`, `row`, `col`, insert/delete/yank/paste) to the action layer using the registry.
- Keyboard bindings switch to lookups in the registry for normal-mode shortcuts, while `:` commands keep using Cobra but continue to call actions so behavior stays unified.
- Deliverable: CLI + keyboard parity for navigation/editing with zero logic left in Cobra.

### Phase 3 – Formatting, Search, Save/Load (5 days)
- Convert style commands (`style`, `style-copy`, `style-paste`), search (`/`, `?`, `search-case`, `fi`), and save/load operations into actions.
- Introduce helper structs for repeated patterns (e.g., style descriptions, search results) to keep `Result.Data` consistent.
- Update docs/tests accordingly.

### Phase 4 – Range/Selection & Info Commands (4 days)
- Migrate selection-aware commands (`visual`, pending US-05 features) and the forthcoming `info` command (US-07).
- Ensure action outputs carry enough metadata for future range exports/info enhancements.

### Phase 5 – Final Wiring & Docs (3 days)
- Remove any remaining direct `appState` references from Cobra.
- Update `SPEC.md §5`, `IMPLEMENTATION_PLAN.md`, `AGENTS.md`, and add an “Adding Actions” guide.
- Run full `go test ./...` and manual keyboard smoke test, documenting results.

## Implementation Details
- **State Injection**: `main.go`/`cmd/root.go` creates a single `*app.State` and passes it to adapters. For now, this is a singleton, but all actions expect it via context so multi-session later only changes the factory.
- **Result Rendering**: CLI adapter prints `Result.Message` (if non-empty) and pretty-prints `Result.Data` when needed. Keyboard mode only uses the message. Future MCP can serialize `Result.Data` as JSON.
- **Error Handling**: Actions return errors with context; adapters surface them verbatim. Keyboard loop already logs errors via `InfoWriter` and will continue to do so.
- **Testing Strategy**: For each migrated command, add action-level tests using the sample workbook (`workbook.SampleWorkbook()`) or fixtures under `js/`. CLI/keyboard smoke tests ensure bindings still issue the right actions.

This plan keeps the refactor tractable while delivering immediate benefits (testable logic, reusable commands) and leaving hooks for more advanced telemetry or multi-session features later.
