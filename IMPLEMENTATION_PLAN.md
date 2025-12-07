# Ebenezer Go Implementation Plan

This plan translates the high-level requirements captured in `SPEC.md` into scoped work packages, ordered milestones, and concrete acceptance criteria. Section references (e.g., §3.4) point back to the source spec for traceability.

## 0. Foundational Assumptions
- Go 1.22+ toolchain with modules; builds flow through `build.sh` (Makefile optional per requirements).
- Workspace-write sandbox with no CGO; external deps limited to those listed in §6 unless a change request is raised.
- Unit/integration testing required for every non-trivial package prior to merging.
- CLI UX must remain parity-compatible with the Node prototype (see `js/`) where behavior is defined there but not explicitly restated in the spec.

## 1. Workstreams & Owners
| Workstream | Scope | Key Packages | Primary Dependencies |
|------------|-------|--------------|----------------------|
| CLI & Cobra Surface | Binary entrypoint, flag plumbing, subcommand scaffolding (§3.1) | `cmd/ebenezer`, `internal/app` | App state, config |
| Terminal UI & Keymaps | Raw terminal handling, mode transitions, prompts (§3.2, §8) | `internal/ui`, `internal/commands` | Cobra root, survey |
| Workbook I/O | Load/save XLSX/CSV, clipboard ops (§3.1, §3.4, §9) | `internal/workbook`, `internal/clipboard` | excelize |
| Formatting & Styling Preservation | Style snapshot/reapply, regression protection (§3.6) | `internal/workbook/styles.go`, `internal/commands` | excelize |
| Command Handlers | Navigation, edit, sheet mgmt, history (§3.3–3.6) | `internal/commands/*` | App state, UI, workbook |
| Formula Engine | Evaluation, error reporting (§3.5) | `internal/formula` | workbook |
| Search & History | Column search, MRU filenames (§3.6) | `internal/search`, `internal/history` | UI, workbook |
| Config & Persistence | Viper setup, MRU persistence, autosave (§3.6, §9) | `internal/config` | Cobra |
| MCP Server | Phase 2 transport & handlers (§13) | `pkg/mcp` | shared action layer |

## 2. Milestone Breakdown

### Milestone 0 – Project Bootstrap (Week 1)
**Goals**: Scaffolding, tooling, and automation baseline.
1. [x] Initialize Go module (`go.mod`) with cobra/viper/excelize dependencies (current scope: cobra, keyboard, readline).
2. [x] Provide a repeatable build entry point (`build.sh`); Makefile intentionally deferred per updated requirements.
3. [x] Set up local `go test ./...` workflow (CI optional in this phase).
4. [x] Establish `internal/app`, `internal/ui`, `internal/workbook` package skeletons with docstrings referencing SPEC sections. *(Status: skeletons exist, docstrings TODO.)*
5. [x] Add basic unit test harness (table-driven tests) covering cursor math (`internal/app/state_test.go`) and workbook helpers (`internal/workbook/workbook_test.go`).

**Acceptance**: ✅ Binary builds via `build.sh`, smoke tests executed with `go test ./...`, and unit tests exist for foundational packages.

### Milestone 1 – Core CLI & Editing Loop (Weeks 2–4)
Deliverable: Minimum usable CLI covering §3.1–3.4, §8 basics.

Tasks:
1. [x] **App State & Lifecycle** (`internal/app/state.go`): implement state + sample workbook loader. *(Filename history still future work.)*
2. [~] **Workbook I/O** (`internal/workbook/io.go`): CSV + Excelize-based `.xlsx` loading done; saving, delimiter overrides, and style snapshots still pending.
3. [ ] **Terminal Raw Mode** (`internal/ui/terminal.go`): not started (current demo uses lifted keyboard loop only).
4. [ ] **Command Registry** (`internal/commands/registry.go`): TBD — current commands directly call state.
5. [x] **Navigation Commands**: arrow key handlers, goto, row/column headers wired via Cobra + keyboard shortcuts.
6. [~] **Editing Commands**: cell + row edit/yank/cut/paste/clear implemented; column operations and save/write flows pending.
7. [ ] **Style Snapshot Helpers**: not started.
8. [x] **Status Reporting** (`internal/ui/status.go`): prints cursor/value after commands.
9. [x] **Visual Selection Mode**: rectangular (`v`) and row (`V`) selections update AppState, drive clipboard-aware `y/x/p/d`, and surface status-line summaries with ESC to exit.
10. [x] **Search UX Settings**: `/` `?` `n` `N` implemented with configurable case-sensitivity (default insensitive) via `search-case` command.

Testing:
- [ ] Unit tests (cursor math, clipboard, etc.).
- [ ] CSV/XLSX golden tests.
- [ ] PTY smoke tests.

Exit Criteria (current status):
- Sample CSV load + navigation works via `open`/keyboard ✅.
- Editing/styling guarantees ❌ (future).

### Milestone 2 – Search, History, and Formula Evaluation (Weeks 5–6)
Scope: §3.5–3.7, §7 (history), §8 enhancements.

Tasks:
1. **Formula Engine** (`internal/formula/engine.go`): integrate excelize evaluation, fallback to `#ERR`, warn on circular refs. If Excelize’s runtime proves too heavy, ship behind an opt-in flag without blocking the milestone.
2. **Column Search Module** (`internal/search/search.go`): regex search, match list, navigation, and Vim-style `/` `?` `n` `N` bindings that wrap around the sheet and remember the last query.
3. **Prompt/Survey Integration**: adopt `survey` widgets for sheet picker, filename history, search result selection.
4. **Filename History Persistence**: `internal/history/history.go` using Viper config at `$XDG_CONFIG_HOME`.
5. **Autosave Toggle**: config flag, temp file write + rename per §9.

Testing:
- Unit tests for formula success/failure paths.
- Search tests using synthetic sheet data.
- Config persistence test writing to temp directories.

Exit Criteria:
- `fi` command lists matches, supports arrow navigation.
- MRU filenames persist across runs.
- Formula cells display computed values with raw fallback.

### Milestone 3 – Reliability & Stretch UX (Weeks 7–8)
Scope: §3.4 stretch features, §4 non-functional, §10 logging.

Tasks:
1. **REPL Command (`:`)**: embed Go eval or stub with clear roadmap note if deferred.
2. **Structured Logging**: integrate `slog` or `zerolog`, gating verbosity via `--debug`.
3. **Read-only Safeguards**: detect file permissions before mutating; fail fast on clipboard ops (§4 Reliability).
4. **Autosave Polish**: status notifications, config validation.
5. **Performance Benchmarks**: load 50k-row CSV and document metrics.

Exit Criteria:
- Logging toggle verified via integration tests capturing stderr.
- Benchmarks recorded in `docs/perf.md`.

### Milestone 4 – MCP Server Integration (Weeks 9–11)
Scope: §5 shared commands, §13 protocol.

Tasks:
1. **Action Abstraction Generalization**: ensure CLI commands call shared action methods accessible to MCP handlers.
2. **MCP Transport** (`pkg/mcp/server.go`): stdio server, session manager, request routing.
3. **Capability Implementations**: `workspace/open`, `cursor/get`/`set`, `cell/edit`, row/column ops, sheet mgmt, save.
4. **Security Layer**: working directory sandbox, read-only mode.
5. **Conformance Tests**: mock MCP client exercising command surface; ensure error codes align with spec table (§13.2).

Exit Criteria:
- CLI can optionally run `ebenezer-go mcp serve`.
- Automated tests validate request/response contract.

### Milestone 5 – Distribution & Release Prep (Weeks 12–13)
Scope: §12, §14.

Tasks:
1. **Cross-Compilation Matrix**: `make release` producing macOS/Linux/Windows binaries <25MB.
2. **Packaging Scripts**: draft Homebrew formula, Scoop manifest templates.
3. **Documentation**: update `README`, add `docs/usage.md`, `docs/mcp.md`.
4. **Telemetry Hooks**: structured event emission for MCP calls (§13.5).

Exit Criteria:
- Release artifacts uploaded to test bucket (or stub) via CI.
- Docs reviewed and linked from root README.

## 3. Cross-Cutting Concerns
- **Dependency Management**: Pin versions via `go.mod`; add Renovate config for upgrades.
- **Error Handling**: Centralize in `internal/app/errors.go` with wrap helpers; ensure no `panic` escapes.
- **Testing Debt Register**: maintain `docs/testing.md` to track gaps (e.g., REPL coverage).
- **UX Consistency**: Add automated snapshot tests for help output and status line.
- **Style Integrity**: Maintain XLSX golden fixtures to detect unintended formatting regressions; add CI diff helper comparing style XML hashes.
- **Security Reviews**: before Milestone 4, run `gosec` and document findings.

## 4. Risk Register & Mitigations
| Risk | Impact | Mitigation |
|------|--------|------------|
| Excelize formula limitations prevent parity (§3.5) | Medium | Prototype with real spreadsheets during Milestone 2; fall back to custom parser if blockers found. |
| Terminal raw-mode inconsistencies across OS | High | Abstract via `golang.org/x/term`, add integration tests on Linux/macOS CI. |
| MCP protocol drift | Medium | Define JSON schema fixtures early; add contract tests in Milestone 4. |
| Binary size creep >25MB | Low | Monitor via `make size` target; strip symbols in release builds. |
| Formatting fidelity regressions when editing styled workbooks (§3.6) | High | Introduce style snapshot helpers in Milestone 1 and run golden regression tests on every edit workflow. |

## 5. Deliverables Checklist
- `IMPLEMENTATION_PLAN.md` kept in sync with milestone execution.
- `docs/usage.md`, `docs/mcp.md`, `docs/perf.md`, `docs/testing.md`.
- Automated CI badges reflecting build/test status.
- Issue templates for bug/feature requests tied to milestones.

## 6. Next Steps
1. Harden the new visual-selection clipboard flows with additional workbook fixtures (multi-sheet, styled ranges) and capture any regressions in `internal/app` tests.
2. Implement column search + MRU history per §3.7 so the CLI can jump between matches and remember recently saved files.
3. Stand up the formula-evaluation scaffold (§3.5) so computed values display alongside raw cell contents.
