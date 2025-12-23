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
- `go test ./...` runs existing unit tests. **Always set `GOCACHE=/root/ebenezer/.cache` first**, e.g. `GOCACHE=/root/ebenezer/.cache go test ./...`, because the default `/root/.cache/go-build` is read-only in this environment and causes misleading `testing/internal/testdeps` errors. Add table-driven tests beside new code.
- Keyboard `:` command mode intentionally routes through Cobra so humans can keep typing Vim-like commands, but every Cobra handler must call into the shared action layer so behavior stays reusable across keyboard, CLI, and future MCP/AI integrations.
- Every action **must** declare metadata (name, description, category, arg info, idempotency) and register via `actions.Register`. The action-layer tests iterate the registry and will fail if metadata is missing or incomplete.

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
- Each discrete work step should land as an atomic commit with a meaningful message (continue using Conventional Commit prefixes). If you discover mid-task that another logical change is required, finish the current commit first before starting a new scoped commit.
- Never remove or alter user modifications outside your scope; check `git status` before editing shared files.

## Agent Tips
- When you need styles or formula metadata, prefer the Excelize APIs already in `internal/workbook`—don’t reimplement parsing.
- For new keyboard shortcuts, wire them via `internal/cmd/keyboard.go` so they dispatch through Cobra commands; update docs accordingly.
- If you introduce new dependencies, update `go.mod`/`go.sum` and note the reason in the PR summary.
- Always keep `README.md`, `docs/demo.md`, and the implementation plan in sync with user-facing changes.
- MCP-facing work must rely on the action metadata pipeline: `actions.Discover()` returns the JSON-ready schema that MCP and other headless clients will consume. Whenever you add or change an action, update its metadata, confirm `Discover()` exposes the new shape via `go test ./internal/actions`, and sanity-check the CLI surface with `./ebenezer actions list`.

## User Story Capture Helper
- Trigger this helper whenever the user explicitly asks for a “user story”, “story card”, or similar phrasing. Confirm intent if the request is ambiguous.
- Collect the classic triad (`As a …`, `I want …`, `So that …`) plus acceptance criteria and metadata (priority, tags, related issues) by asking follow-up questions until each field is clear.
- Present the formatted story back to the user for approval before saving anything. Use this template:
  ```
  As a <role>,
  I want <goal>,
  So that <reason>.

  Acceptance Criteria:
  - <criterion>

  Notes:
  - Priority: <P1/P2/etc>
  - Tags: <comma-separated>
  - Source: <link or “chat request”>
  ```
- Once the user approves, append the story to `user-stories.txt` in the repo root. Separate entries with a `---` line and include an ISO-8601 timestamp plus the agent name handling the request.
- If `user-stories.txt` is missing, create it before writing (keep it under git). Never overwrite prior stories; always append.
- Mention in your final message that the story was captured and logged so humans know the workflow ran end-to-end.

## Blood Mirror Critic Helper
- Invoke Blood Mirror whenever the user explicitly requests a harsh critique, “blood mirror”, or expert review of the current codebase/feature. Default to this helper for any broad “review the project” prompt unless the user asks for implementation work instead.
- Before writing the critique, skim the relevant code, `SPEC.md`, `IMPLEMENTATION_PLAN.md`, and `user-stories.txt` so findings anchor to actual requirements (cite story IDs or spec sections when possible).
- Deliver unvarnished observations: call out architectural drift, missing acceptance criteria, stale docs, lack of tests, regressions, or process gaps. Avoid compliments, hedging, or placating language—the helper’s purpose is to push the project toward its vision.
- Structure output as:
  1. **Findings** ordered by severity, each referencing files/lines or story/spec IDs.
  2. **Gaps/Risks** describing unknowns or missing data.
  3. **Next Actions** suggesting concrete remediation steps.
- Do not attempt to fix issues during the review; the helper purely diagnoses. If the user subsequently wants fixes, switch out of Blood Mirror mode and follow normal implementation workflows.

## Landing the Plane (Session Completion)

**When ending a work session**, you MUST complete ALL steps below. Work is NOT complete until `git push` succeeds.

**MANDATORY WORKFLOW:**

1. **File issues for remaining work** - Create issues for anything that needs follow-up
2. **Run quality gates** (if code changed) - Tests, linters, builds
3. **Update issue status** - Close finished work, update in-progress items
4. **PUSH TO REMOTE** - This is MANDATORY:
   ```bash
   git pull --rebase
   bd sync
   git push
   git status  # MUST show "up to date with origin"
   ```
5. **Clean up** - Clear stashes, prune remote branches
6. **Verify** - All changes committed AND pushed
7. **Hand off** - Provide context for next session

**CRITICAL RULES:**
- Work is NOT complete until `git push` succeeds
- NEVER stop before pushing - that leaves work stranded locally
- NEVER say "ready to push when you are" - YOU must push
- If push fails, resolve and retry until it succeeds


<!-- BEGIN BEADS INTEGRATION -->
## Issue Tracking with bd (beads)

**IMPORTANT**: This project uses **bd (beads)** for ALL issue tracking. Do NOT use markdown TODOs, task lists, or other tracking methods.

**Work in direct mode by default:** bd can be slow to start its daemon and emits warnings; add `--no-daemon` to your bd commands unless you explicitly need the daemon.

### Why bd?

- Dependency-aware: Track blockers and relationships between issues
- Git-friendly: Auto-syncs to JSONL for version control
- Agent-optimized: JSON output, ready work detection, discovered-from links
- Prevents duplicate tracking systems and confusion

### Quick Start

**Quick note:** run bd in direct mode (`--no-daemon`) so we avoid the repeated daemon warnings.

**Check for ready work:**

```bash
bd ready --json --no-daemon
```

**Create new issues:**

```bash
bd create "Issue title" --description="Detailed context" -t bug|feature|task -p 0-4 --json
bd create "Issue title" --description="What this issue is about" -p 1 --deps discovered-from:bd-123 --json
```

**Claim and update:**

```bash
bd update bd-42 --status in_progress --json
bd update bd-42 --priority 1 --json
```

**Complete work:**

```bash
bd close bd-42 --reason "Completed" --json
```

### Issue Types

- `bug` - Something broken
- `feature` - New functionality
- `task` - Work item (tests, docs, refactoring)
- `epic` - Large feature with subtasks
- `chore` - Maintenance (dependencies, tooling)

### Priorities

- `0` - Critical (security, data loss, broken builds)
- `1` - High (major features, important bugs)
- `2` - Medium (default, nice-to-have)
- `3` - Low (polish, optimization)
- `4` - Backlog (future ideas)

### Workflow for AI Agents

1. **Check ready work**: `bd ready` shows unblocked issues
2. **Claim your task**: `bd update <id> --status in_progress`
3. **Work on it**: Implement, test, document
4. **Discover new work?** Create linked issue:
   - `bd create "Found bug" --description="Details about what was found" -p 1 --deps discovered-from:<parent-id>`
5. **Complete**: `bd close <id> --reason "Done"`

### Auto-Sync

bd automatically syncs with git:

- Exports to `.beads/issues.jsonl` after changes (5s debounce)
- Imports from JSONL when newer (e.g., after `git pull`)
- No manual export/import needed!

### Important Rules

- ✅ Use bd for ALL task tracking
- ✅ Always use `--json` flag for programmatic use
- ✅ Link discovered work with `discovered-from` dependencies
- ✅ Check `bd ready` before asking "what should I work on?"
- ❌ Do NOT create markdown TODO lists
- ❌ Do NOT use external issue trackers
- ❌ Do NOT duplicate tracking systems

For more details, see README.md and docs/QUICKSTART.md.

<!-- END BEADS INTEGRATION -->
