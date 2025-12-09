# Action Metadata, Logging, and MCP Readiness Plan (v2)

## Objectives
1. **Schema & Enforcement**: Define metadata schema for every action (name, summary, args, categories, idempotency) and enforce it at compile/test time.
2. **Structured Logging**: Provide a pluggable logging pipeline that records action executions (start/end, duration, status, session/file context) with a clear sink (stderr by default, hook for MCP).
3. **Testing Harness**: Expand action tests and add integration checks (CLI/keyboard) to verify metadata/logging, ensuring MCP consumers can trust the action layer.
4. **MCP Alignment**: Shape the metadata/logging formats to feed directly into upcoming MCP server discovery and telemetry without another refactor.

## Milestones
### M1. Metadata Schema & Registration (5 days)
- Introduce `actions.Metadata` (Name, Description, Args []ArgMeta, Category, Idempotent, Experimental flags).
- Update `Action` interface to expose `Metadata() Metadata`.
- Require metadata during registration (`Register` panics if fields missing); add compile-time helpers to reduce boilerplate.
- Migrate existing actions (batch by package) and add a unit test that iterates `actions.List()` ensuring descriptions/args/categories exist.

### M2. Logging Infrastructure (6 days)
- Define `actions.Logger` interface with `Before(ctx, meta, args)` and `After(ctx, meta, args, result, err, duration)` callbacks.
- Extend `actions.Context` to carry a logger + session ID placeholder.
- Implement default stderr JSON logger (fields: timestamp, action, args summary, duration_ms, status, file/path cues). Allow disabling via env/config.
- Update CLI/keyboard adapters (`executeAction`, keyboard helper) to pass loggers and record start/end times.
- Provide a no-op logger for tests and a hook for MCP to supply its own logger.

### M3. MCP-Facing Metadata & Discovery (4 days)
- Serialize metadata to a JSON schema (e.g., `actions.MetadataDTO`) so MCP can expose a `actions/list` capability. **Status**: ✅ landed via `internal/actions/discover.go`.
- Add a helper (`actions.Discover() []MetadataDTO`) that MCP server will call. The DTO includes `{name, description, category, args[], idempotent, experimental}` with JSON tags so MCP responses are machine-friendly.
- Document any additional MCP requirements (e.g., action categories map to MCP namespaces) and ensure metadata fields cover them. The schema mirrors what `/root/roderik/cmd/mcp.go` advertises through `mark3labs/mcp-go`, so agents can reuse client assumptions across both projects.
- Provide a JSON example in `docs/mcp-readiness-summary.md` to keep MCP consumers aligned on field naming and casing.
- CLI and MCP parity checks now exist: `ebenezer actions list` pretty-prints the payload for humans, while the MCP `actions_list` tool streams the compact JSON for remote agents.

### M4. Testing Enhancements (5 days)
- Expand `internal/actions/actions_test.go` to verify metadata per action (non-empty description, arg info matches expected behavior).
- Add logging tests using a fake logger to ensure `Before/After` fire with correct data, including error scenarios.
- Add an integration test (Go test or script) that runs a handful of commands through Cobra/keyboard, captures stderr, and asserts JSON logs exist.
- Document required tests for new actions (metadata assertion + action behavior + logging check) in `AGENTS.md`.

### M5. Documentation & MPC Prep (2 days)
- Update `SPEC.md §5`, `IMPLEMENTATION_PLAN.md`, `AGENTS.md`, and `docs/action-layer-*` with metadata/logging requirements, enforcement steps, and MCP tie-in.
- Draft a short MCP server plan referencing the metadata/logging outputs (e.g., `actions/discover`, `actions/run`).

## Total Estimate
~22 working days (5 + 6 + 4 + 5 + 2), acknowledging the broad code/doc/test impact.

## Risks & Mitigations
- **Large migration footprint**: tackle actions package-by-package, running `gofmt`/tests after each batch; gate merges behind the metadata test.
- **Logging overhead**: default logger will be opt-in (enabled via env flag); when disabled it’s a no-op to avoid hot-path cost.
- **MCP dependency**: coordinate metadata schema with MCP design doc before implementation; treat schema as semi-stable once merged.
