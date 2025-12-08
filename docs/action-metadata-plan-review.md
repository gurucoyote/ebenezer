# Blood Mirror Review – Action Metadata Plan

**Findings**
- The plan assumes metadata/logging/testing can be bolted on without touching existing actions. Reality: every action file must be updated to provide metadata; there’s no allowance for refactoring cost or coordination with other contributors. Expect churn across 20+ files, not captured in the timeline.
- Structured logging is defined vaguely (“JSON lines with fields…”) but no sink/transport strategy exists. Are we logging to stderr, to a file, to MCP clients? Without a concrete destination, logging will be inconsistent and the tests won’t know what to assert.
- Testing workstream talks about asserting metadata presence but doesn’t specify how to enforce it (build failure? lint?). The plan doesn’t mention updating registry registration (e.g., making metadata required at compile time). Without enforcement, contributors will forget to add metadata and the tests won’t catch it reliably.
- Timeline is fantasy. Metadata + logging touches every action, every adapter, and every doc. 12 working days might cover design docs; implementation + reviews + regression tests will easily double that. You’ve already seen how long migrating commands took. This plan underestimates the complexity.
- No mention of how metadata/logging will feed into the MCP server. If we’re investing in metadata, we should document the contract (e.g., JSON schema) so MCP development can proceed in parallel. Otherwise, MCP work will block waiting for the metadata definitions to stabilize.

**Gaps/Risks**
- Registry API changes aren’t defined. Introducing `Metadata` likely requires breaking changes to `Action` interface and registration functions. Plan ignores migration strategy or how to handle third-party actions (if any) during the transition.
- Logging hooks may incur performance penalties (timers, JSON encoding) on every keypress. No mitigation plan (e.g., sampling, disabled-by-default) is discussed.
- Testing workstream skips CLI/manual validation. Logging is notoriously environment-specific; without end-to-end tests (run CLI, inspect logs), we risk shipping logs that don’t actually appear in real workflows.

**Next Actions**
1. Redo the planning with realistic scope/time: break metadata/logging/testing into smaller milestones, include migration of existing actions, and specify enforcement mechanisms (e.g., compile-time checks, go vet rules).
2. Define the logging sink and schema up front. Decide where logs go (stderr vs. file vs. callback). Without that, you can’t implement or test anything meaningful.
3. Coordinate with MCP requirements. If metadata feeds MCP discovery, specify the JSON schema and include MCP stories/tasks so the work converges.
4. Expand testing strategy to include integration/CLI verification, not just unit tests. You need at least one automated test that runs an action through the CLI/keyboard path and asserts logs/metadata show up.
5. Add risk mitigation: how to roll out metadata/logging gradually, how to guard performance, how to educate contributors.
