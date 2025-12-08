# Action Layer Refactor Roadmap (v2)

| Phase | Scope | Duration (est.) | Deliverables |
|-------|-------|-----------------|--------------|
| 0 | Inventory commands/bindings, stabilize baseline tests | 2 days | `docs/action-layer-inventory.md`, smoke tests documented |
| 1 | Create actions package + pilot (move/goto/status) | 3 days | Action interfaces, pilot migrations, pilot tests |
| 2 | Registry + navigation/edit commands | 5 days | Registry + migrated navigation/edit actions & bindings (keyboard `:` commands remain on Cobra but their handlers call actions). |
| 3 | Formatting, search, save/load actions | 5 days | Style/search/save actions, updated tests |
| 4 | Selection/range + info command actions | 4 days | Visual/range/info actions, selection bindings |
| 5 | Final cleanup + docs/tests | 3 days | Docs updated, `go test ./...`, manual keyboard sign-off |

**Total Estimate**: ~22 working days (~4.5 weeks) with buffer for review.

## Notes
- Telemetry/multi-session explicitly deferred but kept feasible via context design.
- Freeze new command additions during Phases 1–4 to avoid dual migrations.
- After each phase, run `go test ./...` and record results in the refactor document to track regressions early.
