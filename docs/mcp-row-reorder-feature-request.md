# Feature Request: Stable Row Reordering for MCP Workflows

## Summary

Ebenezer's MCP surface currently exposes low-level workbook primitives well enough for cell editing, selection management, row insertion/deletion, and formatting work. It does not yet offer an efficient way to reorder whole table rows based on a column predicate while preserving the relative order of the matching and non-matching groups.

That gap forces agents to fall back to external workbook libraries for a common spreadsheet-maintenance workflow that should stay inside the Ebenezer session model.

## User Scenario

Given a workbook table such as:

- `Datum`
- `Erledigt (X oder leer)`
- `Todo Text`
- `Notiz`

an agent should be able to:

1. keep all unfinished rows at the top,
2. move all rows whose `Erledigt` cell contains `X` or `x` to the bottom,
3. insert exactly one blank separator row before the completed block, and
4. preserve the original order within both groups.

This is a stable partition operation over table rows.

## Current Limitation

The current MCP tool set can approximate this only through a noisy sequence of low-level operations such as:

- exporting a range,
- computing the desired row order outside Ebenezer,
- inserting/deleting rows repeatedly, and
- rewriting many cells one by one.

That approach has several problems:

- It is inefficient for agents and hard to express cleanly.
- It increases the risk of partial or inconsistent edits.
- It bypasses the session/action layer when external scripting is used.
- It makes verification harder because the user's intent is "reorder rows" but the tool sequence is dozens of primitive mutations.

## Requested Capability

Add a first-class MCP operation for bulk row reordering over a rectangular table range.

Two viable API shapes:

### Option A: Generic Reorder Tool

`table_reorder`

Suggested inputs:

- `session_id`
- `range`
- `header_rows` (default `1`)
- `mode`: `stable_partition` | `sort`
- `key_column`: column name or relative index
- `predicate`: e.g. `equals`, `equals_ignore_case`, `is_blank`, `is_non_blank`
- `value`
- `matching_position`: `top` | `bottom`
- `separator_blank_rows`: integer
- `preserve_formatting`: boolean

### Option B: Narrower Task-Focused Tool

`rows_move_matching_to_end`

Suggested inputs:

- `session_id`
- `range`
- `header_rows`
- `match_column`
- `match_value`
- `case_insensitive`
- `blank_rows_before_block`
- `preserve_formatting`

## Expected Behavior

Any implementation should:

- treat each row in the selected data range as a single record,
- preserve row integrity across all columns in the range,
- leave header rows untouched,
- preserve relative order within both partitions,
- insert the requested number of blank separator rows between partitions,
- preserve formulas, styles, row heights, and comments where possible,
- fail clearly when merged cells or other sheet structures make the reorder ambiguous.

## Why This Belongs in Ebenezer

This workflow is common in real planning sheets, booking trackers, todo lists, and operational spreadsheets. A high-level row-partition tool would:

- keep automation inside the shared action/session model,
- reduce MCP chatter and error surface,
- better match user intent,
- make behavior easier to test at the action layer, and
- eliminate unnecessary fallbacks to external tools such as OpenPyXL.

## Implementation Notes

The most natural fit appears to be a shared action plus workbook helper, exposed through both Cobra and MCP:

- workbook helper performs stable row partitioning over a bounded range,
- action layer defines the user-facing arguments and metadata,
- MCP tool publishes the capability through `actions.Discover()`,
- CLI/command mode could later expose a related command for interactive users.

This should be treated as row-level table manipulation, not as a cell-by-cell macro.

## Acceptance Criteria

- Agents can request stable partitioning of rows within a rectangular range without external scripting.
- Matching and non-matching rows preserve internal order.
- Header rows remain fixed.
- Optional blank separator rows can be inserted between partitions.
- The operation is discoverable through MCP action metadata.
- Tests cover value movement, order stability, and formatting preservation expectations.
