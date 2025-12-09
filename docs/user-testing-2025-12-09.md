# User Testing Report – 2025-12-09 (Workbook Formatting Session)

## Context
- Scenario: Exploratory edit of `docs/user-stories-status.xlsx` to mirror `user-stories.txt` and experiment with readability tweaks.
- Goal: Treat the workbook as a sandbox (“plaything”) to uncover gaps in the CLI + MCP tooling without assuming the file is authoritative.

## Observations
1. **Source-of-truth mismatch:** All canonical backlog data lives in `user-stories.txt`, yet mirroring it in Excel is manual and error-prone. There is no guardrail preventing drift between the text log and the demo workbook.
2. **Column dimension blind spots:** MCP exposes `style_describe/apply`, but there is no way to read or set column widths via CLI/MCP, so agents fall back to external scripts (OpenPyXL) that bypass session safety.
3. **Palette ambiguity:** The workbook encodes meaning through fill colors (P1 red, Planned gray, etc.), but there is no discoverable schema or helper describing those semantics for automation.
4. **Heuristic opacity:** The “auto width” heuristic (longest line + padding + clamps) exists only in ad-hoc scripts. Reproducing or tuning it later would require spelunking git history instead of reusing a shared helper.

## Enhancement Paths
- **P2: Column Width Inspection & Setting** – Add shared helpers in `internal/workbook` to measure longest cell content, expose `SetColWidth` bindings, and surface them through both Cobra (e.g., `:colwidth auto A:E`) and MCP (`column_width` tool with `mode=auto|set`). Include width-read APIs so automation can diff the before/after state.
- **Backlog Sync Command** – Provide an action that reads `user-stories.txt` and regenerates demo workbooks (CSV/XLSX) deterministically, ensuring color codes and widths stay consistent without manual copying.
- **Style Legend Metadata** – Extend `info_get` (and docs) to publish a style legend per sheet so MCP clients understand what each fill/font denotes, enabling validation and richer automation.
- **Heuristic Library** – Capture the width-estimation heuristic inside the codebase (unit-tested) so both CLI and MCP can reuse it, eliminating the need for external scripting during user sessions.

## Takeaways
- Treat Excel artifacts strictly as demo sandboxes; always update `user-stories.txt` first and regenerate any visual assets via tools.
- Formalize any experimental workflow that proves useful (column auto-fit, palette auditing) into reusable actions so future user tests happen inside Ebenezer instead of external editors.
