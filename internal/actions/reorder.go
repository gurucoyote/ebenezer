package actions

import (
	"fmt"
	"strconv"
	"strings"

	"ebenezer/internal/workbook"
)

var TableReorder Action = tableReorderAction{}

func init() {
	Register(TableReorder)
}

type tableReorderAction struct{}

func (tableReorderAction) Name() string { return "table-reorder" }

func (tableReorderAction) Metadata() Metadata {
	return Metadata{
		Name:        "table-reorder",
		Description: "Reorder table rows using a stable partition on a column predicate",
		Category:    "editing",
		Idempotent:  false,
		Args: []Arg{
			{Name: "mode", Description: "Reorder mode (currently only \"partition\")", Optional: true},
			{Name: "column", Description: "1-based column index or column letter (e.g. \"2\" or \"B\")"},
			{Name: "value", Description: "Comparison value for equals/equals_ignore_case predicates", Optional: true},
			{Name: "predicate", Description: "Predicate: equals|equals_ignore_case|is_blank|is_non_blank (default equals_ignore_case)", Optional: true},
			{Name: "case_insensitive", Description: "Set \"false\" to make predicate case-sensitive (default true)", Optional: true},
			{Name: "separator_rows", Description: "Blank rows inserted between partitions (default 0)", Optional: true},
			{Name: "header_rows", Description: "Number of header rows to keep fixed (default 1)", Optional: true},
			{Name: "range", Description: "A1-style range to restrict operation (e.g. A1:D20)", Optional: true},
		},
	}
}

func (tableReorderAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	if ctx.State.Workbook == nil {
		return Result{}, fmt.Errorf("no workbook loaded")
	}

	opts, err := parseReorderArgs(args)
	if err != nil {
		return Result{}, err
	}

	// Resolve range into StartRow/EndRow if provided.
	if opts.rangeStr != "" {
		startRow, _, endRow, _, parseErr := workbook.ParseRange(opts.rangeStr)
		if parseErr != nil {
			return Result{}, fmt.Errorf("invalid range %q: %w", opts.rangeStr, parseErr)
		}
		opts.spOpts.StartRow = startRow
		opts.spOpts.EndRow = endRow
	}

	matching, err := ctx.State.Workbook.StablePartition(opts.spOpts)
	if err != nil {
		return Result{}, err
	}

	// Mark the workbook as modified.
	ctx.State.SetDirty()

	position := "bottom"
	if opts.spOpts.MatchingPosition == "top" {
		position = "top"
	}

	totalDataRows := opts.spOpts.EndRow - opts.spOpts.StartRow + 1
	nonMatching := totalDataRows - matching

	msg := fmt.Sprintf("reordered %d rows: %d matching moved to %s, %d non-matching", totalDataRows, matching, position, nonMatching)
	return Result{Message: msg}, nil
}

// parsedReorderArgs holds the intermediate result of argument parsing.
type parsedReorderArgs struct {
	spOpts   workbook.StablePartitionOpts
	rangeStr string
}

// parseReorderArgs converts positional string args into a StablePartitionOpts.
// Arg layout mirrors columnwidth.go: positional tokens with optional key=value flags.
func parseReorderArgs(args []string) (parsedReorderArgs, error) {
	result := parsedReorderArgs{
		spOpts: workbook.StablePartitionOpts{
			HeaderRows:   1,
			Predicate:    "equals_ignore_case",
			MatchingPosition: "bottom",
		},
	}

	// Collect positional args and key=value pairs separately.
	var positional []string
	caseInsensitive := true // default

	for _, raw := range args {
		token := strings.TrimSpace(raw)
		if token == "" {
			continue
		}
		lower := strings.ToLower(token)
		if eq := strings.IndexRune(lower, '='); eq >= 0 {
			key := strings.TrimSpace(lower[:eq])
			valStr := strings.TrimSpace(token[eq+1:]) // preserve original case for value
			switch key {
			case "predicate":
				result.spOpts.Predicate = lower[eq+1:]
			case "case_insensitive":
				lowerVal := strings.ToLower(strings.TrimSpace(valStr))
				caseInsensitive = lowerVal != "false"
			case "separator_rows":
				n, err := strconv.Atoi(strings.TrimSpace(lower[eq+1:]))
				if err != nil || n < 0 {
					return result, fmt.Errorf("invalid separator_rows %q", valStr)
				}
				result.spOpts.SeparatorRows = n
			case "header_rows":
				n, err := strconv.Atoi(strings.TrimSpace(lower[eq+1:]))
				if err != nil || n < 1 {
					return result, fmt.Errorf("invalid header_rows %q", valStr)
				}
				result.spOpts.HeaderRows = n
			case "range":
				result.rangeStr = valStr
			case "value":
				result.spOpts.Value = valStr
			case "mode":
				mode := strings.TrimSpace(lower[eq+1:])
				if mode != "partition" {
					return result, fmt.Errorf("unsupported mode %q (only \"partition\" is supported)", mode)
				}
			default:
				return result, fmt.Errorf("unknown option %q", key)
			}
			continue
		}
		positional = append(positional, token)
	}

	// Resolve positional args: [mode] [column] [value] [predicate]
	// Mode comes first if it matches "partition".
	idx := 0
	if idx < len(positional) && strings.ToLower(positional[idx]) == "partition" {
		idx++
	}

	if idx >= len(positional) {
		return result, fmt.Errorf("column argument is required")
	}

	// Resolve column: could be a number or a letter.
	col, err := resolveColumn(positional[idx])
	if err != nil {
		return result, fmt.Errorf("invalid column %q: %w", positional[idx], err)
	}
	result.spOpts.KeyColumn = col
	idx++

	// Remaining positional: value, then predicate.
	if idx < len(positional) {
		result.spOpts.Value = positional[idx]
		idx++
	}
	if idx < len(positional) {
		result.spOpts.Predicate = strings.ToLower(positional[idx])
		idx++
	}

	// Apply case_insensitive override: if true and predicate is "equals", upgrade.
	if caseInsensitive && result.spOpts.Predicate == "equals" {
		result.spOpts.Predicate = "equals_ignore_case"
	}

	// Validate predicate.
	validPreds := map[string]bool{
		"equals":              true,
		"equals_ignore_case":  true,
		"is_blank":            true,
		"is_non_blank":        true,
	}
	if !validPreds[result.spOpts.Predicate] {
		return result, fmt.Errorf("unsupported predicate %q", result.spOpts.Predicate)
	}

	// equals/equals_ignore_case require a value.
	if (result.spOpts.Predicate == "equals" || result.spOpts.Predicate == "equals_ignore_case") && result.spOpts.Value == "" {
		return result, fmt.Errorf("value is required for predicate %q", result.spOpts.Predicate)
	}

	return result, nil
}

// resolveColumn converts a column spec (number like "2" or letter like "B") to a 1-based column index.
func resolveColumn(spec string) (int, error) {
	// Try as a number first.
	if n, err := strconv.Atoi(strings.TrimSpace(spec)); err == nil {
		if n < 1 {
			return 0, fmt.Errorf("must be >= 1")
		}
		return n, nil
	}
	// Try as column letter(s) via workbook.ParseColumnSpec.
	start, _, err := workbook.ParseColumnSpec(spec)
	if err != nil {
		return 0, err
	}
	return start, nil
}


