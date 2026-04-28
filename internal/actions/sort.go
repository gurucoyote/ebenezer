package actions

import (
	"fmt"
	"strconv"
	"strings"

	"ebenezer/internal/workbook"
)

var TableSort Action = tableSortAction{}

func init() {
	Register(TableSort)
}

type tableSortAction struct{}

func (tableSortAction) Name() string { return "table-sort" }

func (tableSortAction) Metadata() Metadata {
	return Metadata{
		Name:        "table-sort",
		Description: "Sort table rows by one or more columns",
		Category:    "editing",
		Idempotent:  false,
		Args: []Arg{
			{Name: "keys", Description: "Sort keys as column[:direction] pairs, e.g. \"B:desc C:asc\" or \"2:desc\""},
			{Name: "header_rows", Description: "Number of header rows to keep fixed (default 1)", Optional: true},
			{Name: "range", Description: "A1-style range to restrict operation (e.g. A1:D20)", Optional: true},
		},
	}
}

func (tableSortAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	if ctx.State.Workbook == nil {
		return Result{}, fmt.Errorf("no workbook loaded")
	}

	opts, err := parseSortArgs(args)
	if err != nil {
		return Result{}, err
	}

	// Resolve range into StartRow/EndRow if provided.
	if opts.rangeStr != "" {
		startRow, _, endRow, _, parseErr := workbook.ParseRange(opts.rangeStr)
		if parseErr != nil {
			return Result{}, fmt.Errorf("invalid range %q: %w", opts.rangeStr, parseErr)
		}
		opts.sortOpts.StartRow = startRow
		opts.sortOpts.EndRow = endRow
	}

	if err := ctx.State.Workbook.SortRows(opts.sortOpts); err != nil {
		return Result{}, err
	}

	ctx.State.SetDirty()

	keyDescs := make([]string, len(opts.sortOpts.Keys))
	for i, k := range opts.sortOpts.Keys {
		dir := "asc"
		if !k.Ascending {
			dir = "desc"
		}
		keyDescs[i] = fmt.Sprintf("%s %s", workbook.ColumnName(k.Column), dir)
	}
	msg := fmt.Sprintf("sorted %d rows by %s", opts.sortOpts.EndRow-opts.sortOpts.StartRow+1, strings.Join(keyDescs, ", "))
	return Result{Message: msg}, nil
}

type parsedSortArgs struct {
	sortOpts workbook.SortOpts
	rangeStr string
}

// parseSortArgs converts positional string args into SortOpts.
// Syntax: [keys...] [header_rows=N] [range=A1:D20]
// Keys can be given as positional tokens or via by= flag.
func parseSortArgs(args []string) (parsedSortArgs, error) {
	result := parsedSortArgs{
		sortOpts: workbook.SortOpts{
			HeaderRows: 1,
		},
	}

	var positional []string
	for _, raw := range args {
		token := strings.TrimSpace(raw)
		if token == "" {
			continue
		}
		lower := strings.ToLower(token)
		if eq := strings.IndexRune(lower, '='); eq >= 0 {
			key := strings.TrimSpace(lower[:eq])
			valStr := strings.TrimSpace(token[eq+1:])
			switch key {
			case "header_rows":
				n, err := strconv.Atoi(strings.TrimSpace(valStr))
				if err != nil || n < 1 {
					return result, fmt.Errorf("invalid header_rows %q", valStr)
				}
				result.sortOpts.HeaderRows = n
			case "range":
				result.rangeStr = valStr
			case "by":
				keys, err := parseSortKeys(valStr)
				if err != nil {
					return result, err
				}
				result.sortOpts.Keys = append(result.sortOpts.Keys, keys...)
			default:
				return result, fmt.Errorf("unknown option %q", key)
			}
			continue
		}
		positional = append(positional, token)
	}

	// Parse positional tokens as sort keys.
	// Each key is "COLUMN[:DIRECTION]" where direction defaults to asc.
	for _, token := range positional {
		keys, err := parseSortKeys(token)
		if err != nil {
			return result, err
		}
		result.sortOpts.Keys = append(result.sortOpts.Keys, keys...)
	}

	return result, nil
}

// parseSortKeys parses a string like "B:desc C:asc" or "2:desc" into SortKey slice.
// Direction defaults to asc. Column may be a letter or 1-based number.
func parseSortKeys(s string) ([]workbook.SortKey, error) {
	var keys []workbook.SortKey
	parts := strings.Fields(s)
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		// Split column and direction by colon.
		var colSpec, dirSpec string
		if colon := strings.LastIndex(part, ":"); colon >= 0 {
			colSpec = strings.TrimSpace(part[:colon])
			dirSpec = strings.ToLower(strings.TrimSpace(part[colon+1:]))
		} else {
			colSpec = part
		}
		col, err := resolveColumn(colSpec)
		if err != nil {
			return nil, fmt.Errorf("invalid sort column %q: %w", colSpec, err)
		}
		ascending := true
		switch dirSpec {
		case "desc", "d", "descending":
			ascending = false
		case "", "asc", "a", "ascending":
			ascending = true
		default:
			return nil, fmt.Errorf("invalid sort direction %q for column %q", dirSpec, colSpec)
		}
		keys = append(keys, workbook.SortKey{Column: col, Ascending: ascending})
	}
	return keys, nil
}
