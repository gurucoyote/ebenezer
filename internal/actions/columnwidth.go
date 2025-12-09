package actions

import (
	"fmt"
	"sort"
	"strconv"
	"strings"

	"ebenezer/internal/app"
	"ebenezer/internal/workbook"
)

var (
	ColumnWidth Action = columnWidthAction{}
)

func init() {
	Register(ColumnWidth)
}

type columnWidthAction struct{}

type columnWidthPayload struct {
	Column   string  `json:"column"`
	Index    int     `json:"index"`
	Width    float64 `json:"width"`
	Source   string  `json:"source"`
	Explicit bool    `json:"explicit"`
}

func (columnWidthAction) Name() string { return "colwidth" }

func (columnWidthAction) Metadata() Metadata {
	return Metadata{
		Name:        "colwidth",
		Description: "Inspect or set column widths (show|set|auto)",
		Category:    "formatting",
		Args: []Arg{
			{Name: "mode", Description: "show|set|auto"},
			{Name: "columns", Description: "Optional column span (A:D, 2:5, A1:D9)", Optional: true},
			{Name: "value", Description: "Width when mode=set", Optional: true},
			{Name: "options", Description: "key=value pairs for auto (min=,max=,padding=,bonus=,factor=)", Optional: true, Variadic: true},
		},
	}
}

func (columnWidthAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	if len(args) == 0 {
		return Result{}, fmt.Errorf("usage: colwidth <show|set|auto> ...")
	}
	mode := strings.ToLower(strings.TrimSpace(args[0]))
	switch mode {
	case "show":
		return columnWidthShow(ctx, args[1:])
	case "set":
		return columnWidthSet(ctx, args[1:])
	case "auto":
		return columnWidthAuto(ctx, args[1:])
	default:
		return Result{}, fmt.Errorf("unknown colwidth mode %q", mode)
	}
}

func columnWidthShow(ctx Context, args []string) (Result, error) {
	spec := ""
	if len(args) > 0 {
		spec = args[0]
	}
	start, end, err := ctx.State.ResolveColumnSpan(spec)
	if err != nil {
		return Result{}, err
	}
	infos, err := ctx.State.ColumnWidths(start, end)
	if err != nil {
		return Result{}, err
	}
	payload := buildColumnWidthPayloads(infos)
	msg := fmt.Sprintf("column widths %s:\n%s", spanLabel(start, end), formatColumnWidthEntries(infos))
	return Result{Message: msg, Data: payload}, nil
}

func columnWidthSet(ctx Context, args []string) (Result, error) {
	if len(args) < 2 {
		return Result{}, fmt.Errorf("usage: colwidth set <columns> <width>")
	}
	spec := args[0]
	width, err := strconv.ParseFloat(args[1], 64)
	if err != nil {
		return Result{}, fmt.Errorf("invalid width %q", args[1])
	}
	start, end, err := ctx.State.ResolveColumnSpan(spec)
	if err != nil {
		return Result{}, err
	}
	infos, err := ctx.State.SetColumnWidth(start, end, width)
	if err != nil {
		return Result{}, err
	}
	payload := buildColumnWidthPayloads(infos)
	msg := fmt.Sprintf("set width %.2f for %s\n%s", width, spanLabel(start, end), formatColumnWidthEntries(infos))
	return Result{Message: msg, Data: payload}, nil
}

func columnWidthAuto(ctx Context, args []string) (Result, error) {
	spec, opts, err := parseAutoArgs(args)
	if err != nil {
		return Result{}, err
	}
	start, end, err := ctx.State.ResolveColumnSpan(spec)
	if err != nil {
		return Result{}, err
	}
	infos, err := ctx.State.AutoColumnWidth(start, end, opts)
	if err != nil {
		return Result{}, err
	}
	payload := buildColumnWidthPayloads(infos)
	msg := fmt.Sprintf("auto-fit %s (min %.1f / max %.1f)\n%s", spanLabel(start, end), opts.MinWidth, opts.MaxWidth, formatColumnWidthEntries(infos))
	return Result{Message: msg, Data: payload}, nil
}

func parseAutoArgs(args []string) (string, workbook.ColumnWidthOptions, error) {
	opts := workbook.DefaultColumnWidthOptions()
	spec := ""
	for _, raw := range args {
		token := strings.TrimSpace(raw)
		if token == "" {
			continue
		}
		lower := strings.ToLower(token)
		if eq := strings.IndexRune(lower, '='); eq >= 0 {
			key := strings.TrimSpace(lower[:eq])
			valStr := strings.TrimSpace(lower[eq+1:])
			if valStr == "" {
				return "", opts, fmt.Errorf("missing value for %s", key)
			}
			value, err := strconv.ParseFloat(valStr, 64)
			if err != nil {
				return "", opts, fmt.Errorf("invalid %s value %q", key, valStr)
			}
			switch key {
			case "min":
				opts.MinWidth = value
			case "max":
				opts.MaxWidth = value
			case "padding":
				opts.Padding = value
			case "bonus":
				opts.MultilineBonus = value
			case "factor":
				opts.CharacterFactor = value
			default:
				return "", opts, fmt.Errorf("unknown option %s", key)
			}
			continue
		}
		if spec == "" {
			spec = token
			continue
		}
		return "", opts, fmt.Errorf("unexpected argument %q", token)
	}
	if opts.MaxWidth > 0 && opts.MinWidth > opts.MaxWidth {
		opts.MaxWidth = opts.MinWidth
	}
	return spec, opts, nil
}

func buildColumnWidthPayloads(infos []app.ColumnWidthInfo) []columnWidthPayload {
	payload := make([]columnWidthPayload, len(infos))
	for i, info := range infos {
		payload[i] = columnWidthPayload{
			Column:   workbook.ColumnName(info.Column),
			Index:    info.Column,
			Width:    info.Width,
			Source:   info.Source,
			Explicit: info.Explicit,
		}
	}
	sort.Slice(payload, func(i, j int) bool { return payload[i].Index < payload[j].Index })
	return payload
}

func formatColumnWidthEntries(infos []app.ColumnWidthInfo) string {
	var b strings.Builder
	payload := make([]app.ColumnWidthInfo, len(infos))
	copy(payload, infos)
	sort.Slice(payload, func(i, j int) bool { return payload[i].Column < payload[j].Column })
	for i, info := range payload {
		fmt.Fprintf(&b, "- %s (%d): %.1f", workbook.ColumnName(info.Column), info.Column, info.Width)
		if !info.Explicit {
			b.WriteString(" [default]")
		} else {
			switch info.Source {
			case "auto":
				b.WriteString(" [auto]")
			case "set":
				b.WriteString(" [set]")
			}
		}
		if i < len(payload)-1 {
			b.WriteByte('\n')
		}
	}
	if b.Len() == 0 {
		return "(no columns)"
	}
	return b.String()
}

func spanLabel(start, end int) string {
	if start == end {
		return fmt.Sprintf("column %s", workbook.ColumnName(start))
	}
	return fmt.Sprintf("columns %s-%s", workbook.ColumnName(start), workbook.ColumnName(end))
}
