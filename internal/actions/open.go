package actions

import (
	"errors"
	"fmt"
	"path/filepath"
	"strings"

	"ebenezer/internal/workbook"
)

var (
	OpenFile   Action = openAction{}
	SampleData Action = sampleAction{}
)

func init() {
	Register(OpenFile, "open")
	Register(SampleData, "sample")
}

type openAction struct{}

type sampleAction struct{}

func (openAction) Name() string { return "open" }

func (openAction) Metadata() Metadata {
	return Metadata{
		Name:        "open",
		Description: "Open a workbook from disk",
		Category:    "file",
		Args: []Arg{
			{Name: "path", Description: "Path to CSV/XLSX file"},
			{Name: "--sheet", Description: "Optional sheet when opening XLSX", Optional: true},
		},
	}
}

func (openAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	if len(args) == 0 {
		return Result{}, errors.New("open requires a file path")
	}
	path := args[0]
	sheet := ""
	for _, arg := range args[1:] {
		if strings.HasPrefix(arg, "--sheet=") {
			sheet = strings.TrimPrefix(arg, "--sheet=")
		}
	}
	wb, sheets, active, err := workbook.FromFile(path, sheet, workbook.WithCSVDelimiter(ctx.State.CSVDelimiter()))
	if err != nil {
		return Result{}, err
	}
	if strings.EqualFold(filepath.Ext(path), ".xlsx") {
		scanSheet := sheet
		if scanSheet == "" {
			scanSheet = active
		}
		if warnings, err := workbook.ScanRichText(path, scanSheet); err != nil {
			wb.Warnings = append(wb.Warnings, fmt.Sprintf("rich text scan failed: %v", err))
		} else {
			wb.RichTextRuns = workbook.RichTextRunMap(warnings, scanSheet)
			wb.RichTextSheet = scanSheet
			for _, warning := range warnings {
				wb.Warnings = append(wb.Warnings, fmt.Sprintf("rich text in %s!%s (%d runs)", warning.Sheet, warning.Cell, warning.Runs))
			}
		}
	}
	ctx.State.LoadWorkbook(wb, path, sheets, active)
	msg := fmt.Sprintf("loaded %s [%s] (%d rows)\n", filepath.Base(wb.Name), wb.Sheet, len(wb.Cells))
	if len(wb.Warnings) > 0 {
		var b strings.Builder
		b.WriteString(msg)
		b.WriteString("warnings:\n")
		for _, warning := range wb.Warnings {
			b.WriteString("  - ")
			b.WriteString(warning)
			b.WriteByte('\n')
		}
		msg = b.String()
	}
	return Result{Message: msg}, nil
}

func (sampleAction) Name() string { return "sample" }

func (sampleAction) Metadata() Metadata {
	return Metadata{
		Name:        "sample",
		Description: "Load the built-in sample workbook",
		Category:    "file",
		Idempotent:  true,
	}
}

func (sampleAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	wb := workbook.SampleWorkbook()
	ctx.State.LoadWorkbook(wb, "", []string{wb.Sheet}, wb.ActiveCell)
	return Result{Message: "loaded sample workbook\n"}, nil
}
