package actions

import (
	"fmt"
	"strings"

	"ebenezer/internal/workbook"
)

var RichTextScan Action = richTextScanAction{}

func init() {
	Register(RichTextScan, "richtext-scan")
}

type richTextScanAction struct{}

func (richTextScanAction) Name() string { return "richtext-scan" }

func (richTextScanAction) Metadata() Metadata {
	return Metadata{
		Name:        "richtext-scan",
		Description: "Scan an XLSX file for inline rich text runs",
		Category:    "file",
		Args: []Arg{
			{Name: "path", Description: "Path to XLSX file"},
			{Name: "--sheet", Description: "Optional sheet to scan", Optional: true},
		},
	}
}

func (richTextScanAction) Exec(ctx Context, args []string) (Result, error) {
	if len(args) == 0 {
		return Result{}, fmt.Errorf("richtext-scan requires a file path")
	}
	path := args[0]
	sheet := ""
	for _, arg := range args[1:] {
		if strings.HasPrefix(arg, "--sheet=") {
			sheet = strings.TrimPrefix(arg, "--sheet=")
		}
	}
	warnings, err := workbook.ScanRichText(path, sheet)
	if err != nil {
		return Result{}, err
	}
	if len(warnings) == 0 {
		return Result{Message: "no rich text runs found\n", Data: warnings}, nil
	}
	var b strings.Builder
	fmt.Fprintf(&b, "rich text runs found (%d):\n", len(warnings))
	for _, warning := range warnings {
		fmt.Fprintf(&b, "  - %s!%s (%d runs)\n", warning.Sheet, warning.Cell, warning.Runs)
	}
	return Result{Message: b.String(), Data: warnings}, nil
}
