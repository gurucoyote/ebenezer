package actions

import (
	"errors"
	"fmt"
	"path/filepath"
	"strings"

	"ebenezer/internal/workbook"
)

var (
	SheetList   Action = sheetListAction{}
	SheetSwitch Action = sheetSwitchAction{}
	SheetCreate Action = sheetCreateAction{}
)

func init() {
	Register(SheetList, "ps-list")
	Register(SheetSwitch, "ps")
	Register(SheetCreate, "ns")
}

type sheetListAction struct{}

type sheetSwitchAction struct{}

type sheetCreateAction struct{}

func (sheetListAction) Name() string { return "ps-list" }

func (sheetListAction) Metadata() Metadata {
	return Metadata{
		Name:        "ps-list",
		Description: "List sheets in the current workbook",
		Category:    "sheet",
		Idempotent:  true,
	}
}

func (sheetListAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	if len(ctx.State.SheetNames) == 0 {
		return Result{Message: "no sheet metadata available\n"}, nil
	}
	current := ""
	if ctx.State.Workbook != nil {
		current = ctx.State.Workbook.Sheet
	}
	var b strings.Builder
	for _, name := range ctx.State.SheetNames {
		marker := " "
		if strings.EqualFold(name, current) {
			marker = "*"
		}
		fmt.Fprintf(&b, "%s %s\n", marker, name)
	}
	return Result{Message: b.String()}, nil
}

func (sheetSwitchAction) Name() string { return "ps" }

func (sheetSwitchAction) Metadata() Metadata {
	return Metadata{
		Name:        "ps",
		Description: "List or switch sheets",
		Category:    "sheet",
		Args:        []Arg{{Name: "sheet", Description: "Optional sheet name", Optional: true}},
	}
}

func (sheetSwitchAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	if len(args) == 0 {
		return SheetList.Exec(ctx, nil)
	}
	if ctx.State.SourcePath == "" {
		return Result{}, errors.New("current workbook not backed by a file; pass a filename to switch sheets")
	}
	sheet := args[0]
	wb, sheets, active, err := workbook.FromFile(ctx.State.SourcePath, sheet)
	if err != nil {
		return Result{}, err
	}
	ctx.State.LoadWorkbook(wb, ctx.State.SourcePath, sheets, active)
	return Result{Message: fmt.Sprintf("switched to sheet %s\n", wb.Sheet)}, nil
}

func (sheetCreateAction) Name() string { return "ns" }

func (sheetCreateAction) Metadata() Metadata {
	return Metadata{
		Name:        "ns",
		Description: "Create a new sheet, optionally copying an existing one",
		Category:    "sheet",
		Args: []Arg{
			{Name: "name", Description: "New sheet name"},
			{Name: "--copy", Description: "Optional source sheet", Optional: true},
		},
	}
}

func (sheetCreateAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	if len(args) == 0 {
		return Result{}, errors.New("sheet name required")
	}
	if ctx.State.SourcePath == "" || strings.ToLower(filepath.Ext(ctx.State.SourcePath)) != ".xlsx" {
		return Result{}, errors.New("sheet creation requires an .xlsx file saved on disk")
	}
	name := args[0]
	copyFrom := ""
	for _, arg := range args[1:] {
		if strings.HasPrefix(arg, "--copy=") {
			copyFrom = strings.TrimPrefix(arg, "--copy=")
		}
	}
	if err := workbook.AddSheet(ctx.State.SourcePath, name, copyFrom); err != nil {
		return Result{}, err
	}
	wb, sheets, active, err := workbook.FromFile(ctx.State.SourcePath, name)
	if err != nil {
		return Result{}, err
	}
	ctx.State.LoadWorkbook(wb, ctx.State.SourcePath, sheets, active)
	return Result{Message: fmt.Sprintf("created sheet %s\n", name)}, nil
}
