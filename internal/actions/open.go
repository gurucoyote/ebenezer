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
	wb, sheets, active, err := workbook.FromFile(path, sheet)
	if err != nil {
		return Result{}, err
	}
	ctx.State.LoadWorkbook(wb, path, sheets, active)
	msg := fmt.Sprintf("loaded %s [%s] (%d rows)\n", filepath.Base(wb.Name), wb.Sheet, len(wb.Cells))
	return Result{Message: msg}, nil
}

func (sampleAction) Name() string { return "sample" }

func (sampleAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	wb := workbook.SampleWorkbook()
	ctx.State.LoadWorkbook(wb, "", []string{wb.Sheet}, wb.ActiveCell)
	return Result{Message: "loaded sample workbook\n"}, nil
}
