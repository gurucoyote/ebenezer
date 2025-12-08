package actions

import (
	"fmt"
	"strings"
)

var Goto Action = gotoAction{}

var ColumnHeader Action = columnHeaderAction{}
var RowHeader Action = rowHeaderAction{}

func init() {
	Register(Goto)
	Register(ColumnHeader)
	Register(RowHeader)
}

type gotoAction struct{}

type columnHeaderAction struct{}

type rowHeaderAction struct{}

func (gotoAction) Name() string { return "goto" }

func (gotoAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	if len(args) != 1 {
		return Result{}, fmt.Errorf("goto requires a cell address")
	}
	target := strings.TrimSpace(args[0])
	if err := ctx.State.Goto(target); err != nil {
		return Result{}, err
	}
	msg := fmt.Sprintf("→ %s = %q\n", strings.ToUpper(target), ctx.State.CurrentValue())
	return Result{Message: msg}, nil
}

func (columnHeaderAction) Name() string { return "colheader" }

func (columnHeaderAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	msg := fmt.Sprintf("column %d header: %q\n", ctx.State.Cursor.Col, ctx.State.ColumnHeader())
	return Result{Message: msg}, nil
}

func (rowHeaderAction) Name() string { return "rowheader" }

func (rowHeaderAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	msg := fmt.Sprintf("row %d header: %q\n", ctx.State.Cursor.Row, ctx.State.RowHeader())
	return Result{Message: msg}, nil
}
