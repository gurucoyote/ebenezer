package actions

import "fmt"

var (
	RowYank        Action = rowYankAction{}
	RowCut         Action = rowCutAction{}
	RowDelete      Action = rowDeleteAction{}
	RowInsertAbove Action = rowInsertAboveAction{}
	RowInsertBelow Action = rowInsertBelowAction{}
)

func init() {
	Register(RowYank)
	Register(RowCut)
	Register(RowDelete)
	Register(RowInsertAbove)
	Register(RowInsertBelow)
}

type rowYankAction struct{}

type rowCutAction struct{}

type rowDeleteAction struct{}

type rowInsertAboveAction struct{}

type rowInsertBelowAction struct{}

func (rowYankAction) Name() string { return "row-yank" }

func (rowYankAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	ctx.State.YankCurrentRow()
	return Result{Message: fmt.Sprintf("yanked row %d\n", ctx.State.Cursor.Row)}, nil
}

func (rowCutAction) Name() string { return "row-cut" }

func (rowCutAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	ctx.State.CutCurrentRow()
	return Result{Message: fmt.Sprintf("cut row %d\n", ctx.State.Cursor.Row)}, nil
}

func (rowDeleteAction) Name() string { return "row-delete" }

func (rowDeleteAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	ctx.State.DeleteCurrentRow()
	return Result{Message: "row deleted\n"}, nil
}

func (rowInsertAboveAction) Name() string { return "row-insert-above" }

func (rowInsertAboveAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	ctx.State.InsertRowAbove()
	return Result{Message: fmt.Sprintf("inserted row above %d\n", ctx.State.Cursor.Row)}, nil
}

func (rowInsertBelowAction) Name() string { return "row-insert-below" }

func (rowInsertBelowAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	ctx.State.InsertRowBelow()
	return Result{Message: fmt.Sprintf("inserted row below %d\n", ctx.State.Cursor.Row)}, nil
}
