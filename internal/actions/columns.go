package actions

import "fmt"

var (
	ColumnInsertLeft  Action = columnInsertLeftAction{}
	ColumnInsertRight Action = columnInsertRightAction{}
	ColumnDelete      Action = columnDeleteAction{}
)

func init() {
	Register(ColumnInsertLeft)
	Register(ColumnInsertRight)
	Register(ColumnDelete)
}

type columnInsertLeftAction struct{}
type columnInsertRightAction struct{}
type columnDeleteAction struct{}

func (columnInsertLeftAction) Name() string { return "column-insert-left" }

func (columnInsertLeftAction) Metadata() Metadata {
	return Metadata{
		Name:        "column-insert-left",
		Description: "Insert a blank column before the cursor",
		Category:    "editing",
	}
}

func (columnInsertLeftAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	ctx.State.InsertColumnLeft()
	return Result{Message: fmt.Sprintf("inserted column before %s\n", ctx.State.Address())}, nil
}

func (columnInsertRightAction) Name() string { return "column-insert-right" }

func (columnInsertRightAction) Metadata() Metadata {
	return Metadata{
		Name:        "column-insert-right",
		Description: "Insert a blank column after the cursor",
		Category:    "editing",
	}
}

func (columnInsertRightAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	ctx.State.InsertColumnRight()
	return Result{Message: fmt.Sprintf("inserted column after %s\n", ctx.State.Address())}, nil
}

func (columnDeleteAction) Name() string { return "column-delete" }

func (columnDeleteAction) Metadata() Metadata {
	return Metadata{
		Name:        "column-delete",
		Description: "Delete the column at the cursor",
		Category:    "editing",
	}
}

func (columnDeleteAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	ctx.State.DeleteCurrentColumn()
	return Result{Message: "column deleted\n"}, nil
}
