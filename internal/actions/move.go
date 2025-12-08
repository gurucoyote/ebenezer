package actions

import (
	"fmt"
)

var Move Action = moveAction{}

func init() {
	Register(Move)
}

type moveAction struct{}

func (moveAction) Name() string { return "move" }

func (moveAction) Metadata() Metadata {
	return Metadata{
		Name:        "move",
		Description: "Move the cursor one cell in the specified direction",
		Category:    "navigation",
		Args: []Arg{
			{Name: "direction", Description: "one of left/right/up/down"},
		},
		Idempotent: true,
	}
}

func (moveAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	if len(args) != 1 {
		return Result{}, fmt.Errorf("move requires direction")
	}
	var deltaRow, deltaCol int
	switch args[0] {
	case "left":
		deltaCol = -1
	case "right":
		deltaCol = 1
	case "up":
		deltaRow = -1
	case "down":
		deltaRow = 1
	default:
		return Result{}, fmt.Errorf("unknown direction %s", args[0])
	}
	ctx.State.Move(deltaRow, deltaCol)
	msg := fmt.Sprintf("→ %s = %q\n", ctx.State.Address(), ctx.State.CurrentValue())
	return Result{Message: msg}, nil
}
