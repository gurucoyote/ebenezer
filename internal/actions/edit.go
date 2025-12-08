package actions

import (
	"fmt"
	"strings"
)

var (
	Edit  Action = editAction{}
	Clear Action = clearAction{}
	Yank  Action = yankCellAction{}
	Cut   Action = cutCellAction{}
	Paste Action = pasteAction{}
)

func init() {
	Register(Edit)
	Register(Clear)
	Register(Yank)
	Register(Cut)
	Register(Paste)
}

type editAction struct{}

type clearAction struct{}

type yankCellAction struct{}

type cutCellAction struct{}

type pasteAction struct{}

func (editAction) Name() string { return "edit" }

func (editAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	value := strings.Join(args, " ")
	ctx.State.EditCurrentCell(value)
	msg := fmt.Sprintf("%s = %q\n", ctx.State.Address(), ctx.State.CurrentValue())
	return Result{Message: msg}, nil
}

func (clearAction) Name() string { return "clear" }

func (clearAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	summary := ctx.State.SelectionSummary()
	ctx.State.ClearCurrentCell()
	if summary != "" {
		return Result{Message: fmt.Sprintf("cleared %s\n", summary)}, nil
	}
	return Result{Message: fmt.Sprintf("%s cleared\n", ctx.State.Address())}, nil
}

func (yankCellAction) Name() string { return "yank" }

func (yankCellAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	summary := ctx.State.SelectionSummary()
	value := ctx.State.YankCurrentCell()
	if summary != "" {
		return Result{Message: fmt.Sprintf("yanked %s\n", summary)}, nil
	}
	return Result{Message: fmt.Sprintf("yanked %s = %q\n", ctx.State.Address(), value)}, nil
}

func (cutCellAction) Name() string { return "cut" }

func (cutCellAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	summary := ctx.State.SelectionSummary()
	value := ctx.State.CutCurrentCell()
	if summary != "" {
		return Result{Message: fmt.Sprintf("cut %s\n", summary)}, nil
	}
	return Result{Message: fmt.Sprintf("cut %s = %q\n", ctx.State.Address(), value)}, nil
}

func (pasteAction) Name() string { return "paste" }

func (pasteAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	before := false
	for _, arg := range args {
		if arg == "--before" || arg == "before" {
			before = true
		}
	}
	summary := ctx.State.SelectionSummary()
	if err := ctx.State.PasteClipboard(before); err != nil {
		return Result{}, err
	}
	if summary != "" {
		return Result{Message: fmt.Sprintf("pasted into %s\n", summary)}, nil
	}
	return Result{Message: fmt.Sprintf("pasted into %s\n", ctx.State.Address())}, nil
}
