package actions

import (
	"bytes"

	"ebenezer/internal/ui/status"
)

var Status Action = statusAction{}

func init() {
	Register(Status)
}

type statusAction struct{}

func (statusAction) Name() string { return "status" }

func (statusAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	var buf bytes.Buffer
	status.Print(&buf, ctx.State)
	return Result{Message: buf.String()}, nil
}
