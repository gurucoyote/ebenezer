package actions

import (
	"fmt"
	"io"

	"ebenezer/internal/app"
)

// Context supplies state and optional writer targets to actions.
type Context struct {
	State  *app.State
	Writer io.Writer
}

// Result captures action output in a transport-agnostic way.
type Result struct {
	Message string
	Data    any
}

// Action defines the minimal interface all actions must satisfy.
type Action interface {
	Name() string
	Metadata() Metadata
	Exec(ctx Context, args []string) (Result, error)
}

// EnsureState returns an error if the context is missing the app state.
func EnsureState(ctx Context) error {
	if ctx.State == nil {
		return fmt.Errorf("action requires app state")
	}
	return nil
}

// NewContext is a helper for adapters to construct a Context.
func NewContext(state *app.State, w io.Writer) Context {
	return Context{State: state, Writer: w}
}
