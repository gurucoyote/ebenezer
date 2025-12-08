package actions

import (
	"errors"
	"fmt"
	"strings"
)

var (
	SearchForward        Action = searchAction{forward: true}
	SearchBackward       Action = searchAction{forward: false}
	SearchRepeatForward  Action = searchRepeatAction{forward: true}
	SearchRepeatBackward Action = searchRepeatAction{forward: false}
)

func init() {
	Register(SearchForward, "search")
	Register(SearchBackward, "search-reverse")
	Register(SearchRepeatForward, "search-next")
	Register(SearchRepeatBackward, "search-prev")
}

type searchAction struct {
	forward bool
}

type searchRepeatAction struct {
	forward bool
}

func (s searchAction) Name() string {
	if s.forward {
		return "search"
	}
	return "search-reverse"
}

func (s searchAction) Metadata() Metadata {
	desc := "Search forward for text in the current sheet"
	name := "search"
	if !s.forward {
		desc = "Search backward for text in the current sheet"
		name = "search-reverse"
	}
	return Metadata{
		Name:        name,
		Description: desc,
		Category:    "search",
		Args:        []Arg{{Name: "pattern", Description: "Text to locate", Variadic: true}},
	}
}

func (s searchAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	pattern := strings.Join(args, " ")
	if strings.TrimSpace(pattern) == "" {
		return Result{}, errors.New("search term required")
	}
	if err := ctx.State.Search(pattern, s.forward); err != nil {
		return Result{}, err
	}
	return Result{Message: fmt.Sprintf("found %s = %q\n", ctx.State.Address(), ctx.State.CurrentValue())}, nil
}

func (s searchRepeatAction) Name() string {
	if s.forward {
		return "search-next"
	}
	return "search-prev"
}

func (s searchRepeatAction) Metadata() Metadata {
	desc := "Repeat the last search in the same direction"
	name := "search-next"
	if !s.forward {
		desc = "Repeat the last search in the opposite direction"
		name = "search-prev"
	}
	return Metadata{
		Name:        name,
		Description: desc,
		Category:    "search",
		Idempotent:  true,
	}
}

func (s searchRepeatAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	if err := ctx.State.RepeatSearch(s.forward); err != nil {
		return Result{}, err
	}
	return Result{Message: fmt.Sprintf("found %s = %q\n", ctx.State.Address(), ctx.State.CurrentValue())}, nil
}
