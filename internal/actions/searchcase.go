package actions

import (
	"errors"
	"fmt"
	"strings"
)

var SearchCase Action = searchCaseAction{}

func init() {
	Register(SearchCase)
}

type searchCaseAction struct{}

func (searchCaseAction) Name() string { return "search-case" }

func (searchCaseAction) Metadata() Metadata {
	return Metadata{
		Name:        "search-case",
		Description: "View or change search case-sensitivity",
		Category:    "search",
		Args:        []Arg{{Name: "mode", Description: "sensitive|insensitive|toggle", Optional: true}},
		Idempotent:  true,
	}
}

func (searchCaseAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	if len(args) == 0 {
		return Result{Message: formatSearchCase(ctx.State.SearchCaseSensitive)}, nil
	}
	mode := strings.ToLower(strings.TrimSpace(args[0]))
	switch mode {
	case "sensitive", "on", "true", "1":
		ctx.State.SetSearchCaseSensitivity(true)
	case "insensitive", "off", "false", "0":
		ctx.State.SetSearchCaseSensitivity(false)
	case "toggle":
		ctx.State.SetSearchCaseSensitivity(!ctx.State.SearchCaseSensitive)
	default:
		return Result{}, errors.New("expected sensitive|insensitive|toggle")
	}
	return Result{Message: formatSearchCase(ctx.State.SearchCaseSensitive)}, nil
}

func formatSearchCase(enabled bool) string {
	mode := "insensitive"
	if enabled {
		mode = "sensitive"
	}
	return fmt.Sprintf("search is %s\n", mode)
}
