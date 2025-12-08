package actions

import (
	"fmt"
	"strings"
)

var (
	StyleDescribe Action = styleDescribeAction{}
	StyleCopy     Action = styleCopyAction{}
	StylePaste    Action = stylePasteAction{}
)

func init() {
	Register(StyleDescribe)
	Register(StyleCopy)
	Register(StylePaste)
}

type styleDescribeAction struct{}

type styleCopyAction struct{}

type stylePasteAction struct{}

func (styleDescribeAction) Name() string { return "style" }

func (styleDescribeAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	target := ctx.State.Address()
	if len(args) == 1 && strings.TrimSpace(args[0]) != "" {
		target = strings.ToUpper(args[0])
	}
	style, ok := ctx.State.StyleAt(target)
	if !ok {
		return Result{Message: fmt.Sprintf("%s: no style metadata available\n", target)}, nil
	}
	return Result{Message: fmt.Sprintf("%s: %s\n", target, style.Describe())}, nil
}

func (styleCopyAction) Name() string { return "style-copy" }

func (styleCopyAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	rangeStr := ""
	if len(args) == 1 {
		rangeStr = args[0]
	}
	if err := ctx.State.CopyStyle(rangeStr); err != nil {
		return Result{}, err
	}
	target := rangeStr
	if strings.TrimSpace(target) == "" {
		target = ctx.State.Address()
	}
	return Result{Message: fmt.Sprintf("style copied from %s\n", strings.ToUpper(target))}, nil
}

func (stylePasteAction) Name() string { return "style-paste" }

func (stylePasteAction) Exec(ctx Context, args []string) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	rangeStr := ""
	if len(args) == 1 {
		rangeStr = args[0]
	}
	if err := ctx.State.PasteStyle(rangeStr); err != nil {
		return Result{}, err
	}
	target := rangeStr
	if strings.TrimSpace(target) == "" {
		target = ctx.State.Address()
	}
	return Result{Message: fmt.Sprintf("style pasted into %s\n", strings.ToUpper(target))}, nil
}
