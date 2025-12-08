package actions

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

var (
	Save       Action = saveAction{}
	SaveForce  Action = saveForceAction{}
	SaveAs     Action = saveAsAction{}
	Write      Action = writeAction{}
	WriteForce Action = writeForceAction{}
)

func init() {
	Register(Save)
	Register(SaveForce)
	Register(SaveAs)
	Register(Write)
	Register(WriteForce)
}

const (
	flagForce = "--force"
)

type saveAction struct{}

type saveForceAction struct{}

type saveAsAction struct{}

type writeAction struct{}

type writeForceAction struct{}

func (saveAction) Name() string { return "save" }

func (saveAction) Metadata() Metadata {
	return Metadata{
		Name:        "save",
		Description: "Save the current workbook to its existing path or an optional new path",
		Category:    "file",
		Args:        []Arg{{Name: "path", Description: "Optional path", Optional: true}},
	}
}

func (saveAction) Exec(ctx Context, args []string) (Result, error) {
	return runSave(ctx, args, false)
}

func (saveForceAction) Name() string { return "save!" }

func (saveForceAction) Metadata() Metadata {
	return Metadata{
		Name:        "save!",
		Description: "Force save the workbook, overwriting existing files",
		Category:    "file",
		Args:        []Arg{{Name: "path", Description: "Optional path", Optional: true}},
	}
}

func (saveForceAction) Exec(ctx Context, args []string) (Result, error) {
	return runSave(ctx, args, true)
}

func (saveAsAction) Name() string { return "saveas" }

func (saveAsAction) Metadata() Metadata {
	return Metadata{
		Name:        "saveas",
		Description: "Save the current workbook to a new path",
		Category:    "file",
		Args:        []Arg{{Name: "path", Description: "Destination path"}},
	}
}

func (saveAsAction) Exec(ctx Context, args []string) (Result, error) {
	if len(args) != 1 {
		return Result{}, errors.New("saveas requires a path")
	}
	return runSave(ctx, args, false)
}

func (writeAction) Name() string { return "w" }

func (writeAction) Metadata() Metadata {
	return Metadata{
		Name:        "w",
		Description: "Vim-style save command (alias of save)",
		Category:    "file",
		Args:        []Arg{{Name: "path", Description: "Optional path", Optional: true}},
	}
}

func (writeAction) Exec(ctx Context, args []string) (Result, error) {
	return runSave(ctx, args, false)
}

func (writeForceAction) Name() string { return "w!" }

func (writeForceAction) Metadata() Metadata {
	return Metadata{
		Name:        "w!",
		Description: "Vim-style forced save",
		Category:    "file",
		Args:        []Arg{{Name: "path", Description: "Optional path", Optional: true}},
	}
}

func (writeForceAction) Exec(ctx Context, args []string) (Result, error) {
	return runSave(ctx, args, true)
}

func runSave(ctx Context, args []string, force bool) (Result, error) {
	if err := EnsureState(ctx); err != nil {
		return Result{}, err
	}
	state := ctx.State
	if state.Workbook == nil {
		return Result{}, errors.New("no workbook loaded")
	}
	target := state.SourcePath
	forceFlag := force
	filteredArgs := make([]string, 0, len(args))
	for _, arg := range args {
		if arg == flagForce {
			forceFlag = true
			continue
		}
		filteredArgs = append(filteredArgs, arg)
	}
	if len(filteredArgs) > 0 {
		target = filteredArgs[0]
	}
	if strings.TrimSpace(target) == "" {
		return Result{}, errors.New("please provide a filename")
	}
	if !forceFlag && shouldBlockOverwrite(state.SourcePath, target) {
		return Result{}, fmt.Errorf("%s exists (use :w! or --force)", target)
	}
	if err := state.Save(target); err != nil {
		return Result{}, err
	}
	return Result{Message: fmt.Sprintf("saved %s\n", target)}, nil
}

func shouldBlockOverwrite(sourcePath, target string) bool {
	if strings.TrimSpace(sourcePath) != "" {
		if sameFile(sourcePath, target) {
			return false
		}
	}
	_, err := os.Stat(target)
	return err == nil
}

func sameFile(a, b string) bool {
	aAbs, err1 := filepath.Abs(a)
	bAbs, err2 := filepath.Abs(b)
	if err1 != nil || err2 != nil {
		return a == b
	}
	return aAbs == bAbs
}
