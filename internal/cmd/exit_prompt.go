package cmd

import (
	"bufio"
	"fmt"
	"io"
	"os"
	"strings"

	"ebenezer/internal/actions"
	"github.com/spf13/cobra"
)

type exitChoice int

const (
	exitCancel exitChoice = iota
	exitSave
	exitDiscard
)

func promptExitChoice(in io.Reader, out io.Writer) (exitChoice, error) {
	reader := bufio.NewReader(in)
	fmt.Fprint(out, "\nUnsaved changes. Save before exiting? (y=save / n=discard / c=cancel): ")
	for {
		line, err := reader.ReadString('\n')
		if err != nil && err != io.EOF {
			return exitCancel, err
		}
		trimmed := strings.TrimSpace(strings.ToLower(line))
		switch trimmed {
		case "y", "yes":
			return exitSave, nil
		case "n", "no", "d", "discard":
			return exitDiscard, nil
		case "", "c", "cancel":
			return exitCancel, nil
		default:
			fmt.Fprint(out, "Please enter y (save), n (discard), or c (cancel): ")
		}
		if err == io.EOF {
			return exitCancel, nil
		}
	}
}

func promptSavePath(in io.Reader, out io.Writer) (string, error) {
	reader := bufio.NewReader(in)
	fmt.Fprint(out, "Save as (leave blank to cancel): ")
	line, err := reader.ReadString('\n')
	if err != nil && err != io.EOF {
		return "", err
	}
	path := strings.TrimSpace(line)
	return path, nil
}

func saveOnExit(c *cobra.Command) error {
	if appState.SourcePath != "" {
		_, err := executeAction(c, actions.Save, nil)
		return err
	}
	path, err := promptSavePath(os.Stdin, c.OutOrStdout())
	if err != nil {
		return err
	}
	if strings.TrimSpace(path) == "" {
		return errPromptCanceled
	}
	_, err = executeAction(c, actions.Save, []string{path})
	return err
}
