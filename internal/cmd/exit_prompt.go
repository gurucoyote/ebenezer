package cmd

import (
	"bufio"
	"errors"
	"fmt"
	"io"
	"os"
	"strings"

	"ebenezer/internal/actions"
	"ebenezer/internal/ui/keyboard"
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

func registerQuitCommands(c *cobra.Command) {
	if rootCmd == nil {
		return
	}
	if rootCmd.Commands() != nil {
		for _, child := range rootCmd.Commands() {
			switch child.Name() {
			case "q", "q!", "wq":
				return
			}
		}
	}
	rootCmd.AddCommand(quitCmd(c), quitForceCmd(c), writeQuitCmd(c))
}

func unregisterQuitCommands() {
	removeCommand(rootCmd, "q", "q!", "wq")
}

func removeCommand(cmd *cobra.Command, names ...string) {
	if cmd == nil {
		return
	}
	remove := map[string]struct{}{}
	for _, name := range names {
		remove[name] = struct{}{}
	}
	for _, child := range cmd.Commands() {
		if _, ok := remove[child.Name()]; ok {
			cmd.RemoveCommand(child)
		}
	}
}

func quitCmd(c *cobra.Command) *cobra.Command {
	return &cobra.Command{
		Use:    "q",
		Short:  "Quit (refuses if unsaved changes exist)",
		Hidden: true,
		RunE: func(cmd *cobra.Command, args []string) error {
			if appState.IsDirty() {
				fmt.Fprintln(cmd.OutOrStdout(), "No write since last change (add ! to override)")
				return nil
			}
			return keyboard.ErrQuit
		},
	}
}

func quitForceCmd(c *cobra.Command) *cobra.Command {
	return &cobra.Command{
		Use:    "q!",
		Short:  "Quit without saving",
		Hidden: true,
		RunE: func(cmd *cobra.Command, args []string) error {
			return keyboard.ErrQuit
		},
	}
}

func writeQuitCmd(c *cobra.Command) *cobra.Command {
	return &cobra.Command{
		Use:    "wq",
		Short:  "Save and quit",
		Hidden: true,
		RunE: func(cmd *cobra.Command, args []string) error {
			if err := saveOnExit(cmd); err != nil {
				if errors.Is(err, errPromptCanceled) {
					return nil
				}
				return err
			}
			return keyboard.ErrQuit
		},
	}
}
