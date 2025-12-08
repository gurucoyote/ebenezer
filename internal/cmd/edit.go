package cmd

import (
	"ebenezer/internal/actions"
	"github.com/spf13/cobra"
)

var editCmd = &cobra.Command{
	Use:   "edit [value]",
	Short: "Set the current cell's value",
	Args:  cobra.MinimumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.Edit, args)
		return err
	},
}

var clearCmd = &cobra.Command{
	Use:   "clear",
	Short: "Clear the current cell",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.Clear, nil)
		return err
	},
}

var yankCmd = &cobra.Command{
	Use:   "yank",
	Short: "Copy the current cell into the clipboard",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.Yank, nil)
		return err
	},
}

var cutCmd = &cobra.Command{
	Use:   "cut",
	Short: "Cut the current cell into the clipboard",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.Cut, nil)
		return err
	},
}

var pasteBefore bool

var pasteCmd = &cobra.Command{
	Use:   "paste",
	Short: "Paste the clipboard into the current location",
	RunE: func(cmd *cobra.Command, args []string) error {
		pasteArgs := make([]string, len(args))
		copy(pasteArgs, args)
		if pasteBefore {
			pasteArgs = append(pasteArgs, "--before")
		}
		_, err := executeAction(cmd, actions.Paste, pasteArgs)
		return err
	},
}

func init() {
	rootCmd.AddCommand(editCmd)
	rootCmd.AddCommand(clearCmd)
	rootCmd.AddCommand(yankCmd)
	rootCmd.AddCommand(cutCmd)
	rootCmd.AddCommand(pasteCmd)
	pasteCmd.Flags().BoolVar(&pasteBefore, "before", false, "Paste rows above the current row when clipboard holds rows")
}
