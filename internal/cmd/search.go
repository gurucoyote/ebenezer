package cmd

import (
	"ebenezer/internal/actions"
	"github.com/spf13/cobra"
)

var searchCmd = &cobra.Command{
	Use:   "search [pattern]",
	Short: "Search for text within the current sheet",
	Args:  cobra.MinimumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		reverse, _ := cmd.Flags().GetBool("reverse")
		action := actions.SearchForward
		if reverse {
			action = actions.SearchBackward
		}
		_, err := executeAction(cmd, action, args)
		return err
	},
}

var searchNextCmd = &cobra.Command{
	Use:    "search-next",
	Short:  "Repeat the last search direction (n)",
	Hidden: true,
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.SearchRepeatForward, nil)
		return err
	},
}

var searchPrevCmd = &cobra.Command{
	Use:    "search-prev",
	Short:  "Repeat the last search in the opposite direction (N)",
	Hidden: true,
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.SearchRepeatBackward, nil)
		return err
	},
}

func init() {
	searchCmd.Flags().Bool("reverse", false, "search backward (same as '?')")
	rootCmd.AddCommand(searchCmd)
	rootCmd.AddCommand(searchNextCmd)
	rootCmd.AddCommand(searchPrevCmd)
	markInteractive(searchCmd, searchNextCmd, searchPrevCmd)
}
