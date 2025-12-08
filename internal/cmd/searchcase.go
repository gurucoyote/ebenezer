package cmd

import (
	"github.com/spf13/cobra"

	"ebenezer/internal/actions"
)

var searchCaseCmd = &cobra.Command{
	Use:   "search-case [sensitive|insensitive|toggle]",
	Short: "View or change the search case-sensitivity setting",
	Args:  cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.SearchCase, args)
		return err
	},
}

func init() {
	rootCmd.AddCommand(searchCaseCmd)
}
