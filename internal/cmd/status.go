package cmd

import (
	"github.com/spf13/cobra"

	"ebenezer/internal/actions"
)

var statusCmd = &cobra.Command{
	Use:   "status",
	Short: "Print the current cursor/value",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.Status, nil)
		return err
	},
}

func init() {
	rootCmd.AddCommand(statusCmd)
}
