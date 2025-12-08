package cmd

import (
	"github.com/spf13/cobra"

	"ebenezer/internal/actions"
)

var moveCmd = &cobra.Command{
	Use:   "move [direction]",
	Short: "Move the demo cursor",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.Move, args)
		return err
	},
}

func init() {
	rootCmd.AddCommand(moveCmd)
}
