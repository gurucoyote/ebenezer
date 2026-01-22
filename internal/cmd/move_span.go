package cmd

import (
	"github.com/spf13/cobra"

	"ebenezer/internal/actions"
)

var moveSpanCmd = &cobra.Command{
	Use:   "move-span [direction]",
	Short: "Move the cursor to the edge of a filled span (Ctrl+Arrow behavior)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.MoveSpan, args)
		return err
	},
}

func init() {
	rootCmd.AddCommand(moveSpanCmd)
	markInteractive(moveSpanCmd)
}
