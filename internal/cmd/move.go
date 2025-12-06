package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

var moveCmd = &cobra.Command{
	Use:   "move [direction]",
	Short: "Demo movement command invoked by keyboard bindings",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		fmt.Fprintf(cmd.OutOrStdout(), "moved %s\n", args[0])
		return nil
	},
}

func init() {
	rootCmd.AddCommand(moveCmd)
}
