package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

var moveCmd = &cobra.Command{
	Use:   "move [direction]",
	Short: "Move the demo cursor",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		switch args[0] {
		case "left":
			appState.Move(0, -1)
		case "right":
			appState.Move(0, 1)
		case "up":
			appState.Move(-1, 0)
		case "down":
			appState.Move(1, 0)
		default:
			return fmt.Errorf("unknown direction %s", args[0])
		}
		fmt.Fprintf(cmd.OutOrStdout(), "→ %s = %q\n", appState.Address(), appState.CurrentValue())
		return nil
	},
}

func init() {
	rootCmd.AddCommand(moveCmd)
}
