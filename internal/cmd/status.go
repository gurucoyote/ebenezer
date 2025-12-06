package cmd

import (
	"github.com/spf13/cobra"
)

var statusCmd = &cobra.Command{
	Use:   "status",
	Short: "Print the current cursor/value",
	Run: func(cmd *cobra.Command, args []string) {
		cmd.Printf("%s = %q\n", appState.Address(), appState.CurrentValue())
	},
}

func init() {
	rootCmd.AddCommand(statusCmd)
}
