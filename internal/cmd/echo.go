package cmd

import (
	"fmt"
	"strings"

	"github.com/spf13/cobra"
)

var echoCmd = &cobra.Command{
	Use:   "echo [text]",
	Short: "Echo text to stdout (placeholder command for testing)",
	Args:  cobra.MinimumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		fmt.Fprintln(cmd.OutOrStdout(), strings.Join(args, " "))
		return nil
	},
}

func init() {
	rootCmd.AddCommand(echoCmd)
}
