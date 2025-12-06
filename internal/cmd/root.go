package cmd

import (
	"context"

	"github.com/spf13/cobra"
)

var rootCmd = &cobra.Command{
	Use:   "ebenezer",
	Short: "Ebenezer spreadsheet CLI (skeleton)",
}

// Execute runs the root command with the provided arguments/context.
func Execute(ctx context.Context, args []string) error {
	if len(args) > 0 {
		rootCmd.SetArgs(args)
	}
	return rootCmd.ExecuteContext(ctx)
}

// Root exposes the root Cobra command for wiring from other packages.
func Root() *cobra.Command {
	return rootCmd
}
