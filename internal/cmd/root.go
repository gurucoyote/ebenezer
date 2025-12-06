package cmd

import (
	"context"

	"ebenezer/internal/app"
	"ebenezer/internal/ui/status"

	"github.com/spf13/cobra"
)

var (
	rootCmd = &cobra.Command{
		Use:   "ebenezer",
		Short: "Ebenezer spreadsheet CLI (skeleton)",
	}
	appState = app.NewState()
)

func init() {
	rootCmd.PersistentPreRun = func(cmd *cobra.Command, args []string) {
		status.Print(cmd.OutOrStdout(), appState)
	}
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

// AppState exposes the shared in-memory state for commands/keyboard bindings.
func AppState() *app.State {
	return appState
}
