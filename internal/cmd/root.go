package cmd

import (
	"context"
	"fmt"

	"ebenezer/internal/app"
	"ebenezer/internal/ui/status"
	"ebenezer/internal/workbook"

	"github.com/spf13/cobra"
)

var (
	rootCmd = &cobra.Command{
		Use:   "ebenezer [FILE]",
		Short: "Ebenezer spreadsheet CLI (skeleton)",
		Args:  cobra.MaximumNArgs(1),
	}
	appState  = app.NewState()
	rootSheet string
)

func init() {
	rootCmd.RunE = rootRun
	rootCmd.PersistentPreRun = func(cmd *cobra.Command, args []string) {
		status.Print(cmd.OutOrStdout(), appState)
	}
	rootCmd.PersistentFlags().StringVar(&rootSheet, "sheet", "", "Sheet to load when opening .xlsx files")
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

func rootRun(cmd *cobra.Command, args []string) error {
	if len(args) > 0 {
		if err := loadWorkbookFromArg(cmd, args[0], rootSheet); err != nil {
			return err
		}
	} else {
		fmt.Fprintln(cmd.OutOrStdout(), "no file provided, using sample workbook")
		wb := workbook.SampleWorkbook()
		appState.LoadWorkbook(wb, "", []string{wb.Sheet}, "")
	}
	return runKeyboardMode(cmd)
}

func loadWorkbookFromArg(cmd *cobra.Command, path, sheet string) error {
	wb, sheets, active, err := workbook.FromFile(path, sheet)
	if err != nil {
		return err
	}
	appState.LoadWorkbook(wb, path, sheets, active)
	fmt.Fprintf(cmd.OutOrStdout(), "loaded %s [%s]\n", wb.Name, wb.Sheet)
	return nil
}
