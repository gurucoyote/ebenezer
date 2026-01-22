package cmd

import (
	"context"
	"fmt"
	"strings"

	"ebenezer/internal/app"
	"ebenezer/internal/ui/status"
	"ebenezer/internal/workbook"

	"github.com/spf13/cobra"
)

var (
	rootCmd = &cobra.Command{
		Use:     "ebenezer [FILE]",
		Short:   "Ebenezer spreadsheet CLI",
		Long:    "Ebenezer is a headless-yet-interactive spreadsheet editor. Run with a file to open keyboard mode, or use subcommands for automated workflows.",
		Example: "  ebenezer data.xlsx\n  ebenezer --sheet Budget data.xlsx\n  ebenezer -q richtext-scan data.xlsx\n  ebenezer mcp serve",
		Args:    cobra.MaximumNArgs(1),
	}
	appState         = app.NewState()
	rootSheet        string
	rootCSVDelimiter string
	rootQuiet        bool
)

func init() {
	rootCmd.RunE = rootRun
	rootCmd.PersistentPreRunE = rootPreRun
	rootCmd.PersistentFlags().StringVar(&rootSheet, "sheet", "", "Sheet to load when opening .xlsx files")
	rootCmd.PersistentFlags().StringVar(&rootCSVDelimiter, "delimiter", string(workbook.DefaultCSVDelimiter), "Delimiter used when reading/writing CSV files")
	rootCmd.PersistentFlags().BoolVarP(&rootQuiet, "quiet", "q", false, "Suppress JSON action logs")
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
	status.Print(cmd.OutOrStdout(), appState)
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
	wb, sheets, active, err := workbook.FromFile(path, sheet, workbook.WithCSVDelimiter(appState.CSVDelimiter()))
	if err != nil {
		return err
	}
	appState.LoadWorkbook(wb, path, sheets, active)
	fmt.Fprintf(cmd.OutOrStdout(), "loaded %s [%s]\n", wb.Name, wb.Sheet)
	return nil
}

func rootPreRun(cmd *cobra.Command, args []string) error {
	delimiter, err := parseDelimiterFlag(rootCSVDelimiter)
	if err != nil {
		return fmt.Errorf("invalid --delimiter value %q: %w", rootCSVDelimiter, err)
	}
	appState.SetCSVDelimiter(delimiter)
	return nil
}

func parseDelimiterFlag(value string) (rune, error) {
	trimmed := strings.TrimSpace(value)
	if trimmed == "" {
		return workbook.DefaultCSVDelimiter, nil
	}
	if trimmed == `\t` {
		return '\t', nil
	}
	runes := []rune(trimmed)
	if len(runes) != 1 {
		return 0, fmt.Errorf("delimiter must be a single character")
	}
	return runes[0], nil
}
