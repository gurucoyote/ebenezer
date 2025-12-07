package cmd

import (
	"errors"
	"fmt"
	"strings"

	"ebenezer/internal/workbook"
	"github.com/spf13/cobra"
)

var sheetSelectCmd = &cobra.Command{
	Use:   "ps [sheet]",
	Short: "List or switch sheets in the current workbook",
	RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) == 0 {
			return listSheets(cmd)
		}
		return switchSheet(cmd, args[0])
	},
}

func init() {
	rootCmd.AddCommand(sheetSelectCmd)
}

func listSheets(cmd *cobra.Command) error {
	if len(appState.SheetNames) == 0 {
		fmt.Fprintln(cmd.OutOrStdout(), "no sheet metadata available")
		return nil
	}
	current := ""
	if appState.Workbook != nil {
		current = appState.Workbook.Sheet
	}
	for _, name := range appState.SheetNames {
		marker := " "
		if strings.EqualFold(name, current) {
			marker = "*"
		}
		fmt.Fprintf(cmd.OutOrStdout(), "%s %s\n", marker, name)
	}
	return nil
}

func switchSheet(cmd *cobra.Command, sheet string) error {
	if appState.SourcePath == "" {
		return errors.New("current workbook not backed by a file; pass a filename to switch sheets")
	}
	wb, sheets, active, err := workbook.FromFile(appState.SourcePath, sheet)
	if err != nil {
		return err
	}
	appState.LoadWorkbook(wb, appState.SourcePath, sheets, active)
	fmt.Fprintf(cmd.OutOrStdout(), "switched to sheet %s\n", wb.Sheet)
	return nil
}
