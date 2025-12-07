package cmd

import (
	"errors"
	"fmt"
	"path/filepath"
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

var sheetNewCmd = &cobra.Command{
	Use:   "ns <name> [copy-from]",
	Short: "Create a new sheet (optionally by copying an existing one)",
	Args:  cobra.MinimumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		return createSheet(cmd, args)
	},
}

func init() {
	rootCmd.AddCommand(sheetSelectCmd, sheetNewCmd)
}

func init() {
	sheetNewCmd.Flags().String("copy", "", "Copy contents from existing sheet")
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

func createSheet(cmd *cobra.Command, args []string) error {
	if appState.SourcePath == "" || strings.ToLower(filepath.Ext(appState.SourcePath)) != ".xlsx" {
		return errors.New("sheet creation requires an .xlsx file saved on disk")
	}
	name := args[0]
	copyFrom, _ := cmd.Flags().GetString("copy")
	if copyFrom == "" && len(args) > 1 {
		copyFrom = args[1]
	}
	if err := workbook.AddSheet(appState.SourcePath, name, copyFrom); err != nil {
		return err
	}
	wb, sheets, active, err := workbook.FromFile(appState.SourcePath, name)
	if err != nil {
		return err
	}
	appState.LoadWorkbook(wb, appState.SourcePath, sheets, active)
	fmt.Fprintf(cmd.OutOrStdout(), "created sheet %s\n", name)
	return nil
}
