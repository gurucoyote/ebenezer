package cmd

import (
	"fmt"

	"ebenezer/internal/workbook"
	"github.com/spf13/cobra"
)

var openCmd = &cobra.Command{
	Use:   "open [path]",
	Short: "Open a workbook (.csv or .xlsx)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		wb, err := workbook.FromFile(args[0], openSheet)
		if err != nil {
			return err
		}
		appState.LoadWorkbook(wb)
		fmt.Fprintf(cmd.OutOrStdout(), "loaded %s [%s] (%d rows)\n", wb.Name, wb.Sheet, len(wb.Cells))
		return nil
	},
}

var openSheet string

func init() {
	rootCmd.AddCommand(openCmd)
	openCmd.Flags().StringVar(&openSheet, "sheet", "", "Sheet name when opening .xlsx files")
}
