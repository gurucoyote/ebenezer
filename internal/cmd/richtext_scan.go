package cmd

import (
	"ebenezer/internal/actions"
	"github.com/spf13/cobra"
)

var richTextScanCmd = &cobra.Command{
	Use:   "richtext-scan [file] [--sheet NAME]",
	Short: "Scan an XLSX file for inline rich text runs",
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.RichTextScan, args)
		return err
	},
}

func init() {
	rootCmd.AddCommand(richTextScanCmd)
}
