package cmd

import (
	"ebenezer/internal/actions"
	"github.com/spf13/cobra"
)

func init() {
	saveCmd.Flags().Bool("force", false, "overwrite existing files")
	saveAsCmd.Flags().Bool("force", false, "overwrite existing files")

	rootCmd.AddCommand(saveCmd)
	rootCmd.AddCommand(saveAsCmd)
	rootCmd.AddCommand(writeCmd)
	rootCmd.AddCommand(writeForceCmd)
	rootCmd.AddCommand(saveForceCmd)
	markInteractive(saveCmd, saveAsCmd, saveForceCmd, writeCmd, writeForceCmd)
}

var saveCmd = &cobra.Command{
	Use:   "save [path]",
	Short: "Save the current workbook",
	Args:  cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		force, _ := cmd.Flags().GetBool("force")
		extra := []string{}
		if force {
			extra = append(extra, "--force")
		}
		args = append(extra, args...)
		_, err := executeAction(cmd, actions.Save, args)
		return err
	},
}

var saveForceCmd = &cobra.Command{
	Use:    "save!",
	Short:  "Force save the current workbook",
	Hidden: true,
	Args:   cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.SaveForce, args)
		return err
	},
}

var saveAsCmd = &cobra.Command{
	Use:   "saveas <path>",
	Short: "Save the current workbook to a new file",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		force, _ := cmd.Flags().GetBool("force")
		extra := []string{}
		if force {
			extra = append(extra, "--force")
		}
		args = append(extra, args...)
		_, err := executeAction(cmd, actions.SaveAs, args)
		return err
	},
}

var writeCmd = &cobra.Command{
	Use:    "w [path]",
	Short:  "Vim-style save command",
	Hidden: true,
	Args:   cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.Write, args)
		return err
	},
}

var writeForceCmd = &cobra.Command{
	Use:    "w! [path]",
	Short:  "Vim-style force save",
	Hidden: true,
	Args:   cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		_, err := executeAction(cmd, actions.WriteForce, args)
		return err
	},
}
