package cmd

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

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
}

var saveCmd = &cobra.Command{
	Use:   "save [path]",
	Short: "Save the current workbook",
	Args:  cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		force, _ := cmd.Flags().GetBool("force")
		return runSave(cmd, args, force)
	},
}

var saveForceCmd = &cobra.Command{
	Use:    "save!",
	Short:  "Force save the current workbook",
	Hidden: true,
	Args:   cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		return runSave(cmd, args, true)
	},
}

var saveAsCmd = &cobra.Command{
	Use:   "saveas <path>",
	Short: "Save the current workbook to a new file",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		force, _ := cmd.Flags().GetBool("force")
		return runSave(cmd, args, force)
	},
}

var writeCmd = &cobra.Command{
	Use:    "w [path]",
	Short:  "Vim-style save command",
	Hidden: true,
	Args:   cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		return runSave(cmd, args, false)
	},
}

var writeForceCmd = &cobra.Command{
	Use:    "w! [path]",
	Short:  "Vim-style force save",
	Hidden: true,
	Args:   cobra.MaximumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		return runSave(cmd, args, true)
	},
}

func runSave(cmd *cobra.Command, args []string, force bool) error {
	if appState.Workbook == nil {
		return errors.New("no workbook loaded")
	}
	target := appState.SourcePath
	if len(args) > 0 {
		target = args[0]
	}
	if strings.TrimSpace(target) == "" {
		return errors.New("please provide a filename")
	}
	if !force && shouldBlockOverwrite(target) {
		return fmt.Errorf("%s exists (use :w! or --force)", target)
	}
	if err := appState.Save(target); err != nil {
		return err
	}
	fmt.Fprintf(cmd.OutOrStdout(), "saved %s\n", target)
	return nil
}

func shouldBlockOverwrite(path string) bool {
	if appState.SourcePath != "" {
		if sameFile(path, appState.SourcePath) {
			return false
		}
	}
	_, err := os.Stat(path)
	return err == nil
}

func sameFile(a, b string) bool {
	aAbs, err1 := filepath.Abs(a)
	bAbs, err2 := filepath.Abs(b)
	if err1 != nil || err2 != nil {
		return a == b
	}
	return aAbs == bAbs
}
