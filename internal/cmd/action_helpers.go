package cmd

import (
	"fmt"

	"ebenezer/internal/actions"
	"github.com/spf13/cobra"
)

func executeAction(cmd *cobra.Command, action actions.Action, args []string) (actions.Result, error) {
	ctx := actions.NewContext(appState, cmd.OutOrStdout())
	res, err := action.Exec(ctx, args)
	if err != nil {
		return actions.Result{}, err
	}
	if res.Message != "" {
		fmt.Fprint(cmd.OutOrStdout(), res.Message)
	}
	return res, nil
}
