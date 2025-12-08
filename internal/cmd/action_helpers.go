package cmd

import (
	"fmt"
	"time"

	"ebenezer/internal/actions"
	"github.com/spf13/cobra"
)

func executeAction(cmd *cobra.Command, action actions.Action, args []string) (actions.Result, error) {
	ctx := actions.NewContext(appState, cmd.OutOrStdout())
	logger := actions.JSONLogger{Writer: cmd.ErrOrStderr()}
	ctx.Logger = logger
	meta := action.Metadata()
	ctx.Logger.Before(ctx, meta, args)
	start := time.Now()
	res, err := action.Exec(ctx, args)
	ctx.Logger.After(ctx, meta, args, res, err, time.Since(start))
	if err != nil {
		return actions.Result{}, err
	}
	if res.Message != "" {
		fmt.Fprint(cmd.OutOrStdout(), res.Message)
	}
	return res, nil
}
