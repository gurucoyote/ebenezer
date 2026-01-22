package cmd

import "github.com/spf13/cobra"

func markInteractive(cmds ...*cobra.Command) {
	for _, cmd := range cmds {
		if cmd == nil {
			continue
		}
		if cmd.Annotations == nil {
			cmd.Annotations = map[string]string{}
		}
		cmd.Annotations[interactiveAnnotation] = "true"
	}
}
