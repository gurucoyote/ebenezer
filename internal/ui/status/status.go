package status

import (
	"fmt"
	"io"

	"ebenezer/internal/app"
)

// Print writes a single-line summary of the current cursor/value.
func Print(w io.Writer, state *app.State) {
	if state == nil || w == nil {
		return
	}
	message := fmt.Sprintf("%s | %s", state.Address(), state.CurrentValue())
	if summary := state.SelectionSummary(); summary != "" {
		message = fmt.Sprintf("%s | VISUAL %s", message, summary)
	}
	fmt.Fprintln(w, message)
}
