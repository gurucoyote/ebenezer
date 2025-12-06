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
	fmt.Fprintf(w, "%s | %s\n", state.Address(), state.CurrentValue())
}
