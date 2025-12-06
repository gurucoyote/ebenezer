package app

// State captures the high-level CLI session data.
type State struct {
	WorkbookPath string
	ActiveSheet  string
}

// NewState initializes a default state used by early wiring.
func NewState() *State {
	return &State{
		WorkbookPath: "untitled.xlsx",
		ActiveSheet:  "Sheet1",
	}
}
