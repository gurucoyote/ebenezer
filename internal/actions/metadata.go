package actions

// Metadata captures discovery information about an action so other frontends
// (CLI, MCP, automation) can understand capabilities without inspecting the
// implementation directly.
type Metadata struct {
	Name         string
	Description  string
	Category     string
	Args         []Arg
	Idempotent   bool
	Experimental bool
}

// Arg describes an action argument for discovery/documentation purposes.
type Arg struct {
	Name        string
	Description string
	Optional    bool
	Variadic    bool
}
