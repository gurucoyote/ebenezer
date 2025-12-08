package actions

import (
	"fmt"
	"strings"
	"sync"
)

type registry struct {
	sync.RWMutex
	entries map[string]Action
}

var defaultRegistry = &registry{entries: map[string]Action{}}

// Register adds an action to the global registry and associates any aliases.
func Register(action Action, aliases ...string) {
	defaultRegistry.Register(action, aliases...)
}

func (r *registry) Register(action Action, aliases ...string) {
	if action == nil {
		panic("actions: cannot register nil action")
	}
	r.Lock()
	defer r.Unlock()
	names := append([]string{action.Name()}, aliases...)
	for _, name := range names {
		key := strings.ToLower(strings.TrimSpace(name))
		if key == "" {
			continue
		}
		r.entries[key] = action
	}
}

// Lookup retrieves an action by name (case-insensitive).
func Lookup(name string) (Action, bool) {
	return defaultRegistry.Lookup(name)
}

func (r *registry) Lookup(name string) (Action, bool) {
	r.RLock()
	defer r.RUnlock()
	a, ok := r.entries[strings.ToLower(strings.TrimSpace(name))]
	return a, ok
}

// MustLookup returns the action or panics if missing.
func MustLookup(name string) Action {
	if action, ok := Lookup(name); ok {
		return action
	}
	panic(fmt.Sprintf("actions: %s not registered", name))
}

// List returns all registered actions.
func List() []Action {
	defaultRegistry.RLock()
	defer defaultRegistry.RUnlock()
	result := make([]Action, 0, len(defaultRegistry.entries))
	seen := map[Action]struct{}{}
	for _, action := range defaultRegistry.entries {
		if _, ok := seen[action]; ok {
			continue
		}
		seen[action] = struct{}{}
		result = append(result, action)
	}
	return result
}
