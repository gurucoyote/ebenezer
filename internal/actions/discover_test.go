package actions

import (
	"strings"
	"testing"
)

func TestDiscoverReturnsDeterministicMetadata(t *testing.T) {
	want := ListMetadata()
	got := Discover()
	if len(got) != len(want) {
		t.Fatalf("discover len mismatch: got %d want %d", len(got), len(want))
	}

	// Ensure sorted order (case-insensitive) to make MCP responses stable.
	for i := 1; i < len(got); i++ {
		if strings.ToLower(got[i-1].Name) > strings.ToLower(got[i].Name) {
			t.Fatalf("discover results not sorted: %q before %q", got[i-1].Name, got[i].Name)
		}
	}

	metaByName := map[string]Metadata{}
	for _, meta := range want {
		metaByName[meta.Name] = meta
	}
	for _, dto := range got {
		meta, ok := metaByName[dto.Name]
		if !ok {
			t.Fatalf("discover returned unknown action %q", dto.Name)
		}
		if dto.Description != meta.Description {
			t.Fatalf("description mismatch for %s", dto.Name)
		}
		if dto.Category != meta.Category {
			t.Fatalf("category mismatch for %s", dto.Name)
		}
		if dto.Idempotent != meta.Idempotent {
			t.Fatalf("idempotent mismatch for %s", dto.Name)
		}
		if dto.Experimental != meta.Experimental {
			t.Fatalf("experimental mismatch for %s", dto.Name)
		}
		if len(dto.Args) != len(meta.Args) {
			t.Fatalf("args length mismatch for %s: got %d want %d", dto.Name, len(dto.Args), len(meta.Args))
		}
		for i, arg := range dto.Args {
			if arg.Name != meta.Args[i].Name {
				t.Fatalf("arg %d name mismatch for %s", i, dto.Name)
			}
			if arg.Description != meta.Args[i].Description {
				t.Fatalf("arg %d description mismatch for %s", i, dto.Name)
			}
			if arg.Optional != meta.Args[i].Optional {
				t.Fatalf("arg %d optional mismatch for %s", i, dto.Name)
			}
			if arg.Variadic != meta.Args[i].Variadic {
				t.Fatalf("arg %d variadic mismatch for %s", i, dto.Name)
			}
		}
	}
}
