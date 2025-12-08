package actions

import "testing"

func TestAllActionsExposeMetadata(t *testing.T) {
	actions := List()
	if len(actions) == 0 {
		t.Fatal("expected registered actions")
	}
	for _, action := range actions {
		meta := action.Metadata()
		if meta.Name == "" {
			t.Fatalf("action %T missing metadata name", action)
		}
		if meta.Description == "" {
			t.Fatalf("action %s missing description", meta.Name)
		}
		if meta.Category == "" {
			t.Fatalf("action %s missing category", meta.Name)
		}
	}

	metas := ListMetadata()
	if len(metas) != len(actions) {
		t.Fatalf("expected metadata list to match actions: got %d vs %d", len(metas), len(actions))
	}
}
