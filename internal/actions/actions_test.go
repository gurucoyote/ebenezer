package actions

import (
	"path/filepath"
	"strings"
	"testing"

	"ebenezer/internal/app"
)

func TestMoveAction(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)
	if _, err := Move.Exec(ctx, []string{"down"}); err != nil {
		t.Fatalf("move down: %v", err)
	}
	if st.Cursor.Row != 2 {
		t.Fatalf("expected row 2, got %d", st.Cursor.Row)
	}
	if _, err := Move.Exec(ctx, []string{"left"}); err != nil {
		t.Fatalf("move left: %v", err)
	}
	if st.Cursor.Col != 1 {
		t.Fatalf("expected col clamp to 1, got %d", st.Cursor.Col)
	}
}

func TestGotoAction(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)
	res, err := Goto.Exec(ctx, []string{"B2"})
	if err != nil {
		t.Fatalf("goto: %v", err)
	}
	if st.Address() != "B2" {
		t.Fatalf("expected cursor at B2, got %s", st.Address())
	}
	if !strings.Contains(res.Message, "B2") {
		t.Fatalf("expected message to mention B2, got %q", res.Message)
	}
}

func TestStatusAction(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)
	res, err := Status.Exec(ctx, nil)
	if err != nil {
		t.Fatalf("status: %v", err)
	}
	if res.Message == "" {
		t.Fatalf("expected status message")
	}
}

func TestClearAction(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)
	_, err := Edit.Exec(ctx, []string{"hello"})
	if err != nil {
		t.Fatalf("edit: %v", err)
	}
	if _, err := Clear.Exec(ctx, nil); err != nil {
		t.Fatalf("clear: %v", err)
	}
	if got := st.CurrentValue(); got != "" {
		t.Fatalf("expected cell cleared, got %q", got)
	}
}

func TestCutAndPasteActions(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)
	_, err := Edit.Exec(ctx, []string{"temp"})
	if err != nil {
		t.Fatalf("edit: %v", err)
	}
	if _, err := Cut.Exec(ctx, nil); err != nil {
		t.Fatalf("cut: %v", err)
	}
	if st.CurrentValue() != "" {
		t.Fatalf("expected source cleared after cut")
	}
	if err := st.Goto("B1"); err != nil {
		t.Fatalf("goto B1: %v", err)
	}
	if _, err := Paste.Exec(ctx, nil); err != nil {
		t.Fatalf("paste: %v", err)
	}
	if got := st.CurrentValue(); got != "temp" {
		t.Fatalf("expected paste result 'temp', got %q", got)
	}
}
func TestStyleCopyPasteActions(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)
	if _, err := StyleCopy.Exec(ctx, nil); err != nil {
		t.Fatalf("style copy: %v", err)
	}
	if _, err := StylePaste.Exec(ctx, []string{"A1"}); err != nil {
		t.Fatalf("style paste: %v", err)
	}
}

func TestSaveAction(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)
	tmp := t.TempDir() + "/out.xlsx"
	if _, err := Save.Exec(ctx, []string{tmp}); err != nil {
		t.Fatalf("save: %v", err)
	}
}

func TestSearchActions(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)
	if _, err := SearchForward.Exec(ctx, []string{"Item"}); err != nil {
		t.Fatalf("search: %v", err)
	}
	if _, err := SearchRepeatForward.Exec(ctx, nil); err != nil {
		t.Fatalf("repeat search: %v", err)
	}
	if _, err := SearchBackward.Exec(ctx, []string{"Item"}); err != nil {
		t.Fatalf("reverse search: %v", err)
	}
}

func TestSampleAndOpenActions(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)
	_, err := SampleData.Exec(ctx, nil)
	if err != nil {
		t.Fatalf("sample: %v", err)
	}
	tmp := filepath.Join(t.TempDir(), "test.csv")
	if err := st.Workbook.Save(tmp); err != nil {
		t.Fatalf("save temp csv: %v", err)
	}
	_, err = OpenFile.Exec(ctx, []string{tmp})
	if err != nil {
		t.Fatalf("open: %v", err)
	}
	if ctx.State.SourcePath != tmp {
		t.Fatalf("expected source path %s, got %s", tmp, ctx.State.SourcePath)
	}
}

func TestSheetListAction(t *testing.T) {
	st := app.NewState()
	st.SheetNames = []string{"Sheet1", "Sheet2"}
	st.Workbook.Sheet = "Sheet2"
	ctx := NewContext(st, nil)
	res, err := SheetList.Exec(ctx, nil)
	if err != nil {
		t.Fatalf("sheet list: %v", err)
	}
	if !strings.Contains(res.Message, "* Sheet2") {
		t.Fatalf("expected current sheet marker, got %q", res.Message)
	}
}

func TestSearchCaseAction(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)
	if _, err := SearchCase.Exec(ctx, nil); err != nil {
		t.Fatalf("search-case show: %v", err)
	}
	if st.SearchCaseSensitive {
		t.Fatalf("expected default insensitive")
	}
	if _, err := SearchCase.Exec(ctx, []string{"sensitive"}); err != nil {
		t.Fatalf("search-case set: %v", err)
	}
	if !st.SearchCaseSensitive {
		t.Fatalf("expected sensitive state")
	}
}

func TestInfoAction(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)
	res, err := Info.Exec(ctx, nil)
	if err != nil {
		t.Fatalf("info current: %v", err)
	}
	if !strings.Contains(res.Message, "sheet: Sheet1") {
		t.Fatalf("expected sheet in info output, got %q", res.Message)
	}
}

func TestColumnActions(t *testing.T) {
	st := app.NewState()
	ctx := NewContext(st, nil)
	if _, err := ColumnInsertRight.Exec(ctx, nil); err != nil {
		t.Fatalf("column insert right: %v", err)
	}
	if _, err := ColumnInsertLeft.Exec(ctx, nil); err != nil {
		t.Fatalf("column insert left: %v", err)
	}
	if _, err := ColumnDelete.Exec(ctx, nil); err != nil {
		t.Fatalf("column delete: %v", err)
	}
}
