package mcp

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"ebenezer/internal/workbook"
	"github.com/mark3labs/mcp-go/mcp"
)

func TestNormalizeOptionsDefaults(t *testing.T) {
	opts := normalizeOptions(Options{})
	if opts.Name != "ebenezer" {
		t.Fatalf("unexpected default name: %s", opts.Name)
	}
	if opts.Version == "" {
		t.Fatalf("expected version default")
	}
}

func TestWorkspaceLifecycleTools(t *testing.T) {
	handler := &toolHandler{sessions: NewSessionManager()}
	path := writeSampleWorkbook(t)

	openRes := mustCall(t, handler.workspaceOpenTool, map[string]any{"path": path})
	var openPayload workspaceOpenResponse
	decodeResult(t, openRes, &openPayload)
	if openPayload.SessionID == "" {
		t.Fatalf("expected session id in open response")
	}
	if len(openPayload.Sheets) == 0 {
		t.Fatalf("expected sheets list")
	}

	cursorRes := mustCall(t, handler.cursorGetTool, map[string]any{"session_id": openPayload.SessionID})
	var cursor cursorResponse
	decodeResult(t, cursorRes, &cursor)
	if cursor.Address != "A1" {
		t.Fatalf("expected default cursor A1, got %s", cursor.Address)
	}

	setRes := mustCall(t, handler.cursorSetTool, map[string]any{"session_id": openPayload.SessionID, "address": "B2"})
	decodeResult(t, setRes, &cursor)
	if cursor.Address != "B2" {
		t.Fatalf("expected cursor at B2, got %s", cursor.Address)
	}
	if cursor.Value == "" {
		t.Fatalf("expected value at B2")
	}

	infoRes := mustCall(t, handler.infoGetTool, map[string]any{"session_id": openPayload.SessionID, "details": true})
	var info infoGetResponse
	decodeResult(t, infoRes, &info)
	if info.Info.ActiveSheet == "" {
		t.Fatalf("expected active sheet")
	}

	savePath := filepath.Join(t.TempDir(), "saved.csv")
	saveRes := mustCall(t, handler.workbookSaveTool, map[string]any{
		"session_id": openPayload.SessionID,
		"path":       savePath,
	})
	var savePayload workbookSaveResponse
	decodeResult(t, saveRes, &savePayload)
	if _, err := os.Stat(savePayload.Path); err != nil {
		t.Fatalf("expected file saved: %v", err)
	}

	editRes := mustCall(t, handler.cellEditTool, map[string]any{
		"session_id": openPayload.SessionID,
		"address":    "C2",
		"value":      "99",
	})
	var editPayload cellEditResponse
	decodeResult(t, editRes, &editPayload)
	if editPayload.Value != "99" || editPayload.Address != "C2" {
		t.Fatalf("unexpected edit payload: %+v", editPayload)
	}

	clearRes := mustCall(t, handler.cellClearTool, map[string]any{
		"session_id": openPayload.SessionID,
		"address":    "C2",
	})
	var clearPayload cellEditResponse
	decodeResult(t, clearRes, &clearPayload)
	if clearPayload.Value != "" {
		t.Fatalf("expected cleared cell to be blank, got %q", clearPayload.Value)
	}

	rowInsert := mustCall(t, handler.rowInsertBelowTool, map[string]any{
		"session_id": openPayload.SessionID,
		"address":    "A2",
	})
	var rowResp rowOpResponse
	decodeResult(t, rowInsert, &rowResp)
	if rowResp.Operation != "insert_below" {
		t.Fatalf("unexpected row op response %+v", rowResp)
	}

	rowDelete := mustCall(t, handler.rowDeleteTool, map[string]any{
		"session_id": openPayload.SessionID,
		"address":    "A2",
	})
	decodeResult(t, rowDelete, &rowResp)
	if rowResp.Operation != "delete" {
		t.Fatalf("unexpected row delete response %+v", rowResp)
	}

	colInsert := mustCall(t, handler.columnInsertRightTool, map[string]any{
		"session_id": openPayload.SessionID,
		"address":    "A1",
	})
	var colResp columnOpResponse
	decodeResult(t, colInsert, &colResp)
	if colResp.Operation != "insert_right" {
		t.Fatalf("unexpected column insert response %+v", colResp)
	}

	colDelete := mustCall(t, handler.columnDeleteTool, map[string]any{
		"session_id": openPayload.SessionID,
	})
	decodeResult(t, colDelete, &colResp)
	if colResp.Operation != "delete" {
		t.Fatalf("unexpected column delete response %+v", colResp)
	}

	colWidthShow := mustCall(t, handler.columnWidthTool, map[string]any{
		"session_id": openPayload.SessionID,
		"mode":       "show",
	})
	var widthResp columnWidthResponse
	decodeResult(t, colWidthShow, &widthResp)
	if len(widthResp.Columns) == 0 {
		t.Fatalf("expected width metadata, got %+v", widthResp)
	}

	colWidthSet := mustCall(t, handler.columnWidthTool, map[string]any{
		"session_id": openPayload.SessionID,
		"mode":       "set",
		"columns":    "B",
		"width":      28,
	})
	decodeResult(t, colWidthSet, &widthResp)
	if len(widthResp.Columns) != 1 || widthResp.Columns[0].Width != 28 {
		t.Fatalf("expected explicit width set response, got %+v", widthResp)
	}

	colWidthAuto := mustCall(t, handler.columnWidthTool, map[string]any{
		"session_id": openPayload.SessionID,
		"mode":       "auto",
		"columns":    "A:C",
		"min_width":  12,
		"max_width":  40,
	})
	decodeResult(t, colWidthAuto, &widthResp)
	if len(widthResp.Columns) != 3 {
		t.Fatalf("expected auto width to touch three columns, got %+v", widthResp)
	}

	selSet := mustCall(t, handler.selectionSetTool, map[string]any{
		"session_id": openPayload.SessionID,
		"range":      "A1:B2",
	})
	var selResp selectionResponse
	decodeResult(t, selSet, &selResp)
	if !selResp.Active || selResp.Summary == "" {
		t.Fatalf("expected active selection, got %+v", selResp)
	}

	selClear := mustCall(t, handler.selectionClearTool, map[string]any{
		"session_id": openPayload.SessionID,
	})
	decodeResult(t, selClear, &selResp)
	if selResp.Active {
		t.Fatalf("expected selection cleared, got %+v", selResp)
	}

	selExport := mustCall(t, handler.selectionExportTool, map[string]any{
		"session_id": openPayload.SessionID,
		"range":      "A1:A1",
	})
	var exportResp selectionExportResponse
	decodeResult(t, selExport, &exportResp)
	if len(exportResp.Values) != 1 || len(exportResp.Values[0]) != 1 {
		t.Fatalf("unexpected export payload %+v", exportResp)
	}
	if exportResp.Rows != 1 || exportResp.Cols != 1 || exportResp.Range == "" {
		t.Fatalf("expected normalized range metadata, got %+v", exportResp)
	}
	if exportResp.SavedPath != "" {
		t.Fatalf("did not expect savedPath for in-memory export")
	}

	exportCSV := filepath.Join(t.TempDir(), "selection.csv")
	selExportCSV := mustCall(t, handler.selectionExportTool, map[string]any{
		"session_id": openPayload.SessionID,
		"range":      "A1:B2",
		"path":       exportCSV,
	})
	var exportCSVResp selectionExportResponse
	decodeResult(t, selExportCSV, &exportCSVResp)
	if exportCSVResp.SavedPath != exportCSV || exportCSVResp.Format != "csv" {
		t.Fatalf("expected csv export metadata, got %+v", exportCSVResp)
	}
	if exportCSVResp.Rows != 2 || exportCSVResp.Cols != 2 || exportCSVResp.Bytes == 0 {
		t.Fatalf("unexpected csv export dimensions %+v", exportCSVResp)
	}
	if _, err := os.Stat(exportCSV); err != nil {
		t.Fatalf("expected csv export file: %v", err)
	}

	exportXLSX := filepath.Join(t.TempDir(), "selection.xlsx")
	selExportXLSX := mustCall(t, handler.selectionExportTool, map[string]any{
		"session_id": openPayload.SessionID,
		"range":      "A1:B2",
		"path":       exportXLSX,
	})
	var exportXLSXResp selectionExportResponse
	decodeResult(t, selExportXLSX, &exportXLSXResp)
	if exportXLSXResp.Format != "xlsx" || exportXLSXResp.SavedPath != exportXLSX {
		t.Fatalf("expected xlsx export metadata, got %+v", exportXLSXResp)
	}
	if exportXLSXResp.Rows != 2 || exportXLSXResp.Cols != 2 || exportXLSXResp.Bytes == 0 {
		t.Fatalf("unexpected xlsx export dimensions %+v", exportXLSXResp)
	}
	if _, err := os.Stat(exportXLSX); err != nil {
		t.Fatalf("expected xlsx export file: %v", err)
	}

	clipSet := mustCall(t, handler.clipboardSetTool, map[string]any{
		"session_id": openPayload.SessionID,
		"kind":       "cell",
		"value":      "X",
	})
	var clipResp clipboardGetResponse
	decodeResult(t, clipSet, &clipResp)
	if clipResp.Kind != "cell" || clipResp.CellValue != "X" {
		t.Fatalf("unexpected clipboard set resp %+v", clipResp)
	}

	clipGet := mustCall(t, handler.clipboardGetTool, map[string]any{
		"session_id": openPayload.SessionID,
	})
	decodeResult(t, clipGet, &clipResp)
	if clipResp.Kind != "cell" || clipResp.CellValue != "X" {
		t.Fatalf("unexpected clipboard get resp %+v", clipResp)
	}
	clipRangeVals := [][]string{{"1", "2"}, {"3", "4"}}
	clipRangeSet := mustCall(t, handler.clipboardSetTool, map[string]any{
		"session_id":  openPayload.SessionID,
		"kind":        "range",
		"rangeValues": clipRangeVals,
	})
	decodeResult(t, clipRangeSet, &clipResp)
	if clipResp.Kind != "range" || len(clipResp.RangeValues) != 2 {
		t.Fatalf("unexpected clipboard range set resp %+v", clipResp)
	}
	clipRangeGet := mustCall(t, handler.clipboardGetTool, map[string]any{
		"session_id": openPayload.SessionID,
	})
	decodeResult(t, clipRangeGet, &clipResp)
	if clipResp.Kind != "range" || clipResp.RangeValues[0][0] != "1" || clipResp.RangeValues[1][1] != "4" {
		t.Fatalf("unexpected clipboard range get resp %+v", clipResp)
	}

	stylePayload := [][]workbook.CellStyle{{{
		FillColor: "FF0000",
		FontColor: "000000",
		Bold:      true,
	}}}
	styleApply := mustCall(t, handler.styleApplyTool, map[string]any{
		"session_id": openPayload.SessionID,
		"range":      "B2",
		"styles":     stylePayload,
	})
	var applyResp styleApplyResponse
	decodeResult(t, styleApply, &applyResp)
	if applyResp.StyleStatus != "payload_applied" || applyResp.CellsAffected != 1 {
		t.Fatalf("unexpected style apply resp %+v", applyResp)
	}

	styledDescribe := mustCall(t, handler.styleDescribeTool, map[string]any{
		"session_id": openPayload.SessionID,
		"range":      "B2",
	})
	var describeResp styleDescribeResponse
	decodeResult(t, styledDescribe, &describeResp)
	if len(describeResp.Styles) == 0 {
		t.Fatalf("expected style metadata for B2")
	}
	if describeResp.Styles[0].Style.FillColor != "FF0000" {
		t.Fatalf("expected fill color FF0000, got %+v", describeResp.Styles[0].Style)
	}

	styleCopy := mustCall(t, handler.styleApplyTool, map[string]any{
		"session_id": openPayload.SessionID,
		"range":      "C3",
		"style_from": "B2",
	})
	decodeResult(t, styleCopy, &applyResp)
	if applyResp.StyleStatus != "copied_from_range" {
		t.Fatalf("expected copied_from_range status, got %+v", applyResp)
	}

	styledExport := mustCall(t, handler.selectionExportTool, map[string]any{
		"session_id": openPayload.SessionID,
		"range":      "B2:C3",
	})
	var styledExportResp selectionExportResponse
	decodeResult(t, styledExport, &styledExportResp)
	if len(styledExportResp.Styles) < 2 {
		t.Fatalf("expected styles in export, got %+v", styledExportResp.Styles)
	}

	closeRes := mustCall(t, handler.workspaceCloseTool, map[string]any{"session_id": openPayload.SessionID})
	var closePayload map[string]string
	decodeResult(t, closeRes, &closePayload)
	if closePayload["status"] != "closed" {
		t.Fatalf("expected closed status, got %v", closePayload)
	}
	if _, err := handler.sessions.Get(openPayload.SessionID); err == nil {
		t.Fatalf("expected session to be removed")
	}
}

func TestWorkspaceOpenCreateIfMissing(t *testing.T) {
	handler := &toolHandler{sessions: NewSessionManager()}

	t.Run("creates_csv_when_missing", func(t *testing.T) {
		path := filepath.Join(t.TempDir(), "new_workbook.csv")
		openRes := mustCall(t, handler.workspaceOpenTool, map[string]any{
			"path":             path,
			"create_if_missing": true,
		})
		var openPayload workspaceOpenResponse
		decodeResult(t, openRes, &openPayload)
		if openPayload.SessionID == "" {
			t.Fatalf("expected session id")
		}
		if len(openPayload.Sheets) == 0 || openPayload.Sheets[0] != "Sheet1" {
			t.Fatalf("expected Sheet1, got %v", openPayload.Sheets)
		}
		if _, err := os.Stat(path); err != nil {
			t.Fatalf("expected file to exist on disk: %v", err)
		}

		cursorRes := mustCall(t, handler.cursorGetTool, map[string]any{"session_id": openPayload.SessionID})
		var cursor cursorResponse
		decodeResult(t, cursorRes, &cursor)
		if cursor.Address != "A1" {
			t.Fatalf("expected cursor at A1, got %s", cursor.Address)
		}
	})

	t.Run("creates_xlsx_when_missing", func(t *testing.T) {
		path := filepath.Join(t.TempDir(), "new_workbook.xlsx")
		openRes := mustCall(t, handler.workspaceOpenTool, map[string]any{
			"path":              path,
			"create_if_missing": true,
		})
		var openPayload workspaceOpenResponse
		decodeResult(t, openRes, &openPayload)
		if openPayload.SessionID == "" {
			t.Fatalf("expected session id")
		}
		if _, err := os.Stat(path); err != nil {
			t.Fatalf("expected xlsx file to exist on disk: %v", err)
		}
	})

	t.Run("fails_without_create_if_missing", func(t *testing.T) {
		path := filepath.Join(t.TempDir(), "nonexistent.csv")
		req := mcp.CallToolRequest{}
		req.Params.Arguments = map[string]any{"path": path}
		_, err := handler.workspaceOpenTool(context.Background(), req)
		if err == nil {
			t.Fatalf("expected error for missing file without create_if_missing")
		}
	})

	t.Run("opens_existing_file_with_create_if_missing", func(t *testing.T) {
		path := writeSampleWorkbook(t)
		openRes := mustCall(t, handler.workspaceOpenTool, map[string]any{
			"path":              path,
			"create_if_missing": true,
		})
		var openPayload workspaceOpenResponse
		decodeResult(t, openRes, &openPayload)
		if openPayload.SessionID == "" {
			t.Fatalf("expected session id for existing file")
		}
		if len(openPayload.Sheets) == 0 {
			t.Fatalf("expected sheets for existing file")
		}
	})
}

func TestCellEditOutOfBoundsAddress(t *testing.T) {
	handler := &toolHandler{sessions: NewSessionManager()}
	path := writeSampleWorkbook(t)

	openRes := mustCall(t, handler.workspaceOpenTool, map[string]any{"path": path})
	var openPayload workspaceOpenResponse
	decodeResult(t, openRes, &openPayload)

	// Edit a cell well beyond current grid dimensions.
	editRes := mustCall(t, handler.cellEditTool, map[string]any{
		"session_id": openPayload.SessionID,
		"address":    "Z10",
		"value":      "out_of_bounds",
	})
	var editPayload cellEditResponse
	decodeResult(t, editRes, &editPayload)
	if editPayload.Address != "Z10" {
		t.Fatalf("expected address Z10, got %s", editPayload.Address)
	}
	if editPayload.Value != "out_of_bounds" {
		t.Fatalf("expected value out_of_bounds, got %s", editPayload.Value)
	}

	// Verify cursor landed at the target cell, not a clamped position.
	cursorRes := mustCall(t, handler.cursorGetTool, map[string]any{"session_id": openPayload.SessionID})
	var cursor cursorResponse
	decodeResult(t, cursorRes, &cursor)
	if cursor.Address != "Z10" {
		t.Fatalf("expected cursor at Z10, got %s", cursor.Address)
	}

	// Clear the out-of-bounds cell; cursor should move there too.
	clearRes := mustCall(t, handler.cellClearTool, map[string]any{
		"session_id": openPayload.SessionID,
		"address":    "Z10",
	})
	var clearPayload cellEditResponse
	decodeResult(t, clearRes, &clearPayload)
	if clearPayload.Address != "Z10" {
		t.Fatalf("expected clear address Z10, got %s", clearPayload.Address)
	}
}

func writeSampleWorkbook(t *testing.T) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "sample.csv")
	if err := workbook.SampleWorkbook().Save(path); err != nil {
		t.Fatalf("save sample workbook: %v", err)
	}
	return path
}

type toolFunc func(context.Context, mcp.CallToolRequest) (*mcp.CallToolResult, error)

func mustCall(t *testing.T, fn toolFunc, args map[string]any) *mcp.CallToolResult {
	t.Helper()
	req := mcp.CallToolRequest{}
	req.Params.Arguments = args
	res, err := fn(context.Background(), req)
	if err != nil {
		t.Fatalf("tool call failed: %v", err)
	}
	return res
}

func decodeResult[T any](t *testing.T, res *mcp.CallToolResult, dest *T) {
	t.Helper()
	if len(res.Content) == 0 {
		t.Fatalf("empty tool result")
	}
	text, ok := res.Content[0].(mcp.TextContent)
	if !ok {
		t.Fatalf("expected text content")
	}
	if err := json.Unmarshal([]byte(text.Text), dest); err != nil {
		t.Fatalf("unmarshal result: %v", err)
	}
}
