package mcp

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	"ebenezer/internal/actions"
	"ebenezer/internal/app"
	"ebenezer/internal/discovery"
	"ebenezer/internal/workbook"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
)

// Options control the metadata exposed by the MCP server.
type Options struct {
	Name    string
	Version string
}

// Serve starts the MCP server on stdio. The context is currently unused but
// accepted to mirror future session management hooks.
func Serve(ctx context.Context, opts Options) error {
	_ = ctx // placeholder until session management is wired
	s := newServer(normalizeOptions(opts))
	return server.ServeStdio(s)
}

func newServer(opts Options) *server.MCPServer {
	handler := &toolHandler{
		sessions: NewSessionManager(),
	}
	s := server.NewMCPServer(
		opts.Name,
		opts.Version,
		server.WithToolCapabilities(true),
	)
	s.AddTool(
		mcp.NewTool(
			"actions_list",
			mcp.WithDescription("List available Ebenezer actions and metadata."),
		),
		handler.actionsListTool,
	)
	s.AddTool(
		mcp.NewTool(
			"workspace_open",
			mcp.WithDescription("Open a workbook and allocate a new MCP session."),
			mcp.WithString("path", mcp.Required(), mcp.Description("Path to CSV/XLSX file")),
			mcp.WithString("sheet", mcp.Description("Optional sheet name when opening XLSX files")),
			mcp.WithString("mode", mcp.Description("Optional mode: read_only or read_write (default)")),
			mcp.WithBoolean("create_if_missing", mcp.Description("When true, create a new empty workbook if the file does not exist")),
		),
		handler.workspaceOpenTool,
	)
	s.AddTool(
		mcp.NewTool(
			"workspace_close",
			mcp.WithDescription("Close an existing MCP session."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier returned by workspace_open")),
		),
		handler.workspaceCloseTool,
	)
	s.AddTool(
		mcp.NewTool(
			"workbook_save",
			mcp.WithDescription("Save the workbook for the given session."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("path", mcp.Description("Optional destination path; defaults to original source path")),
		),
		handler.workbookSaveTool,
	)
	s.AddTool(
		mcp.NewTool(
			"cursor_get",
			mcp.WithDescription("Return the current cursor location/value for a session."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
		),
		handler.cursorGetTool,
	)
	s.AddTool(
		mcp.NewTool(
			"cursor_set",
			mcp.WithDescription("Move the cursor to an explicit cell address."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("address", mcp.Required(), mcp.Description("Excel-style address such as B12")),
		),
		handler.cursorSetTool,
	)
	s.AddTool(
		mcp.NewTool(
			"info_get",
			mcp.WithDescription("Return structured workbook metadata for the session."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithBoolean("details", mcp.Description("Include extended metadata (file size, timestamps)")),
		),
		handler.infoGetTool,
	)
	s.AddTool(
		mcp.NewTool(
			"cell_edit",
			mcp.WithDescription("Edit the specified cell (defaults to current cursor)."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("value", mcp.Required(), mcp.Description("Value to write into the cell")),
			mcp.WithString("address", mcp.Description("Optional Excel-style address (e.g., C5)")),
		),
		handler.cellEditTool,
	)
	s.AddTool(
		mcp.NewTool(
			"cell_clear",
			mcp.WithDescription("Clear the specified cell or current cursor."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("address", mcp.Description("Optional address to clear before returning cursor to that location")),
		),
		handler.cellClearTool,
	)
	s.AddTool(
		mcp.NewTool(
			"row_insert_above",
			mcp.WithDescription("Insert a blank row above the cursor (or optional address)."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("address", mcp.Description("Optional address used to position the cursor before inserting")),
		),
		handler.rowInsertAboveTool,
	)
	s.AddTool(
		mcp.NewTool(
			"row_insert_below",
			mcp.WithDescription("Insert a blank row below the cursor (or optional address)."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("address", mcp.Description("Optional address used to position the cursor before inserting")),
		),
		handler.rowInsertBelowTool,
	)
	s.AddTool(
		mcp.NewTool(
			"row_delete",
			mcp.WithDescription("Delete the row at the cursor (or optional address)."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("address", mcp.Description("Optional address used to position the cursor before deleting")),
		),
		handler.rowDeleteTool,
	)
	s.AddTool(
		mcp.NewTool(
			"column_insert_left",
			mcp.WithDescription("Insert a blank column before the cursor (or optional address)."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("address", mcp.Description("Optional address used to position the cursor before inserting")),
		),
		handler.columnInsertLeftTool,
	)
	s.AddTool(
		mcp.NewTool(
			"column_insert_right",
			mcp.WithDescription("Insert a blank column after the cursor (or optional address)."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("address", mcp.Description("Optional address used to position the cursor before inserting")),
		),
		handler.columnInsertRightTool,
	)
	s.AddTool(
		mcp.NewTool(
			"column_delete",
			mcp.WithDescription("Delete the column at the cursor (or optional address)."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("address", mcp.Description("Optional address used to position the cursor before deleting")),
		),
		handler.columnDeleteTool,
	)
	s.AddTool(
		mcp.NewTool(
			"column_width",
			mcp.WithDescription("Inspect or set column widths (show|set|auto)."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("mode", mcp.Required(), mcp.Description("show|set|auto")),
			mcp.WithString("columns", mcp.Description("Optional column span such as A:D or 2:4")),
			mcp.WithNumber("width", mcp.Description("Width when mode=set")),
			mcp.WithNumber("min_width", mcp.Description("Minimum width when mode=auto")),
			mcp.WithNumber("max_width", mcp.Description("Maximum width when mode=auto")),
			mcp.WithNumber("padding", mcp.Description("Padding added by the auto heuristic")),
			mcp.WithNumber("bonus", mcp.Description("Multiline bonus added by the auto heuristic")),
			mcp.WithNumber("factor", mcp.Description("Character width factor for the auto heuristic")),
		),
		handler.columnWidthTool,
	)
	s.AddTool(
		mcp.NewTool(
			"selection_set",
			mcp.WithDescription("Activate a rectangular selection (e.g., A1:D4)."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("range", mcp.Required(), mcp.Description("Excel-style range such as A1:C3")),
		),
		handler.selectionSetTool,
	)
	s.AddTool(
		mcp.NewTool(
			"selection_clear",
			mcp.WithDescription("Clear the active selection if one exists."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
		),
		handler.selectionClearTool,
	)
	s.AddTool(
		mcp.NewTool(
			"selection_export",
			mcp.WithDescription("Export the current selection or provided range as values or a file."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("range", mcp.Description("Optional range override")),
			mcp.WithString("path", mcp.Description("Optional .csv or .xlsx destination for the exported range")),
		),
		handler.selectionExportTool,
	)
	s.AddTool(
		mcp.NewTool(
			"clipboard_get",
			mcp.WithDescription("Return the current clipboard contents."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
		),
		handler.clipboardGetTool,
	)
	s.AddTool(
		mcp.NewTool(
			"clipboard_set",
			mcp.WithDescription("Set clipboard contents for future paste operations."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("kind", mcp.Required(), mcp.Description("Clipboard kind: cell, row, or range")),
			mcp.WithString("value", mcp.Description("Value for cell clipboard kind")),
			mcp.WithArray("rows", mcp.Items(map[string]any{"type": "array", "items": map[string]any{"type": "string"}}), mcp.Description("Rows payload for row clipboard kind")),
			mcp.WithArray("range_values", mcp.Items(map[string]any{"type": "array", "items": map[string]any{"type": "string"}}), mcp.Description("2D array payload for range clipboard kind")),
		),
		handler.clipboardSetTool,
	)
	s.AddTool(
		mcp.NewTool(
			"style_describe",
			mcp.WithDescription("Return style metadata for a given cell or range."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("range", mcp.Description("Optional range expression; defaults to current cell")),
		),
		handler.styleDescribeTool,
	)
	s.AddTool(
		mcp.NewTool(
			"style_apply",
			mcp.WithDescription("Apply formatting to a cell or range."),
			mcp.WithString("session_id", mcp.Required(), mcp.Description("Session identifier")),
			mcp.WithString("range", mcp.Description("Destination cell or range")),
			mcp.WithString("style_from", mcp.Description("Optional source range to copy styles from")),
			mcp.WithArray("styles", mcp.Items(map[string]any{"type": "array", "items": map[string]any{"type": "string"}}), mcp.Description("Optional 2D array of style objects to apply")),
		),
		handler.styleApplyTool,
	)
	return s
}

type toolHandler struct {
	sessions *SessionManager
}

// DTOs for tool arguments/responses.
type (
	workspaceOpenArgs struct {
		Path            string `json:"path"`
		Sheet           string `json:"sheet"`
		Mode            string `json:"mode"`
		CreateIfMissing bool   `json:"create_if_missing"`
	}
	workspaceOpenResponse struct {
		SessionID   string   `json:"sessionId"`
		Path        string   `json:"path"`
		Sheets      []string `json:"sheets"`
		ActiveSheet string   `json:"activeSheet"`
		ReadOnly    bool     `json:"readOnly"`
	}
	sessionIDArgs struct {
		SessionID string `json:"session_id"`
	}
	workbookSaveArgs struct {
		SessionID string `json:"session_id"`
		Path      string `json:"path"`
	}
	workbookSaveResponse struct {
		SessionID string `json:"sessionId"`
		Path      string `json:"path"`
	}
	cursorSetArgs struct {
		SessionID string `json:"session_id"`
		Address   string `json:"address"`
	}
	cursorResponse struct {
		SessionID string `json:"sessionId"`
		Sheet     string `json:"sheet"`
		Row       int    `json:"row"`
		Col       int    `json:"col"`
		Address   string `json:"address"`
		Value     string `json:"value"`
	}
	infoGetArgs struct {
		SessionID string `json:"session_id"`
		Details   bool   `json:"details"`
	}
	infoGetResponse struct {
		SessionID string               `json:"sessionId"`
		Info      actions.WorkbookInfo `json:"info"`
	}
	cellEditArgs struct {
		SessionID string `json:"session_id"`
		Value     string `json:"value"`
		Address   string `json:"address"`
	}
	cellEditResponse struct {
		SessionID   string `json:"sessionId"`
		Sheet       string `json:"sheet"`
		Address     string `json:"address"`
		Row         int    `json:"row"`
		Col         int    `json:"col"`
		Value       string `json:"value"`
		StyleStatus string `json:"styleStatus"`
	}
	cellClearArgs struct {
		SessionID string `json:"session_id"`
		Address   string `json:"address"`
	}
	rowOpArgs struct {
		SessionID string `json:"session_id"`
		Address   string `json:"address"`
	}
	rowOpResponse struct {
		SessionID string `json:"sessionId"`
		Sheet     string `json:"sheet"`
		Row       int    `json:"row"`
		Operation string `json:"operation"`
	}
	selectionSetArgs struct {
		SessionID string `json:"session_id"`
		Range     string `json:"range"`
	}
	selectionResponse struct {
		SessionID string `json:"sessionId"`
		Range     string `json:"range"`
		Summary   string `json:"summary"`
		Active    bool   `json:"active"`
	}
	selectionExportArgs struct {
		SessionID string `json:"session_id"`
		Range     string `json:"range"`
		Path      string `json:"path"`
	}
	selectionExportResponse struct {
		SessionID string            `json:"sessionId"`
		Range     string            `json:"range"`
		Values    [][]string        `json:"values,omitempty"`
		Rows      int               `json:"rows"`
		Cols      int               `json:"cols"`
		SavedPath string            `json:"savedPath,omitempty"`
		Format    string            `json:"format,omitempty"`
		Bytes     int64             `json:"bytes,omitempty"`
		Styles    []rangeStyleEntry `json:"styles,omitempty"`
	}
	columnOpArgs struct {
		SessionID string `json:"session_id"`
		Address   string `json:"address"`
	}
	columnOpResponse struct {
		SessionID string `json:"sessionId"`
		Sheet     string `json:"sheet"`
		Col       int    `json:"col"`
		Operation string `json:"operation"`
	}
	columnWidthArgs struct {
		SessionID string   `json:"session_id"`
		Mode      string   `json:"mode"`
		Columns   string   `json:"columns"`
		Width     float64  `json:"width"`
		MinWidth  *float64 `json:"min_width"`
		MaxWidth  *float64 `json:"max_width"`
		Padding   *float64 `json:"padding"`
		Bonus     *float64 `json:"bonus"`
		Factor    *float64 `json:"factor"`
	}
	columnWidthResponse struct {
		SessionID string             `json:"sessionId"`
		Mode      string             `json:"mode"`
		Range     string             `json:"range"`
		Columns   []columnWidthEntry `json:"columns"`
	}
	columnWidthEntry struct {
		Column   string  `json:"column"`
		Index    int     `json:"index"`
		Width    float64 `json:"width"`
		Source   string  `json:"source"`
		Explicit bool    `json:"explicit"`
	}
	clipboardGetResponse struct {
		SessionID   string     `json:"sessionId"`
		Kind        string     `json:"kind"`
		CellValue   string     `json:"cellValue,omitempty"`
		Rows        [][]string `json:"rows,omitempty"`
		RangeValues [][]string `json:"rangeValues,omitempty"`
	}
	clipboardSetArgs struct {
		SessionID   string     `json:"session_id"`
		Kind        string     `json:"kind"`
		Value       string     `json:"value"`
		Rows        [][]string `json:"rows"`
		RangeValues [][]string `json:"rangeValues"`
	}
	styleDescribeArgs struct {
		SessionID string `json:"session_id"`
		Range     string `json:"range"`
	}
	styleDescribeResponse struct {
		SessionID string            `json:"sessionId"`
		Range     string            `json:"range"`
		Styles    []rangeStyleEntry `json:"styles"`
		Values    [][]string        `json:"values,omitempty"`
	}
	styleApplyArgs struct {
		SessionID string                 `json:"session_id"`
		Range     string                 `json:"range"`
		StyleFrom string                 `json:"style_from"`
		Styles    [][]workbook.CellStyle `json:"styles"`
	}
	styleApplyResponse struct {
		SessionID     string `json:"sessionId"`
		Range         string `json:"range"`
		CellsAffected int    `json:"cellsAffected"`
		StyleStatus   string `json:"styleStatus"`
		Source        string `json:"source,omitempty"`
	}
	rangeStyleEntry struct {
		Row       int                `json:"row"`
		Col       int                `json:"col"`
		Address   string             `json:"address"`
		RowOffset int                `json:"rowOffset"`
		ColOffset int                `json:"colOffset"`
		Style     workbook.CellStyle `json:"style"`
		Value     string             `json:"value,omitempty"`
	}
)

func (h *toolHandler) actionsListTool(ctx context.Context, _ mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	data, err := discovery.MarshalActions(false)
	if err != nil {
		return nil, err
	}
	return textResult(json.RawMessage(data))
}

func (h *toolHandler) workspaceOpenTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args workspaceOpenArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	if strings.TrimSpace(args.Path) == "" {
		return nil, errors.New("path is required")
	}
	if args.CreateIfMissing {
		if _, err := os.Stat(args.Path); err != nil && os.IsNotExist(err) {
			wb := workbook.NewEmpty(filepath.Base(args.Path))
			if err := wb.Save(args.Path); err != nil {
				return nil, fmt.Errorf("create empty workbook: %w", err)
			}
		}
	}
	readOnly := strings.EqualFold(args.Mode, "read_only") || strings.EqualFold(args.Mode, "readonly")
	session, err := h.sessions.Open(args.Path, args.Sheet, readOnly)
	if err != nil {
		return nil, err
	}
	resp := workspaceOpenResponse{
		SessionID:   session.ID,
		Path:        session.State.SourcePath,
		Sheets:      append([]string(nil), session.State.SheetNames...),
		ActiveSheet: session.State.Workbook.Sheet,
		ReadOnly:    session.ReadOnly,
	}
	return jsonResult(resp)
}

func (h *toolHandler) workspaceCloseTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args sessionIDArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	if strings.TrimSpace(args.SessionID) == "" {
		return nil, errors.New("session_id is required")
	}
	if !h.sessions.Close(args.SessionID) {
		return nil, fmt.Errorf("session %s not found", args.SessionID)
	}
	return jsonResult(map[string]string{"sessionId": args.SessionID, "status": "closed"})
}

func (h *toolHandler) workbookSaveTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args workbookSaveArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	if session.ReadOnly {
		return nil, fmt.Errorf("session %s is read-only", session.ID)
	}
	target := strings.TrimSpace(args.Path)
	if target == "" {
		target = session.State.SourcePath
	}
	if strings.TrimSpace(target) == "" {
		return nil, errors.New("no destination path; provide path for workbook_save")
	}
	if err := session.State.Save(target); err != nil {
		return nil, err
	}
	return jsonResult(workbookSaveResponse{SessionID: session.ID, Path: target})
}

func (h *toolHandler) cursorGetTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	session, err := h.sessionFromRequest(req)
	if err != nil {
		return nil, err
	}
	state := session.State
	if state.Workbook == nil {
		return nil, errors.New("no workbook loaded in session")
	}
	resp := cursorResponse{
		SessionID: session.ID,
		Sheet:     state.Workbook.Sheet,
		Row:       state.Cursor.Row,
		Col:       state.Cursor.Col,
		Address:   state.Address(),
		Value:     state.CurrentValue(),
	}
	return jsonResult(resp)
}

func (h *toolHandler) cursorSetTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args cursorSetArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	if strings.TrimSpace(args.Address) == "" {
		return nil, errors.New("address is required")
	}
	if err := session.State.Goto(args.Address); err != nil {
		return nil, err
	}
	resp := cursorResponse{
		SessionID: session.ID,
		Sheet:     session.State.Workbook.Sheet,
		Row:       session.State.Cursor.Row,
		Col:       session.State.Cursor.Col,
		Address:   session.State.Address(),
		Value:     session.State.CurrentValue(),
	}
	return jsonResult(resp)
}

func (h *toolHandler) infoGetTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args infoGetArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	ctxActions := actions.NewContext(session.State, nil)
	var infoArgs []string
	if args.Details {
		infoArgs = []string{"--details"}
	}
	result, err := actions.Info.Exec(ctxActions, infoArgs)
	if err != nil {
		return nil, err
	}
	info, ok := result.Data.(actions.WorkbookInfo)
	if !ok {
		return nil, errors.New("unexpected info payload type")
	}
	return jsonResult(infoGetResponse{SessionID: session.ID, Info: info})
}

func (h *toolHandler) cellEditTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args cellEditArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	if strings.TrimSpace(args.Value) == "" {
		return nil, errors.New("value is required")
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	if strings.TrimSpace(args.Address) != "" {
		row, col, err := workbook.ParseCellAddress(args.Address)
		if err != nil {
			return nil, err
		}
		session.State.Workbook.SetCell(row, col, args.Value)
		session.State.Cursor = app.Cursor{Row: row, Col: col}
		session.State.SetDirty()
		session.State.Workbook.ActiveCell = session.State.Address()
	} else {
		if _, err := runAction(session, actions.Edit, []string{args.Value}); err != nil {
			return nil, err
		}
	}
	resp := cellEditResponse{
		SessionID:   session.ID,
		Sheet:       session.State.Workbook.Sheet,
		Row:         session.State.Cursor.Row,
		Col:         session.State.Cursor.Col,
		Address:     session.State.Address(),
		Value:       session.State.CurrentValue(),
		StyleStatus: "retained",
	}
	return jsonResult(resp)
}

func (h *toolHandler) cellClearTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args cellClearArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	if strings.TrimSpace(args.Address) != "" {
		row, col, err := workbook.ParseCellAddress(args.Address)
		if err != nil {
			return nil, err
		}
		session.State.Workbook.ClearCell(row, col)
		session.State.Cursor = app.Cursor{Row: row, Col: col}
		session.State.SetDirty()
		session.State.Workbook.ActiveCell = session.State.Address()
	} else {
		if _, err := runAction(session, actions.Clear, nil); err != nil {
			return nil, err
		}
	}
	resp := cellEditResponse{
		SessionID:   session.ID,
		Sheet:       session.State.Workbook.Sheet,
		Row:         session.State.Cursor.Row,
		Col:         session.State.Cursor.Col,
		Address:     session.State.Address(),
		Value:       session.State.CurrentValue(),
		StyleStatus: "retained",
	}
	return jsonResult(resp)
}

func (h *toolHandler) rowInsertAboveTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	return h.rowOp(ctx, req, actions.RowInsertAbove, "insert_above")
}

func (h *toolHandler) rowInsertBelowTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	return h.rowOp(ctx, req, actions.RowInsertBelow, "insert_below")
}

func (h *toolHandler) rowDeleteTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	return h.rowOp(ctx, req, actions.RowDelete, "delete")
}

func (h *toolHandler) rowOp(ctx context.Context, req mcp.CallToolRequest, action actions.Action, op string) (*mcp.CallToolResult, error) {
	_ = ctx
	var args rowOpArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	if strings.TrimSpace(args.Address) != "" {
		if err := session.State.Goto(args.Address); err != nil {
			return nil, err
		}
	}
	if _, err := runAction(session, action, nil); err != nil {
		return nil, err
	}
	resp := rowOpResponse{
		SessionID: session.ID,
		Sheet:     session.State.Workbook.Sheet,
		Row:       session.State.Cursor.Row,
		Operation: op,
	}
	return jsonResult(resp)
}

func (h *toolHandler) selectionSetTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args selectionSetArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	if strings.TrimSpace(args.Range) == "" {
		return nil, errors.New("range is required")
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	summary, err := session.State.SetSelectionRange(args.Range)
	if err != nil {
		return nil, err
	}
	resp := selectionResponse{
		SessionID: session.ID,
		Range:     args.Range,
		Summary:   summary,
		Active:    session.State.HasSelection(),
	}
	return jsonResult(resp)
}

func (h *toolHandler) selectionClearTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	session, err := h.sessionFromRequest(req)
	if err != nil {
		return nil, err
	}
	session.State.ClearSelection()
	resp := selectionResponse{
		SessionID: session.ID,
		Active:    false,
	}
	return jsonResult(resp)
}

func (h *toolHandler) columnInsertLeftTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	return h.columnOp(ctx, req, actions.ColumnInsertLeft, "insert_left")
}

func (h *toolHandler) columnInsertRightTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	return h.columnOp(ctx, req, actions.ColumnInsertRight, "insert_right")
}

func (h *toolHandler) columnDeleteTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	return h.columnOp(ctx, req, actions.ColumnDelete, "delete")
}

func (h *toolHandler) columnOp(ctx context.Context, req mcp.CallToolRequest, action actions.Action, op string) (*mcp.CallToolResult, error) {
	_ = ctx
	var args columnOpArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	if strings.TrimSpace(args.Address) != "" {
		if err := session.State.Goto(args.Address); err != nil {
			return nil, err
		}
	}
	if _, err := runAction(session, action, nil); err != nil {
		return nil, err
	}
	resp := columnOpResponse{
		SessionID: session.ID,
		Sheet:     session.State.Workbook.Sheet,
		Col:       session.State.Cursor.Col,
		Operation: op,
	}
	return jsonResult(resp)
}

func (h *toolHandler) columnWidthTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args columnWidthArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	mode := strings.ToLower(strings.TrimSpace(args.Mode))
	if mode == "" {
		return nil, errors.New("mode is required")
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	start, end, err := session.State.ResolveColumnSpan(strings.TrimSpace(args.Columns))
	if err != nil {
		return nil, err
	}
	var infos []app.ColumnWidthInfo
	switch mode {
	case "show":
		infos, err = session.State.ColumnWidths(start, end)
	case "set":
		if session.ReadOnly {
			return nil, fmt.Errorf("session %s is read-only", session.ID)
		}
		if args.Width <= 0 {
			return nil, errors.New("width must be greater than zero")
		}
		infos, err = session.State.SetColumnWidth(start, end, args.Width)
	case "auto":
		if session.ReadOnly {
			return nil, fmt.Errorf("session %s is read-only", session.ID)
		}
		opts := deriveWidthOptions(args)
		if opts.MaxWidth > 0 && opts.MinWidth > opts.MaxWidth {
			opts.MaxWidth = opts.MinWidth
		}
		infos, err = session.State.AutoColumnWidth(start, end, opts)
	default:
		return nil, fmt.Errorf("unknown column_width mode %q", mode)
	}
	if err != nil {
		return nil, err
	}
	resp := columnWidthResponse{
		SessionID: session.ID,
		Mode:      mode,
		Range:     columnRangeLabel(start, end),
		Columns:   mapColumnWidthEntries(infos),
	}
	return jsonResult(resp)
}

func deriveWidthOptions(args columnWidthArgs) workbook.ColumnWidthOptions {
	opts := workbook.DefaultColumnWidthOptions()
	if args.MinWidth != nil {
		opts.MinWidth = *args.MinWidth
	}
	if args.MaxWidth != nil {
		opts.MaxWidth = *args.MaxWidth
	}
	if args.Padding != nil {
		opts.Padding = *args.Padding
	}
	if args.Bonus != nil {
		opts.MultilineBonus = *args.Bonus
	}
	if args.Factor != nil {
		opts.CharacterFactor = *args.Factor
	}
	return opts
}

func mapColumnWidthEntries(infos []app.ColumnWidthInfo) []columnWidthEntry {
	entries := make([]columnWidthEntry, len(infos))
	for i, info := range infos {
		entries[i] = columnWidthEntry{
			Column:   workbook.ColumnName(info.Column),
			Index:    info.Column,
			Width:    info.Width,
			Source:   info.Source,
			Explicit: info.Explicit,
		}
	}
	sort.Slice(entries, func(i, j int) bool { return entries[i].Index < entries[j].Index })
	return entries
}

func columnRangeLabel(start, end int) string {
	if start == end {
		return workbook.ColumnName(start)
	}
	return fmt.Sprintf("%s:%s", workbook.ColumnName(start), workbook.ColumnName(end))
}

func (h *toolHandler) selectionExportTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args selectionExportArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	rangeInput := strings.TrimSpace(args.Range)
	if rangeInput == "" && session.State.HasSelection() {
		if startRow, startCol, endRow, endCol, ok := session.State.SelectionBounds(); ok {
			rangeInput = renderRange(startRow, startCol, endRow, endCol)
		}
	}
	values, styles, normalized, err := session.State.ExportRangeWithStyles(rangeInput)
	if err != nil {
		return nil, err
	}
	rows := len(values)
	cols := 0
	if rows > 0 {
		cols = len(values[0])
	}
	resp := selectionExportResponse{
		SessionID: session.ID,
		Range:     normalized,
		Values:    values,
		Rows:      rows,
		Cols:      cols,
	}
	if startRow, startCol, _, _, err := parseRangeBounds(normalized); err == nil {
		resp.Styles = flattenStyleEntries(styles, startRow, startCol, values)
	}
	path := strings.TrimSpace(args.Path)
	if path != "" {
		format, bytesWritten, err := app.SaveRangeToFile(values, styles, path, workbook.WithCSVDelimiter(session.State.CSVDelimiter()))
		if err != nil {
			return nil, err
		}
		resp.SavedPath = path
		resp.Format = format
		resp.Bytes = bytesWritten
	}
	return jsonResult(resp)
}

func (h *toolHandler) clipboardGetTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	session, err := h.sessionFromRequest(req)
	if err != nil {
		return nil, err
	}
	resp := clipboardSnapshot(session)
	return jsonResult(resp)
}

func (h *toolHandler) clipboardSetTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args clipboardSetArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	switch strings.ToLower(strings.TrimSpace(args.Kind)) {
	case "cell":
		if strings.TrimSpace(args.Value) == "" {
			return nil, errors.New("value required for cell clipboard")
		}
		session.State.Clipboard = app.Clipboard{Kind: app.ClipboardCell, CellValue: args.Value}
	case "row":
		if len(args.Rows) == 0 {
			return nil, errors.New("rows payload required for row clipboard")
		}
		session.State.Clipboard = app.Clipboard{Kind: app.ClipboardRow, Rows: clone2DStrings(args.Rows)}
	case "range":
		if len(args.RangeValues) == 0 {
			return nil, errors.New("rangeValues payload required for range clipboard")
		}
		session.State.Clipboard = app.Clipboard{Kind: app.ClipboardRange, RangeValues: clone2DStrings(args.RangeValues)}
	default:
		return nil, fmt.Errorf("unsupported clipboard kind %q", args.Kind)
	}
	resp := clipboardSnapshot(session)
	return jsonResult(resp)
}

func (h *toolHandler) styleDescribeTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args styleDescribeArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	values, styles, normalized, err := session.State.ExportRangeWithStyles(strings.TrimSpace(args.Range))
	if err != nil {
		return nil, err
	}
	startRow, startCol, _, _, parseErr := parseRangeBounds(normalized)
	if parseErr != nil {
		startRow = session.State.Cursor.Row
		startCol = session.State.Cursor.Col
	}
	resp := styleDescribeResponse{
		SessionID: session.ID,
		Range:     normalized,
		Styles:    flattenStyleEntries(styles, startRow, startCol, values),
		Values:    values,
	}
	return jsonResult(resp)
}

func (h *toolHandler) styleApplyTool(ctx context.Context, req mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	var args styleApplyArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	session, err := h.sessions.Get(strings.TrimSpace(args.SessionID))
	if err != nil {
		return nil, err
	}
	if session.ReadOnly {
		return nil, fmt.Errorf("session %s is read-only", session.ID)
	}
	targetRange := strings.TrimSpace(args.Range)
	if targetRange == "" {
		targetRange = session.State.Address()
	}
	status := ""
	source := ""
	if hasStylePayload(args.Styles) {
		if err := session.State.ApplyStyles(targetRange, args.Styles); err != nil {
			return nil, err
		}
		status = "payload_applied"
		source = "payload"
	} else if strings.TrimSpace(args.StyleFrom) != "" {
		if err := session.State.CopyStyle(args.StyleFrom); err != nil {
			return nil, err
		}
		if err := session.State.PasteStyle(targetRange); err != nil {
			return nil, err
		}
		status = "copied_from_range"
		source = strings.ToUpper(strings.TrimSpace(args.StyleFrom))
	} else {
		return nil, errors.New("provide style_from or styles payload")
	}
	startRow, startCol, endRow, endCol, err := parseRangeBounds(targetRange)
	if err != nil {
		startRow = session.State.Cursor.Row
		startCol = session.State.Cursor.Col
		endRow = startRow
		endCol = startCol
	}
	normalized := renderRange(startRow, startCol, endRow, endCol)
	cells := (endRow - startRow + 1) * (endCol - startCol + 1)
	resp := styleApplyResponse{
		SessionID:     session.ID,
		Range:         normalized,
		CellsAffected: cells,
		StyleStatus:   status,
		Source:        source,
	}
	return jsonResult(resp)
}

func (h *toolHandler) sessionFromRequest(req mcp.CallToolRequest) (*Session, error) {
	var args sessionIDArgs
	if err := decodeArgs(req.Params.Arguments, &args); err != nil {
		return nil, err
	}
	if strings.TrimSpace(args.SessionID) == "" {
		return nil, errors.New("session_id is required")
	}
	return h.sessions.Get(args.SessionID)
}

func decodeArgs(args map[string]interface{}, dest any) error {
	if dest == nil {
		return errors.New("destination is nil")
	}
	if args == nil {
		args = map[string]interface{}{}
	}
	data, err := json.Marshal(args)
	if err != nil {
		return err
	}
	return json.Unmarshal(data, dest)
}

func jsonResult(payload any) (*mcp.CallToolResult, error) {
	data, err := json.Marshal(payload)
	if err != nil {
		return nil, err
	}
	return mcp.NewToolResultText(string(data)), nil
}

func textResult(payload any) (*mcp.CallToolResult, error) {
	switch v := payload.(type) {
	case json.RawMessage:
		return mcp.NewToolResultText(string(v)), nil
	default:
		return jsonResult(v)
	}
}

func normalizeOptions(opts Options) Options {
	if opts.Name == "" {
		opts.Name = "ebenezer"
	}
	if opts.Version == "" {
		opts.Version = "0.0.0-dev"
	}
	return opts
}

func runAction(session *Session, action actions.Action, args []string) (actions.Result, error) {
	ctx := actions.NewContext(session.State, nil)
	return action.Exec(ctx, args)
}

func renderRange(startRow, startCol, endRow, endCol int) string {
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	return fmt.Sprintf("%s%d:%s%d", workbook.ColumnName(startCol), startRow, workbook.ColumnName(endCol), endRow)
}

func clone2DStrings(values [][]string) [][]string {
	if len(values) == 0 {
		return nil
	}
	dup := make([][]string, len(values))
	for i, row := range values {
		if row == nil {
			continue
		}
		copyRow := make([]string, len(row))
		copy(copyRow, row)
		dup[i] = copyRow
	}
	return dup
}

func clipboardSnapshot(session *Session) clipboardGetResponse {
	cp := session.State.Clipboard
	resp := clipboardGetResponse{
		SessionID: session.ID,
	}
	switch cp.Kind {
	case app.ClipboardCell:
		resp.Kind = "cell"
		resp.CellValue = cp.CellValue
	case app.ClipboardRow:
		resp.Kind = "row"
		resp.Rows = cp.Rows
	case app.ClipboardRange:
		resp.Kind = "range"
		resp.RangeValues = cp.RangeValues
	default:
		resp.Kind = "empty"
	}
	return resp
}

func hasStylePayload(styles [][]workbook.CellStyle) bool {
	if len(styles) == 0 {
		return false
	}
	for _, row := range styles {
		if len(row) > 0 {
			return true
		}
	}
	return false
}

func flattenStyleEntries(styles map[int]map[int]workbook.CellStyle, startRow, startCol int, values [][]string) []rangeStyleEntry {
	if len(styles) == 0 {
		return nil
	}
	entries := make([]rangeStyleEntry, 0, len(styles))
	for rOffset, cols := range styles {
		for cOffset, style := range cols {
			row := startRow + rOffset
			col := startCol + cOffset
			entry := rangeStyleEntry{
				Row:       row,
				Col:       col,
				Address:   fmt.Sprintf("%s%d", workbook.ColumnName(col), row),
				RowOffset: rOffset,
				ColOffset: cOffset,
				Style:     style,
			}
			if rOffset < len(values) && cOffset < len(values[rOffset]) {
				entry.Value = values[rOffset][cOffset]
			}
			entries = append(entries, entry)
		}
	}
	sort.Slice(entries, func(i, j int) bool {
		if entries[i].Row == entries[j].Row {
			return entries[i].Col < entries[j].Col
		}
		return entries[i].Row < entries[j].Row
	})
	return entries
}

func parseRangeBounds(rangeStr string) (int, int, int, int, error) {
	parts := strings.Split(strings.ToUpper(strings.TrimSpace(rangeStr)), ":")
	if len(parts) == 0 || len(parts) > 2 {
		return 0, 0, 0, 0, fmt.Errorf("invalid range %s", rangeStr)
	}
	startRow, startCol, err := parseCellAddress(parts[0])
	if err != nil {
		return 0, 0, 0, 0, err
	}
	endRow, endCol := startRow, startCol
	if len(parts) == 2 {
		endRow, endCol, err = parseCellAddress(parts[1])
		if err != nil {
			return 0, 0, 0, 0, err
		}
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	return startRow, startCol, endRow, endCol, nil
}

func parseCellAddress(addr string) (int, int, error) {
	addr = strings.TrimSpace(strings.ToUpper(addr))
	if addr == "" {
		return 0, 0, fmt.Errorf("empty address")
	}
	var letters, digits strings.Builder
	for _, r := range addr {
		switch {
		case r >= 'A' && r <= 'Z':
			letters.WriteRune(r)
		case r >= '0' && r <= '9':
			digits.WriteRune(r)
		default:
			return 0, 0, fmt.Errorf("invalid address %s", addr)
		}
	}
	if letters.Len() == 0 || digits.Len() == 0 {
		return 0, 0, fmt.Errorf("invalid address %s", addr)
	}
	col := columnLettersToNumber(letters.String())
	row, err := strconv.Atoi(digits.String())
	if err != nil {
		return 0, 0, err
	}
	return row, col, nil
}

func columnLettersToNumber(input string) int {
	result := 0
	for _, r := range input {
		if r < 'A' || r > 'Z' {
			continue
		}
		result = result*26 + int(r-'A'+1)
	}
	if result == 0 {
		return 1
	}
	return result
}
