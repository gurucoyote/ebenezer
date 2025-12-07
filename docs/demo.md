# Ebenezer CLI Demo

This walkthrough exercises the current Go skeleton so you can see the keyboard loop drive simple workbook navigation.

## Build
```
go build ./cmd/ebenezer
```

## Run with Sample Workbook or Real Files
```
./ebenezer data.xlsx     # load an existing workbook (csv/xlsx)
# or start with the built-in sample data
./ebenezer
```

### Normal Mode Bindings (current scope)
- Arrow keys: move the cursor (updates stdout with coordinates/value).
- `g`: prompt for a cell address (defaulting to the current cell) and jump there.
- `s`: print the current cell status.
- `c` then `t`: print the current column header (row 1); `r` then `t`: print the current row header (column 1).
- `:`: enter command mode (type commands such as `status`, `open test.csv`, `sample`).
- `q`: exit keyboard mode.

### Command Examples (`:` prompt)
```
:status               # show active cell/value
:open js/test.csv     # load an existing CSV in the repo
:open js/test.xlsx    # load the sample XLSX (use --sheet NAME as needed)
:style B5             # describe formatting of B5 (fill/font/bold/italic)
:colheader            # print the header of the current column (row 1)
:rowheader            # print the header of the current row (column 1)
:sample               # reload built-in sample workbook
```

These commands exercise the workbook loader (`internal/workbook`), shared app state (`internal/app`), and the keyboard loop (`internal/ui/keyboard`).
