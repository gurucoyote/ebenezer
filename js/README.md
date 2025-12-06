# Ebenezer - headless yet interactive spreadsheet editor

## What it isn't

- A replacement for GUI office tools, For that it just isn't feature complete enough
- reliable with formula and live calculations, altho it has some features in that regard: it actually re-calculates most Excel formula while editing a spreadsheet. But it will barf on complex sheets with circular formula

## Installation

Like any comandline tool written in JS and node:

```
$ git clone https://github.com/gurucoyote/ebenezer
$ cd /ebenezer/
# install deps
$ npm install
# or globally, so you can use it anywhere
$ npm install -g
```

## Basic usage

```
$ ebenezer [filename]
```

if not given a filename, it will create an empty spreadsheet.

### keyboard bindings

ebeniezer uses a vim like modal approach.
It starts up in 'normal' mode, where you navigate and enter commands by pressing short key sequences.
Try 'h' or '?' to output the list of currently available key mappings.

```
 q  -  quit the program, no questions asked
 i  -  edit the current cell
 y  -  yank the current cell into paste buffer
 x  -  cut current cell to paste buffer
 O  -  insert blank row above current
 o  -  insert blank row below current
 Y, yy  -  yank the current row into paste buffer
 X, xx  -  cut the current row into paste buffer, leaving a blank row.
 D, dd  -  delete current row, shifting below rows up
 yc  -  yank current column into paste buffer
 dc  -  delete current column and shift remaining columns left.
 xc  -  cut current column into paste buffer.
 P  -  paste buffer above or before row or cell.
 p  -  paste buffer below or after row or cell.
 ns  -  new sheet
 ps  -  pick sheet
 wb  -  write the workbook to disk, asks for filename
 g  -  goto cell address
 :  -  enter repl mode to enter js code
 enter, return, down  -  move one cell down
 left  -  move one cell left
 right  -  move one cell right
 up  -  move one cell up
 h, ?  -  print this help message
```

