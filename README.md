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


