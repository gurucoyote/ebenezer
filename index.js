#! /usr/bin/env node
const path = require("path");
const readline = require("readline");
const repl = require("repl");
const Excel = require("exceljs");
const FormulaParser = require("hot-formula-parser").Parser;
// how long to wait for completion of key sequences
const keyWait = 250;

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});
rl.input.setEncoding("utf-8");
rl.input.setRawMode(true);
const defaultKPL = rl.input.listeners("keypress");
async function run() {
  const eb = {};
  eb.filename = process.argv[2];
  eb.filenames = [eb.filename || "untitled.xlsx"];
  eb.workbook = await load(eb.filename);
  eb.worksheet = eb.workbook.worksheets[0]; //the first one
  eb.row = 1;
  eb.col = 1;
  console.log(`on worksheet ${eb.worksheet.name}`);
  reportCell();
  switchNormalMode();

  const parser = new FormulaParser();
  parser.on("callCellValue", function (cellCoord, done) {
    // console.debug("callCellValue", cellCoord.label);
    if (eb.worksheet.getCell(cellCoord.label).formula) {
      done(parser.parse(eb.worksheet.getCell(cellCoord.label).formula).result);
    } else {
      done(eb.worksheet.getCell(cellCoord.label).value);
    }
  });
  parser.on("callRangeValue", function (startCellCoord, endCellCoord, done) {
    // console.debug("callRangeValue", startCellCoord.label, endCellCoord.label);
    var fragment = [];

    for (
      var row = startCellCoord.row.index;
      row <= endCellCoord.row.index;
      row++
    ) {
      var colFragment = [];

      for (
        var col = startCellCoord.column.index;
        col <= endCellCoord.column.index;
        col++
      ) {
        colFragment.push(eb.worksheet.getRow(row + 1).getCell(col + 1).value);
      }

      fragment.push(colFragment);
    }

    if (fragment) {
      done(fragment);
    }
  });

  let keySequence = "";
  function kp(c, key) {
    if (c && c.match(/[a-zA-Z]/)) {
      keySequence += c;
    } else {
      keySequence += key.name || key.sequence;
    }
    let clearSequenceTimeout;
    clearTimeout(clearSequenceTimeout);
    clearSequenceTimeout = setTimeout(function () {
      // did the user enter a known key sequence?
      if (cmds[keySequence]) cmds[keySequence].f();
      // clear out anything unknown
      keySequence = "";
    }, keyWait);
  }

  const help = {
    help: "print this help message",
    f: () => {
      let compact = {};
      Object.entries(cmds).map((c) => {
        const [k, v] = c;
        if (!compact[v.help]) {
          compact[v.help] = [k];
        } else {
          compact[v.help].push(k);
        }
      });
      Object.entries(compact).map((e) => {
        const [help, cmds] = e;
        console.log(cmds.join(", "), " - ", help);
      });
    },
  };
  const readColumnTitle = {
    help: "read the column title/header",
    f: () => {
      reportCell(eb.worksheet.getRow(1).getCell(eb.col));
    },
  };
  const readRowTitle = {
    help: "read the row title (first cell in row)",
    f: () => {
      reportCell(eb.worksheet.getRow(eb.row).getCell(1));
    },
  };
  const down = {
    help: "move one cell down",
    f: () => {
      eb.row += 1;
      reportCell();
    },
  };
  const cutCell = {
    help: "cut current cell to paste buffer",
    f: () => {
      yankCell.f();
      eb.worksheet.getRow(eb.row).getCell(eb.col).value = null;
    },
  };
  const yankCell = {
    help: "yank the current cell into paste buffer",
    f: () => {
      eb.yank = {
        type: "cell",
        y: eb.worksheet.getRow(eb.row).getCell(eb.col).value,
      };
    },
  };
  const yankRow = {
    help: "yank the current row into paste buffer",
    f: () => {
      console.log(`yanking row ${eb.row}`);
      eb.yank = { type: "row", y: [eb.worksheet.getRow(eb.row).values] };
    },
  };
  const cutRow = {
    help: "cut the current row into paste buffer, leaving a blank row.",
    f: () => {
      yankRow.f();
      console.log(`blanking row ${eb.row}`);
      eb.worksheet.getRow(eb.row).eachCell((c) => {
        c.value = null;
      });
      reportCell();
    },
  };
  const deleteRow = {
    help: "delete current row, shifting below rows up",
    f: () => {
      yankRow.f();
      console.log(`removing row ${eb.row}`);
      eb.worksheet.spliceRows(eb.row, 1);
      reportCell();
    },
  };
  const yankCol = {
    help: "yank current column into paste buffer",
    f: () => {
      eb.yank = {
        type: "col",
        // splice(1) returns all elements except first
        // apparently, inserting values into columns does not need the empty first element that values returns ^^
        y: eb.worksheet.getColumn(eb.col).values.splice(1),
      };
    },
  };
  const deleteCol = {
    help: "delete current column and shift remaining columns left.",
    f: () => {
      yankCol.f();
      eb.worksheet.spliceColumns(eb.col, 1);
    },
  };
  const cutCol = {
    help: "cut current column into paste buffer.",
    f: () => {
      yankCol.f();
      eb.worksheet.getColumn(eb.col).eachCell((cell, _) => {
        cell.value = null;
      });
    },
  };
  function paste(dir) {
    if (eb.yank && eb.yank.type === "row") {
      eb.worksheet.insertRows(eb.row + dir, eb.yank.y);
      eb.row = eb.row + dir;
      reportCell();
    } else if (eb.yank && eb.yank.type === "cell") {
      eb.worksheet.getRow(eb.row).getCell(eb.col).value = eb.yank.y;
      console.log("pasted:");
      reportCell();
    } else if (eb.yank && eb.yank.type === "col") {
      eb.worksheet.spliceColumns(eb.col + dir, 0, eb.yank.y);
      console.log("pasted column");
      eb.col = eb.col + dir;
      reportCell();
    } else {
      console.log("nothing to paste in buffer.");
    }
  }
  const pasteBefore = {
    help: "paste buffer above or before row or cell.",
    f: () => {
      paste(0);
    },
  };
  const pasteAfter = {
    help: "paste buffer below or after row or cell.",
    f: () => {
      paste(1);
    },
  };
  const insertRowBelow = {
    help: "insert blank row below current",
    f: () => {
      eb.row += 1;
      eb.worksheet.insertRow(eb.row);
      reportCell();
    },
  };
  const insertRowAbove = {
    help: "insert blank row above current",
    f: () => {
      eb.worksheet.insertRow(eb.row);
      reportCell();
    },
  };
  const findIncol = {
    help: "find string in current column, returns list of matching cells to select from",
    f: () => {
      switchInsertMode();
      // ask for search string
      rl.question(
        "search>",
        { signal: eb.abortSignal },
        async function (search) {
          // prepare history to select from
          rl.history = [];
          // find matching cells
          const column = eb.worksheet.getColumn(eb.col);
          const colLetter = n2l(eb.col);
          const re = new RegExp(search, "i");
          column.eachCell((c, idx) => {
            if (c.value && c.value.match(re)) {
              rl.history.push(colLetter + idx + " : " + c.value);
            }
          });
          rl.question(
            "goto>",
            { signal: eb.abortSignal },
            async function (answer) {
              // goto selected cell
              console.log("going to " + answer);
              goto(answer);
            }
          );
          rl.write("", {
            sequence: "\x1B[A",
            name: "up",
            ctrl: false,
            meta: false,
            shift: false,
            code: "[A",
          });
        }
      );
    },
  };
  let cmds = {
    q: {
      help: "quit the program, no questions asked",
      f: () => {
        process.exit();
      },
    },
    i: {
      help: "edit the current cell",
      f: () => {
        switchInsertMode();
        const cell = eb.worksheet.getRow(eb.row).getCell(eb.col);
        let cellcontent;
        if (cell.formula) cellcontent = "=" + cell.formula;
        else cellcontent = cell.value;
        rl.history = uniqueColumnValues(eb.col);
        rl.question(
          "cell value> ",
          { signal: eb.abortSignal },
          function (answer) {
            // write the edit to the sheet
            if (answer.charAt(0) === "=") {
              cell.value = { formula: answer.substring(1) };
              // substring(1) = from 2nd char to end of string
            } else {
              cell.value = answer;
            }
            reportCell(cell);
            switchNormalMode();
          }
        );
        // provide default anser that can be edited
        rl.write(cellcontent);
      },
    },
    y: yankCell,
    x: cutCell,
    O: insertRowAbove,
    o: insertRowBelow,
    Y: yankRow,
    yy: yankRow,
    X: cutRow,
    xx: cutRow,
    D: deleteRow,
    dd: deleteRow,
    yc: yankCol,
    dc: deleteCol,
    xc: cutCol,
    P: pasteBefore,
    p: pasteAfter,
    ns: {
      help: "new sheet",
      f: () => {
        // create a worksheet
        switchInsertMode();
        rl.question(
          "sheetname?",
          { signal: eb.abortSignal },
          function (answer) {
            try {
              eb.worksheet = eb.workbook.addWorksheet(answer);
              console.log(`created sheet ${answer}`);
            } catch (e) {
              console.error(e.message);
            } finally {
              switchNormalMode();
            }
          }
        );
      },
    },
    fi: findIncol,
    ps: {
      help: "pick sheet",
      f: () => {
        // select a worksheet
        switchInsertMode();
        // prepare history to select from
        rl.history = [];
        eb.workbook.worksheets.map((ws) => {
          rl.history.push(ws.name);
        });
        rl.question("ws>", { signal: eb.abortSignal }, async function (answer) {
          try {
            const ws = eb.workbook.getWorksheet(answer);
            if (ws) {
              eb.worksheet = ws;
              console.log(`selected ws ${answer}`);
            } else {
              console.log("no such sheet.");
            }
          } catch (e) {
            console.error(e);
          }
          switchNormalMode();
        });
        rl.write("", {
          sequence: "\x1B[A",
          name: "up",
          ctrl: false,
          meta: false,
          shift: false,
          code: "[A",
        });
      },
    },
    wb: {
      help: "write the workbook to disk, asks for filename",
      f: () => {
        switchInsertMode();
        // console.log(
        //   "enter new filename, or use up/down arrow to choose previous"
        // );
        rl.history = [];
        rl.history = [...eb.filenames];
        rl.question(
          "filename>",
          { signal: eb.abortSignal },
          async function (fn) {
            if (await save(eb, fn)) {
              eb.filenames = [...new Set(eb.filenames).add(fn)];
            }
            switchNormalMode();
          }
        );
        // write arrow up to enter history
        rl.write("", {
          sequence: "\x1B[A",
          name: "up",
          ctrl: false,
          meta: false,
          shift: false,
          code: "[A",
        });
      },
    },
    g: {
      help: "goto cell address",
      f: () => {
        // goto cell
        let currCell = eb.worksheet.getRow(eb.row).getCell(eb.col);
        switchInsertMode();
        rl.question("goto> ", { signal: eb.abortSignal }, goto);
        rl.write(currCell.address);
      },
    },
    ":": {
      help: "enter repl mode to enter js code",
      f: () => {
        rl.input.removeAllListeners("keypress");
        const r = repl.start({
          prompt: "repl>",
          ignoreUndefined: true,
        });
        r.context.eb = eb;
        r.context.rl = rl;
        r.context.humanFileSize = humanFileSize;
        r.defineCommand("q", {
          help: "leave current repl",
          action() {
            this.clearBufferedCommand();
            switchNormalMode();
          },
        });
      },
    },
    ct: readColumnTitle,
    rt: readRowTitle,
    enter: down,
    return: down,
    down: down,
    left: {
      help: "move one cell left",
      f: () => {
        eb.col -= 1;
        reportCell();
      },
    },
    right: {
      help: "move one cell right",
      f: () => {
        eb.col += 1;
        reportCell();
      },
    },
    up: {
      help: "move one cell up",
      f: () => {
        eb.row -= 1;
        reportCell();
      },
    },
    h: help,
    "?": help,
  };
  // define commands at later stage:
  cmds["ab"] = {
    help: "a simple test for sequences of keys, outputs 'ab' if that sequence was enterd",
    f: () => {
      console.log("saw:" + keySequence);
    },
  };
  function uniqueColumnValues(col) {
    const column = eb.worksheet.getColumn(col || eb.col);
    let uv = new Set();
    column.eachCell((c) => {
      uv.add(c.value);
    });
    // console.debug(uv);
    return [...uv];
  }
  function goto(gt) {
    try {
      const [_, col, row] = gt.match(/([a-z]+)(\d+)/i);
      eb.row = parseInt(row);
      eb.col = l2n(col);
      currCell = eb.worksheet.getRow(eb.row).getCell(eb.col);
    } catch (e) {
      console.warn(`${gt} is not a valid cell address`);
      // console.log(e);
    } finally {
      reportCell(currCell);
      switchNormalMode();
    }
  }
  function reportCell(cell) {
    if (!cell) {
      eb.row = eb.row < 1 ? 1 : eb.row;
      eb.col = eb.col < 1 ? 1 : eb.col;
      cell = eb.worksheet.getRow(eb.row).getCell(eb.col);
    }
    let delim = ":";
    if (cell.formula) {
      delim = "formula result =";
      // console.log('formula:', cell.formula);
    }
    console.log(cell.address, delim, getCellResult(eb.worksheet, cell.address));
  }
  function getCellResult(worksheet, cellAddress) {
    if (worksheet.getCell(cellAddress).formula) {
      return parser.parse(worksheet.getCell(cellAddress).formula).result;
    } else {
      return worksheet.getCell(cellAddress).value;
    }
  }
  function switchInsertMode() {
    // enter line edit mode
    rl.input.removeAllListeners("keypress");
    // restore default kp listener
    defaultKPL.map((f) => {
      rl.input.on("keypress", f);
    });
    // install abort controller for escape key
    const ac = new AbortController();
    eb.abortSignal = ac.signal;
    // eb.abortSignal.addEventListener('abort', () => {
    //   console.log('The question was aborted');
    // }, { once: true });
    rl.input.on("keypress", (_, key) => {
      if (key.name === "escape") {
        ac.abort();
        switchNormalMode();
      }
      // else console.log(key.name)
    });
  }
  function switchNormalMode() {
    rl.input.setRawMode(true);
    rl.input.removeAllListeners("keypress");
    rl.input.on("keypress", kp);
  }
}
run();
async function load(filename) {
  const workbook = new Excel.Workbook();
  if (!filename) {
    console.log("creating empty workbook");
    workbook.addWorksheet("Mappe 1");
    return workbook;
  }
  const ext = path.extname(filename).toLowerCase();
  console.log(`opening file ${filename}`);
  if (ext === ".xlsx") {
    await workbook.xlsx.readFile(filename);
  } else if (ext === ".csv") {
    const options = {
      parserOptions: {
        delimiter: ";",
        quote: false,
      },
    };
    await workbook.csv.readFile(filename, options);
  } else {
    console.log(`extension is: ${ext}`);
    console.log(
      "please provide a filename with either .xlsx or .csv extension."
    );
    process.exit(1);
  }
  return workbook;
}
async function save(eb, filename) {
  const ext = path.extname(filename).toLowerCase();
  if (ext === ".xlsx") {
    await eb.workbook.xlsx.writeFile(filename);
    console.log(`sheet written to ${filename}`);
    return true;
  } else if (ext === ".csv") {
    const options = {
      sheetName: eb.worksheet.name, // the currently selected
      formatterOptions: {
        delimiter: ";",
        quote: false,
      },
    };
    await eb.workbook.csv.writeFile(filename, options);
    console.log(`sheet written to ${filename}`);
    return true;
  } else {
    console.log(
      "please use either .xlsx or .csv file extension to specify file format"
    );
    return false;
  }
}
function n2l(num) {
  let s = "",
    t;
  while (num > 0) {
    t = (num - 1) % 26;
    s = String.fromCharCode(65 + t) + s;
    num = ((num - t) / 26) | 0;
  }
  return s || undefined;
}

function l2n(str) {
  str = str.toUpperCase();
  let out = 0,
    len = str.length;
  for (pos = 0; pos < len; pos++) {
    out += (str.charCodeAt(pos) - 64) * Math.pow(26, len - pos - 1);
  }
  return out;
}
function humanFileSize(size) {
  var i = Math.floor(Math.log(size) / Math.log(1024));
  return (
    (size / Math.pow(1024, i)).toFixed(2) * 1 +
    " " +
    ["B", "kB", "MB", "GB", "TB"][i]
  );
}
