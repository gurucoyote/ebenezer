const path = require("path");
const readline = require("readline");
const repl = require("repl");
const Excel = require("exceljs");
const FormulaParser = require("hot-formula-parser").Parser;

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
  historySize: 0, // so arrow up/down don't enter text ^^
});
rl.input.setEncoding("utf-8");
rl.input.setRawMode(true);
const defaultKPL = rl.input.listeners("keypress");
async function run() {
  const eb = {};
  eb.filename = process.argv[2];
  eb.filenames = [eb.filename];
  eb.workbook = await load(eb.filename);
  eb.worksheet = eb.workbook.worksheets[0]; //the first one
  eb.row = 1;
  eb.col = 1;
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
  function kp(_, key) {
    keySequence = keySequence + (key.name || key.sequence);
    let clearSequenceTimeout;
    clearTimeout(clearSequenceTimeout);
    clearSequenceTimeout = setTimeout(function () {
      keySequence = "";
    }, 500);
    if (cmds[keySequence]) cmds[keySequence].f();
    eb.row = eb.row < 1 ? 1 : eb.row;
    eb.col = eb.col < 1 ? 1 : eb.col;
    const cell = eb.worksheet.getRow(eb.row).getCell(eb.col);
    reportCell(cell);
  }

  let cmds = {
    ab: {
      help: "a simple test for sequences of keys",
      f: () => {
        console.log("saw:" + keySequence);
      },
    },
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
        rl.question("cell value> ", function (answer) {
          // write the edit to the sheet
          if (answer.charAt(0) === "=") {
            cell.value = { formula: answer.substring(1) };
            // substring(1) = from 2nd char to end of string
          } else {
            cell.value = answer;
          }
          reportCell(cell);
          switchNormalMode();
        });
        // provide default anser that can be edited
        let cellcontent;
        if (cell.formula) cellcontent = "=" + cell.formula;
        else cellcontent = cell.value;
        rl.write(cellcontent);
        // return from the function, so that the latter code won't be executed
        return;
      },
    },
    c: {
      help: "create a worksheet",
      f: () => {
        // create a worksheet
        switchInsertMode();
        rl.question("sheetname?", function (answer) {
          try {
            eb.worksheet = eb.workbook.addWorksheet(answer);
            console.log(`created sheet ${answer}`);
          } catch (e) {
            console.error(e.message);
          } finally {
            switchNormalMode();
            return;
          }
        });
        return;
      },
    },
    s: {
      help: "select a worksheet",
      f: () => {
        // select a worksheet
        switchInsertMode();
        rl.question("ws>", async function (answer) {
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
        eb.workbook.worksheets.map((ws) => {
          rl.history.push(ws.name);
        });
        rl.write("", {
          sequence: "\x1B[A",
          name: "up",
          ctrl: false,
          meta: false,
          shift: false,
          code: "[A",
        });
        return;
      },
    },
    w: {
      help: "write the workbook to disk, asks for filename",
      f: () => {
        switchInsertMode();
        console.log(
          "enter new filename, or use up/down arrow to choose previous"
        );
        rl.question("filename> ", async function (fn) {
          if (await save(eb.workbook, fn)) {
            // eb.filenames.push(fn);
            eb.filenames = [...new Set(eb.filenames).add(fn)];
          }
          switchNormalMode();
        });
        rl.history = eb.filenames;
        // write arrow up to enter history
        rl.write("", {
          sequence: "\x1B[A",
          name: "up",
          ctrl: false,
          meta: false,
          shift: false,
          code: "[A",
        });
        return;
      },
    },
    g: {
      help: "goto cell address",
      f: () => {
        // goto cell
        let currCell = eb.worksheet.getRow(eb.row).getCell(eb.col);
        switchInsertMode();
        rl.question("goto> ", function (gt) {
          try {
            const [_, col, row] = gt.match(/([a-z]+)(\d+)/i);
            eb.row = parseInt(row);
            eb.col = convertLetterToNumber(col);
            currCell = eb.worksheet.getRow(eb.row).getCell(eb.col);
          } catch (e) {
            console.warn(`${gt} is not a valid cell address`);
            // console.log(e);
          } finally {
            reportCell(currCell);
            switchNormalMode();
          }
        });
        rl.write(currCell.address);
        return;
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
        r.defineCommand("q", {
          help: "leave current repl",
          action() {
            this.clearBufferedCommand();
            switchNormalMode();
          },
        });
        return;
      },
    },
    enter: { f: () => {} },
    return: { f: () => {} },
    down: {
      help: "move one cell down",
      f: () => {
        eb.row += 1;
      },
    },
    left: {
      help: "move one cell left",
      f: () => {
        eb.col -= 1;
      },
    },
    right: {
      help: "move one cell right",
      f: () => {
        eb.col += 1;
      },
    },
    up: {
      help: "move one cell up",
      f: () => {
        eb.row -= 1;
      },
    },
    h: {
      help: "print this help message",
      f: () => {
        Object.keys(cmds).forEach((key) => {
          console.log(key, " - ", cmds[key].help);
        });
      },
    },
  };
  function reportCell(cell) {
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
async function save(workbook, filename) {
  const ext = path.extname(filename).toLowerCase();
  if (ext === ".xlsx") {
    await workbook.xlsx.writeFile(filename);
    console.log(`sheet written to ${filename}`);
    return true;
  } else if (ext === ".csv") {
    const options = {
      formatterOptions: {
        delimiter: ";",
        quote: false,
      },
    };
    await workbook.csv.writeFile(filename, options);
    console.log(`sheet written to ${filename}`);
    return true;
  } else {
    console.log(
      "please use either .xlsx or .csv file extension to specify file format"
    );
    return false;
  }
}
function convertLetterToNumber(str) {
  str = str.toUpperCase();
  let out = 0,
    len = str.length;
  for (pos = 0; pos < len; pos++) {
    out += (str.charCodeAt(pos) - 64) * Math.pow(26, len - pos - 1);
  }
  return out;
}
