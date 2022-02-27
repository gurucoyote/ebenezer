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
function convertLetterToNumber(str) {
  str = str.toUpperCase();
  let out = 0,
    len = str.length;
  for (pos = 0; pos < len; pos++) {
    out += (str.charCodeAt(pos) - 64) * Math.pow(26, len - pos - 1);
  }
  return out;
}
async function run() {
  const eb = {};
  eb.filename = process.argv[2];
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

  function kp(_, key) {
    switch (key.name) {
      case "q":
        process.exit();
      case "i":
        switchInsertMode();
        const cell = eb.worksheet.getRow(eb.row).getCell(eb.col);
        rl.question("> ", function (answer) {
          // console.log("// User entered: ", answer);
          // write the edit to the sheet
          if (answer.charAt(0) === "=") {
            cell.formula = answer.substring(1); // substring(1) = from 2nd char to end of string
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
      case "s":
        switchInsertMode();
        rl.question("filename> ", function (fn) {
          save(workbook, fn);
          switchNormalMode();
        });
        rl.write(eb.filename);
        return;
      case "g":
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
      case "r":
        rl.input.removeAllListeners("keypress");
        const r = repl.start({
          prompt: "repl>",
          ignoreUndefined: true,
        });
        r.context.eb = eb;
        r.defineCommand("q", {
          help: "leave current repl",
          action() {
            this.clearBufferedCommand();
            switchNormalMode();
          },
        });
        // we might want to prevent the user to totally exit the app here
        // delete r.commands.exit;
        return;
      case "enter":
      case "return":
      case "down":
        eb.row += 1;
        break;
      case "left":
        eb.col -= 1;
        break;
      case "right":
        eb.col += 1;
        break;
      case "up":
        eb.row -= 1;
        break;
      default:
        // console.log("key", key);
        return;
    }
    eb.row = eb.row < 1 ? 1 : eb.row;
    eb.col = eb.col < 1 ? 1 : eb.col;
    const cell = eb.worksheet.getRow(eb.row).getCell(eb.col);
    reportCell(cell);
  }
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
  console.log(`saving file ${filename}`);
  if (ext === ".xlsx") {
    await workbook.xlsx.writeFile(filename);
  } else if (ext === ".csv") {
    const options = {
      formatterOptions: {
        delimiter: ";",
        quote: false,
      },
    };
    await workbook.csv.writeFile(filename, options);
  } else {
    console.log(
      "please use either .xlsx or .csv file extension to specify file format"
    );
    return;
  }
  console.log(`sheet written to ${filename}`);
}
