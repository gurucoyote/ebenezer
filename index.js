const path = require("path");
const readline = require("readline");
const Excel = require("exceljs");
const FormulaParser = require("hot-formula-parser").Parser;
async function run() {
  const parser = new FormulaParser();
  const workbook = new Excel.Workbook();
  const file = process.argv[2];
  const ext = path.extname(file).toLowerCase();
  console.log(`opening file ${file}`);
  if (ext === ".xlsx") {
    await workbook.xlsx.readFile(file);
  } else if (ext === ".csv") {
    const options = {
      parserOptions: {
        delimiter: ";",
        quote: false,
      },
    };
    await workbook.csv.readFile(file, options);
  } else {
    console.log(`extension is: ${ext}`);
    console.log(
      "please provide a filename with either .xlsx or .csv extension."
    );
    process.exit(1);
  }
  var worksheet = workbook.getWorksheet(1);
  parser.on("callCellValue", function (cellCoord, done) {
    if (worksheet.getCell(cellCoord.label).formula) {
      done(parser.parse(worksheet.getCell(cellCoord.label).formula).result);
    } else {
      done(worksheet.getCell(cellCoord.label).value);
    }
  });

  parser.on("callRangeValue", function (startCellCoord, endCellCoord, done) {
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
        colFragment.push(worksheet.getRow(row + 1).getCell(col + 1).value);
      }

      fragment.push(colFragment);
    }

    if (fragment) {
      done(fragment);
    }
  });
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
    historySize: 0, // so arrow up/down don't enter text ^^
  });
  rl.input.setEncoding("utf-8");
  rl.input.setRawMode(true);
  const defaultKPL = await rl.input.listeners("keypress");
  // console.log(defaultKPL[0].toString( ));
  rl.input.removeAllListeners("keypress");
  let rownumber = 1;
  let colnumber = 1;
  const kp = (_, key) => {
    switch (key.name) {
      case "q":
        process.exit();
      case "i":
        // enter line edit mode
        rl.input.removeAllListeners("keypress");
        // restore default kp listener
        defaultKPL.map((f) => {
          rl.input.on("keypress", f);
        });
        rl.question("> ", function (answer) {
          console.log("User entered: ", answer);
          // return to watching single keys
          rl.input.setRawMode(true);
          rl.input.removeAllListeners("keypress");
          rl.input.on("keypress", kp);
        });
        // provide default anser that can be edited
        var cel = worksheet.getRow(rownumber).getCell(colnumber);
        rl.write(getCellResult(worksheet, cel.address));
        // return from the function, so that the latter code won't be executed
        return;
      case "enter":
      case "return":
      case "down":
        rownumber += 1;
        break;
      case "left":
        colnumber -= 1;
        break;
      case "right":
        colnumber += 1;
        break;
      case "up":
        rownumber -= 1;
        break;
      default:
        // console.log("key", key);
        return;
    }
    rownumber = rownumber < 1 ? 1 : rownumber;
    colnumber = colnumber < 1 ? 1 : colnumber;
    var cel = worksheet.getRow(rownumber).getCell(colnumber);
    console.log(cel.address, getCellResult(worksheet, cel.address));
  };

  rl.input.on("keypress", kp);

  function getCellResult(worksheet, cellLabel) {
    if (worksheet.getCell(cellLabel).formula) {
      return parser.parse(worksheet.getCell(cellLabel).formula).result;
    } else {
      return worksheet.getCell(cellLabel).value;
    }
  }
}
run();
