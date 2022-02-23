const FormulaParser = require('hot-formula-parser').Parser;
const parser = new FormulaParser();

const Excel = require('exceljs');
const workbook = new Excel.Workbook();

function getCellResult(worksheet, cellLabel) {
  if (worksheet.getCell(cellLabel).formula) {
    return parser.parse(worksheet.getCell(cellLabel).formula).result;
  } else {
    return worksheet.getCell(cellLabel).value;
  }
}

workbook.xlsx.readFile('./doc.xlsx').then(() => {
  var worksheet = workbook.getWorksheet(1);

  parser.on('callCellValue', function(cellCoord, done) {
    if (worksheet.getCell(cellCoord.label).formula) {
      done(parser.parse(worksheet.getCell(cellCoord.label).formula).result);
    } else {
      done(worksheet.getCell(cellCoord.label).value);
    }
  });

  parser.on('callRangeValue', function(startCellCoord, endCellCoord, done) {
    var fragment = [];

    for (var row = startCellCoord.row.index; row <= endCellCoord.row.index; row++) {
      var colFragment = [];

      for (var col = startCellCoord.column.index; col <= endCellCoord.column.index; col++) {
        colFragment.push(worksheet.getRow(row + 1).getCell(col + 1).value);
      }

      fragment.push(colFragment);
    }

    if (fragment) {
      done(fragment);
    }
  });

	console.log('value of c13 beforea:');
  console.log(getCellResult(worksheet, 'C13'));
  worksheet.getCell('C12').value = 500;
	console.log('value of c13 aftera:');
  console.log(getCellResult(worksheet, 'C13'));
});
