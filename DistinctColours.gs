
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Distinct colours",
    functionName : "setDistinctColours"
  }];
  sheet.addMenu("Script Center Menu", entries);
};

function setDistinctColours() {
  // Matplotlib colours
  var colours = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'];
  var valueColours = new Map();
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selection = activeSheet.getSelection()
  var range = selection.getActiveRange();
  var values = range.getValues();
  var numRows = range.getNumRows();
  var numColumns = range.getNumColumns();

  for (var rowNum = 0; rowNum <= numRows - 1; rowNum++) {
    for (var colNum = 0; colNum <= numColumns - 1; colNum++) {
      var cellValue = values[rowNum][colNum];
      if(!valueColours.has(cellValue)) {
        valueColours.set(cellValue, colours.shift())
      }
      range.getCell(rowNum + 1, colNum + 1).setBackground(valueColours.get(cellValue));
    }
  }
}
