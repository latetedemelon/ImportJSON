/**
 * Creates a custom menu in Google Sheets when the spreadsheet opens.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Factura')
        .addItem('Importar CSV', 'ImportCSV')
        .addToUi();
  } catch (e) {
    // TODO (Developer) - Handle exception
    console.log('Failed with error: %s', e.error);
  }
}

function importCSVFactura(_csvContent)
{
  var csvSheet = SpreadsheetApp.getActive().getSheetByName("CSV Factura");
  if (csvSheet)
  {
    let data = Utilities.parseCsv(_csvContent, ';');
    csvSheet.clear();
    // Determines the incoming data size.
    let numRows = data.length;
    let numColumns = data[0].length;
    csvSheet.getRange(1, 1, numRows, numColumns).setValues(data);
    return _csvContent;
  }
  else
  {
    var sheetNames = "";
    var sheets = SpreadsheetApp.getActive().getSheets();
    for (var i = 0; i < sheets.length; ++i)
    {
      sheetNames += sheets[i].getName() + ";";
    }
    return sheetNames;
  }
}

function ImportCSV() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ImportCSVDialog.html')
        .setWidth(600)
        .setHeight(425)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Select a file');
  } catch (e) {
    // TODO (Developer) - Handle exception
    console.log('Failed with error: %s', e.error);
  }
}
