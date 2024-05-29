// not currently working

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Adds a custom menu to the spreadsheet.
  ui.createMenu('Batch Processing')
    .addItem('Start Processing', 'showDialog')
    .addToUi();
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showModalDialog(html, 'Batch Processing Parameters');
}
