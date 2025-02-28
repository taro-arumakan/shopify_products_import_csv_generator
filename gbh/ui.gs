function onOpen() {
  SpreadsheetApp.getUi().createMenu('Custom Menu')
    .addItem('Generate Product Import CSV', 'promptUserForSheetAndHeaders')
    .addToUi();
}

function promptUserForSheetAndHeaders() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const sheetNames = sheets.map(sheet => sheet.getName());

  // Prompt for source sheet selection
  const sheetPrompt = ui.prompt('Select the source sheet from the following list:\n' + sheetNames.join('\n'));
  const sourceSheetName = sheetPrompt.getResponseText().trim();

  if (!sourceSheetName || !sheetNames.includes(sourceSheetName)) {
    ui.alert('Invalid source sheet name. Please select a valid sheet from the list.');
    return;
  }

  // Prompt for number of header rows to skip
  const headerPrompt = ui.prompt('Enter the number of header rows to skip:');
  const headerRowsToSkip = parseInt(headerPrompt.getResponseText(), 10);

  if (isNaN(headerRowsToSkip)) {
    ui.alert('Number of header rows to skip must be a number.');
    return;
  }

  // Call the function to create the CSV
  Logger.log('Executing createProductImportCsvSheet');
  createProductImportCsvSheet(sourceSheetName, headerRowsToSkip);
}
