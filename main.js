/**
 * @OnlyCurrentDoc
 */

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('My Custom Menu')
      .addItem('Create Hello Sheet', 'createHelloSheet')
      .addToUi();
}

/**
 * Creates a new sheet tab named "Hello" and writes "Hello world" in cell A1.
 */
function createHelloSheet() {
  // const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SPREADSHEET_ID = '1ETps14eo10a1FTPDuGRYJIjKNeICwvb_MaCE1otK4gs';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const SHEET_NAME = "Hello";
  const MESSAGE = "Hello world2";

  // Check if the sheet already exists
  let newSheet = ss.getSheetByName(SHEET_NAME);

  // If the sheet exists, delete it first to ensure a fresh tab
  if (newSheet) {
    ss.deleteSheet(newSheet);
  }

  // Create the new sheet
  newSheet = ss.insertSheet(SHEET_NAME);

  // Write the message to cell A1
  newSheet.getRange('A1').setValue(MESSAGE);

  // Optional: Alert the user that the task is complete
  // SpreadsheetApp.getUi().alert(`New tab "${SHEET_NAME}" created with the message: "${MESSAGE}"`);
}
