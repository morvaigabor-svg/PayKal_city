/**
 * ExpenseTracker napló szolgáltatás
 */

function writeLog(level, module, message) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(APP.SHEETS.LOG);

  sheet.appendRow([
    new Date(),
    level,
    module,
    message
  ]);
}

function testLog() {
  writeLog(
    "INFO",
    "Test",
    "Log rendszer működik"
  );
}