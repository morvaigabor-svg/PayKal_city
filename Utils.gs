/**
 * Egyedi kiadás azonosító (ID) generálása
 * Formátum: ÉÉÉÉHHNN-Költséghely-Sorszám (pl. 20260723-Rendezvény-005)
 */
function generateExpenseId(costCenter) {
  const now = new Date();

  const date = Utilities.formatDate(
    now,
    APP.TIMEZONE,
    "yyyyMMdd"
  );

  const sheet = SpreadsheetApp
    .openById(CONFIG.SHEET_ID)
    .getSheetByName(APP.SHEETS.EXPENSES);

  const lastRow = sheet.getLastRow();

  const sequence = String(lastRow).padStart(3, "0");

  return date + "-" + costCenter + "-" + sequence;
}