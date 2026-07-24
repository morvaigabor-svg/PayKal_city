/**
 * Main App Controller & Web App Entry Point
 */

/**
 * Web App belépési pont
 */
function doGet() {
  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle(APP.NAME);
}

/**
 * HTML komponensek és modulok beemelése
 */
function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}

/* ==========================================
   FRONTEND ÁLTAL HÍVOTT WRAPPER FÜGGVÉNYEK
   ========================================== */

function getSettings() {
  return getSettingsData();
}

function getExpenseId(costCenter) {
  return generateExpenseId(costCenter);
}

function saveExpense(expense, imageUrls, expenseId, gpsCoords) {
  return saveExpenseData(expense, imageUrls, expenseId, gpsCoords);
}

function saveIncome(incomeData) {
  return saveIncomeData(incomeData);
}

function saveTransfer(transferData) {
  return saveTransferData(transferData);
}

function getPayKalDashboardData(timeFilter) {
  return getPayKalDashboardDataImpl(timeFilter);
}