/**
 * Web App belépési pont
 */
function doGet() {
  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle(APP.NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Belépő felhasználó azonosítása és jogosultság ellenőrzése
 */
function getUserAuth() {
  const activeEmail = Session.getActiveUser().getEmail();
  
  if (!activeEmail) {
    return { authorized: false, reason: "NO_EMAIL" };
  }

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const userSheet = ss.getSheetByName(APP.SHEETS.USERS);
  
  if (!userSheet) {
    throw new Error("A 'Felhasználók' munkalap nem található!");
  }

  const data = userSheet.getDataRange().getValues();
  // Feltételezett fejlécek: A: Email, B: Név, C: Csoport_ID, D: Szerepkör
  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][0]).trim().toLowerCase();
    if (email === activeEmail.toLowerCase()) {
      return {
        authorized: true,
        email: activeEmail,
        name: data[i][1],
        csoportId: data[i][2],
        role: data[i][3]
      };
    }
  }

  // Ha nem található a Felhasználók lapon: Fiókváltó link generálása
  const scriptUrl = ScriptApp.getService().getUrl();
  const switchAccountUrl = "https://accounts.google.com/AccountChooser?continue=" + encodeURIComponent(scriptUrl);

  return {
    authorized: false,
    email: activeEmail,
    switchAccountUrl: switchAccountUrl
  };
}

/* --- API WRAPPER FÜGGVÉNYEK --- */
function getSettings() { return getSettingsData(); }
function getExpenseId(costCenter) { return generateExpenseId(costCenter); }
function saveExpense(expense, imageUrls, expenseId, gpsCoords) { return saveExpenseData(expense, imageUrls, expenseId, gpsCoords); }
function saveIncome(incomeData) { return saveIncomeData(incomeData); }
function saveTransfer(transferData) { return saveTransferData(transferData); }
function getPayKalDashboardData(timeFilter) { return getPayKalDashboardDataImpl(timeFilter); }