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
 * Felhasználó azonosítása és jogosultság ellenőrzése
 */
/**
 * Felhasználó azonosítása és jogosultság ellenőrzése
 */
function getUserAuth() {
  try {
    const activeEmail = Session.getActiveUser().getEmail();
    const scriptUrl = ScriptApp.getService().getUrl();
    const switchUrl = "https://accounts.google.com/AccountChooser?continue=" + encodeURIComponent(scriptUrl);

    // 1. HA ELSŐ BELÉPÉS (A Google még nem adta át az e-mailt)
    if (!activeEmail || activeEmail.trim() === "") {
      return {
        authorized: false,
        needsAuth: true, // JELZÉS: Automatikus átirányítás szükséges!
        switchAccountUrl: switchUrl
      };
    }

    // 2. 'Felhasználók' lap ellenőrzése
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const usersSheet = ss.getSheetByName(APP.SHEETS.USERS);
    
    let isAuthorized = false;
    let userGroup = null;

    if (usersSheet && usersSheet.getLastRow() >= 2) {
      const usersData = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 4).getValues();
      
      for (let i = 0; i < usersData.length; i++) {
        const sheetEmail = String(usersData[i][0]).trim().toLowerCase();
        const currentEmail = activeEmail.trim().toLowerCase();
        
        if (sheetEmail === currentEmail) {
          isAuthorized = true;
          userGroup = usersData[i][2];
          break;
        }
      }
    }

    // 3. HA BENNE VAN A TÁBLÁZATBAN ➔ BEENGETJÜK!
    if (isAuthorized) {
      return {
        authorized: true,
        email: activeEmail,
        csoportId: userGroup
      };
    }

    // 4. HA ISMERT AZ E-MAIL, DE NINCS A TÁBLÁZATBAN ➔ LETILTÓ KÁRTYA!
    return {
      authorized: false,
      needsAuth: false,
      email: activeEmail,
      switchAccountUrl: switchUrl
    };

  } catch (error) {
    const scriptUrl = ScriptApp.getService().getUrl();
    return {
      authorized: false,
      needsAuth: false,
      email: "Azonosítási hiba történt",
      switchAccountUrl: "https://accounts.google.com/AccountChooser?continue=" + encodeURIComponent(scriptUrl)
    };
  }
}

/* --- API WRAPPER FÜGGVÉNYEK --- */
function getSettings() { return getSettingsData(); }
function getExpenseId(costCenter) { return generateExpenseId(costCenter); }
function saveExpense(expense, imageUrls, expenseId, gpsCoords) { return saveExpenseData(expense, imageUrls, expenseId, gpsCoords); }
function saveIncome(incomeData) { return saveIncomeData(incomeData); }
function saveTransfer(transferData) { return saveTransferData(transferData); }
function getPayKalDashboardData(timeFilter) { return getPayKalDashboardDataImpl(timeFilter); }