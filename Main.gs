/**
 * Web App belépési pont
 */
/**
 * Web App belépési pont - Mobilbarát és fiókkonfliktus-védett
 */
function doGet(e) {
  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle(APP.NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Felhasználó azonosítása és pontos koordinátori jogosultságok felépítése
 */
/**
 * 1 db összefogó inicializálás az azonnali induláshoz
 */
function getInitialAppState() {
  const auth = getUserAuth();
  if (!auth.authorized) {
    return { auth: auth };
  }

  const settings = getSettingsData();
  const dashboard = getPayKalDashboardDataImpl("1H", "ALL", auth.csoportId, "LEADERSHIP_GROUP");

  return {
    auth: auth,
    settings: settings,
    dashboard: dashboard
  };
}

/**
 * Felhasználó azonosítása (Mobilos fiókváltási hiba kivédésével)
 */
function getUserAuth() {
  try {
    let activeEmail = "";
    try { activeEmail = Session.getActiveUser().getEmail(); } catch (e) {}
    if (!activeEmail) {
      try { activeEmail = Session.getEffectiveUser().getEmail(); } catch (e) {}
    }

    const scriptUrl = ScriptApp.getService().getUrl();
    const switchUrl = "https://accounts.google.com/AccountChooser?continue=" + encodeURIComponent(scriptUrl);

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const usersSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.USERS) ? APP.SHEETS.USERS : "Felhasználók";
    const usersSheet = ss.getSheetByName(usersSheetName);

    if (!usersSheet || usersSheet.getLastRow() < 2) {
      throw new Error("A 'Felhasználók' munkalap nem található vagy üres!");
    }

    const usersData = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 4).getValues();
    const cleanEmail = String(activeEmail || "").trim().toLowerCase();

    // 1. Keresés az email alapján
    let matchedRow = null;
    if (cleanEmail !== "") {
      matchedRow = usersData.find(row => String(row[0] || "").trim().toLowerCase() === cleanEmail);
    }

    // 2. Ha az email nincs a munkalapon (vagy a Google elrejtette a domainen kívüli cimet)
    if (!matchedRow) {
      return {
        authorized: false,
        needsAuth: false,
        email: cleanEmail || "Külső / Nem azonosított fiók",
        switchAccountUrl: switchUrl,
        message: cleanEmail 
          ? "A(z) " + cleanEmail + " email cím nincs engedélyezve a rendszernyilvántartásban." 
          : "A Google nem adta át a fiókadatokat (Domainen kívüli fiók), vagy a fiók nincs regisztrálva."
      };
    }

    const finalEmail = String(matchedRow[0] || "").trim();
    const userName = String(matchedRow[1] || "").trim();
    const userGroup = String(matchedRow[2] || "").trim() || "KMCS101";
    const userRole = String(matchedRow[3] || "").trim() || "VEZETO";

    const normalizedRole = userRole.toUpperCase().replace(/Á/g, "A").replace(/É/g, "E");
    const isCoordinator = normalizedRole.includes("KOORDIN");

    // Város és csoportok feltérképezése
    const groupsSheet = ss.getSheetByName("Csoportok");
    let userCity = "Budapest";
    const cityGroups = [];

    if (groupsSheet && groupsSheet.getLastRow() >= 2) {
      const groupsData = groupsSheet.getRange(2, 1, groupsSheet.getLastRow() - 1, 2).getValues();
      const groupToCityMap = {};
      groupsData.forEach(row => {
        const gId = String(row[0] || "").trim();
        const gCity = String(row[1] || "").trim();
        if (gId) groupToCityMap[gId] = gCity;
      });
      if (userGroup && groupToCityMap[userGroup]) userCity = groupToCityMap[userGroup];
      if (userCity) {
        Object.keys(groupToCityMap).forEach(gId => {
          if (groupToCityMap[gId].toLowerCase() === userCity.toLowerCase()) cityGroups.push(gId);
        });
      }
    }

    let coordinatorOptions = [];
    if (isCoordinator) {
      coordinatorOptions.push({ id: userGroup, viewType: "LEADERSHIP_GROUP", name: "👑 Saját vezetői csoport (" + userGroup + ")" });
      if (userCity) {
        coordinatorOptions.push({ id: "CITY_SUMMARY", viewType: "CITY_SUMMARY", name: "📊 " + userCity + " - Városi összesítő nézet" });
      }
      cityGroups.forEach(gId => {
        if (gId !== userGroup) {
          coordinatorOptions.push({ id: gId, viewType: "GROUP_READONLY", name: "📁 " + gId + " (Csoportnézet)" });
        }
      });
    }

    return {
      authorized: true,
      email: finalEmail,
      name: userName,
      csoportId: userGroup,
      role: userRole,
      isCoordinator: isCoordinator,
      userCity: userCity,
      coordinatorOptions: coordinatorOptions
    };

  } catch (error) {
    const scriptUrl = ScriptApp.getService().getUrl();
    return {
      authorized: false,
      needsAuth: false,
      email: "Azonosítási hiba történt",
      switchAccountUrl: "https://accounts.google.com/AccountChooser?continue=" + encodeURIComponent(scriptUrl),
      message: "Rendszerhiba történt a bejelentkezés során."
    };
  }
}

/* --- API WRAPPER FÜGGVÉNYEK --- */
function getSettings() { return getSettingsData(); }
function getExpenseId(costCenter) { return generateExpenseId(costCenter); }
function saveExpense(expense, imageUrls, expenseId, gpsCoords) { return saveExpenseData(expense, imageUrls, expenseId, gpsCoords); }
function saveIncome(incomeData) { return saveIncomeData(incomeData); }
function saveTransfer(transferData) { return saveTransferData(transferData); }
function getPayKalDashboardData(timeFilter, selectedProject, targetCsoportId, viewType) { 
  return getPayKalDashboardDataImpl(timeFilter, selectedProject, targetCsoportId, viewType); 
}