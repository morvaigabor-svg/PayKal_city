/**
 * Beállítások és KM Tagok beolvasása az aktív csoport számára
 */
function getSettingsData() {
  try {
    const auth = getUserAuth();
    if (!auth.authorized) throw new Error("Nincs jogosultságod az adatok lekéréséhez!");

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    
    // 1. Beállítások munkalap beolvasása
    const settingsSheet = ss.getSheetByName(APP.SHEETS.SETTINGS);
    const costCenters = [], paymentMethods = [], incomePurposes = [], incomePaymentMethods = [];
    
    if (settingsSheet && settingsSheet.getLastRow() >= 2) {
      const data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 4).getValues();
      
      data.forEach(row => {
        if (row[0] && String(row[0]).trim() !== "") costCenters.push(String(row[0]).trim());
        if (row[1] && String(row[1]).trim() !== "") paymentMethods.push(String(row[1]).trim());
        if (row[2] && String(row[2]).trim() !== "") incomePurposes.push(String(row[2]).trim());
        if (row[3] && String(row[3]).trim() !== "") incomePaymentMethods.push(String(row[3]).trim());
      });
    }

    // 2. KM Tagok munkalap beolvasása (A: Név | B: Csoport ID)
    const membersSheet = ss.getSheetByName(APP.SHEETS.MEMBERS);
    const payers = [];
    
    if (membersSheet && membersSheet.getLastRow() >= 2) {
      const mData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 2).getValues();
      const userGroup = String(auth.csoportId).trim();
      
      mData.forEach(row => {
        const memberName = row[0] ? String(row[0]).trim() : "";
        const groupVal = row[1] ? String(row[1]).trim() : "";
        
        if (memberName !== "" && groupVal === userGroup) {
          payers.push(memberName);
        }
      });
    }

    return {
      costCenters: [...new Set(costCenters)],
      paymentMethods: [...new Set(paymentMethods)],
      incomePaymentMethods: [...new Set(incomePaymentMethods)],
      incomePurposes: [...new Set(incomePurposes)],
      payers: [...new Set(payers)]
    };

  } catch (error) {
    if (typeof writeLog === "function") writeLog("ERROR", "Settings", error.message);
    throw error;
  }
}