/* ==========================================
   1. KIADÁS MENTÉSE (L, M, N oszlop bővítéssel)
   ========================================== */
function saveExpenseData(expense, imageUrls = [], customId = null, gpsCoords = null) {
  try {
    const auth = getUserAuth();
    if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(APP.SHEETS.EXPENSES);
    const id = customId || generateExpenseId(expense.costCenter);

    // Projekt attribútumok keresése
    let projectType = "";
    let projectId = "";
    const activeProjects = getActiveProjectsData(auth.csoportId, ss);
    const matchedProj = activeProjects.find(p => p.name === expense.costCenter);
    if (matchedProj) {
      projectType = matchedProj.type;
      projectId = matchedProj.id;
    }

    sheet.appendRow([
      auth.csoportId,           // A: Csoport_ID
      expense.date,             // B: Dátum
      expense.costCenter,       // C: Költséghely
      expense.amount,           // D: Összeg
      expense.paymentMethod,    // E: Fizetési mód
      "",                       // F: Blokk link
      expense.comment,          // G: Megjegyzés
      auth.email,               // H: Beküldő email
      "",                       // I: GPS
      new Date(),               // J: Feltöltés ideje
      id,                       // K: Egyedi azonosító
      projectType,              // L: Projekt típusa
      projectId,                // M: Projekt_ID
      expense.extraCategory || "" // N: Beállítások F2:F érték
    ]);

    const lastRow = sheet.getLastRow();

    if (imageUrls && imageUrls.length > 0) {
      let labels = imageUrls.map((_, i) => String(i + 1));
      let fullText = labels.join(" | "); 
      let richTextBuilder = SpreadsheetApp.newRichTextValue().setText(fullText);

      let currentOffset = 0;
      imageUrls.forEach((url, i) => {
        let label = String(i + 1);
        let start = currentOffset;
        let end = start + label.length;
        richTextBuilder.setLinkUrl(start, end, url);
        currentOffset = end + 3; 
      });

      sheet.getRange(lastRow, 6).setRichTextValue(richTextBuilder.build());
    }

    if (gpsCoords && gpsCoords.lat && gpsCoords.lng) {
      const mapUrl = "https://maps.google.com/?q=" + gpsCoords.lat + "," + gpsCoords.lng;
      const gpsRichText = SpreadsheetApp.newRichTextValue().setText("📍 Térkép").setLinkUrl(mapUrl).build();
      sheet.getRange(lastRow, 9).setRichTextValue(gpsRichText);
    }

    writeLog("INFO", "Expense", "Kiadás mentve: " + id);
    return { success: true, id: id, message: "Mentés sikeres" };

  } catch (error) {
    writeLog("ERROR", "Expense", error.message);
    throw error;
  }
}

function generateExpenseId(costCenter) {
  const now = new Date();
  const timeStamp = Utilities.formatDate(now, APP.TIMEZONE, "yyyyMMdd");
  const randomSuffix = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  return "EXP-" + timeStamp + "-" + randomSuffix;
}

/* ==========================================
   2. BEVÉTEL MENTÉSE (K és L oszlop bővítéssel)
   ========================================== */
function saveIncomeData(incomeData) {
  try {
    const auth = getUserAuth();
    if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(APP.SHEETS.INCOME);

    // Projekt attribútumok kikeresése a cél (projekt neve) alapján
    let projectType = "";
    let projectId = "";
    if (incomeData.purpose) {
      const activeProjects = getActiveProjectsData(auth.csoportId, ss);
      const matchedProj = activeProjects.find(p => p.name === incomeData.purpose);
      if (matchedProj) {
        projectType = matchedProj.type;
        projectId = matchedProj.id;
      }
    }

    const now = new Date();
    const timeStamp = Utilities.formatDate(now, APP.TIMEZONE, "yyyyMMdd");
    const lastRow = sheet.getLastRow();
    const rowsToAppend = [];
    const ids = [];

    incomeData.payers.forEach((payer, idx) => {
      const seq = String(lastRow + idx).padStart(3, "0");
      const id = "INC-" + timeStamp + "-" + seq;
      ids.push(id);

      rowsToAppend.push([
        auth.csoportId,           // A: Csoport_ID
        incomeData.date,          // B: Dátum
        incomeData.purpose,       // C: Bevétel célja
        Number(incomeData.amount),// D: Összeg
        incomeData.paymentMethod, // E: Befizetés típusa
        payer,                    // F: Befizető neve
        incomeData.comment || "", // G: Megjegyzés
        id,                       // H: Egyedi azonosító
        auth.email,               // I: Rögzítő Email
        now,                      // J: Rögzítés ideje
        projectType,              // K: Projekt típusa
        projectId                 // L: Projekt ID
      ]);
    });

    if (rowsToAppend.length > 0) {
      sheet.getRange(lastRow + 1, 1, rowsToAppend.length, 12).setValues(rowsToAppend);
    }

    writeLog("INFO", "Income", "Bevétel rögzítve: " + ids.join(", "));
    return { success: true, count: incomeData.payers.length };

  } catch (error) {
    writeLog("ERROR", "Income", error.message);
    throw error;
  }
}

/* ==========================================
   3. PÉNZMOZGÁS MENTÉSE (I, J, K oszlop bővítéssel)
   ========================================== */
function saveTransferData(transferData) {
  try {
    const auth = getUserAuth();
    if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(APP.SHEETS.TRANSFERS);

    // Projekt attribútumok kikeresése (ha van kiválasztva projekt)
    let projectType = "";
    let projectId = "";
    if (transferData.project) {
      const activeProjects = getActiveProjectsData(auth.csoportId, ss);
      const matchedProj = activeProjects.find(p => p.name === transferData.project);
      if (matchedProj) {
        projectType = matchedProj.type;
        projectId = matchedProj.id;
      }
    }

    const now = new Date();
    const timeStamp = Utilities.formatDate(now, APP.TIMEZONE, "yyyyMMdd");
    const seq = String(sheet.getLastRow()).padStart(3, "0");
    const transferId = "TRF-" + timeStamp + "-" + seq;

    sheet.appendRow([
      auth.csoportId,             // A: Csoport_ID
      transferData.date,          // B: Dátum
      transferData.type,          // C: Pénzmozgás típusa
      Number(transferData.amount),// D: Összeg
      transferData.comment || "", // E: Megjegyzés
      transferId,                 // F: Egyedi azonosító
      auth.email,                 // G: Rögzítő Email
      now,                        // H: Rögzítés ideje
      transferData.project || "", // I: Projekt neve
      projectType,                // J: Projekt típusa
      projectId                   // K: Projekt ID
    ]);

    writeLog("INFO", "Transfer", "Pénzmozgás rögzítve: " + transferId);
    return { success: true, message: "Pénzmozgás sikeresen rögzítve!" };

  } catch (error) {
    writeLog("ERROR", "Transfer", error.message);
    throw error;
  }
}