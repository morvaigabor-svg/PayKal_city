/* ==========================================
   1. KIADÁS MENTÉSE (Csoport_ID az A oszlopban)
   ========================================== */
function saveExpenseData(expense, imageUrls = [], customId = null, gpsCoords = null) {
  try {
    const auth = getUserAuth();
    if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(APP.SHEETS.EXPENSES);
    const id = customId || generateExpenseId(expense.costCenter);

    sheet.appendRow([
      auth.csoportId,        // A oszlop: Csoport_ID
      expense.date,          // B oszlop: Dátum
      expense.costCenter,    // C oszlop: Költséghely
      expense.amount,        // D oszlop: Összeg
      expense.paymentMethod, // E oszlop: Fizetési mód
      "",                    // F oszlop: Blokk link
      expense.comment,       // G oszlop: Megjegyzés
      auth.email,            // H oszlop: Beküldő email
      "",                    // I oszlop: GPS
      new Date(),            // J oszlop: Feltöltés ideje
      id                     // K oszlop: Egyedi azonosító
    ]);

    const lastRow = sheet.getLastRow();

    // Blokk linkek beszúrása (F oszlop = 6. oszlop)
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

    // GPS link beszúrása (I oszlop = 9. oszlop)
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

/* ==========================================
   2. BEVÉTEL MENTÉSE (Csoport_ID az A oszlopban)
   ========================================== */
function saveIncomeData(incomeData) {
  try {
    const auth = getUserAuth();
    if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(APP.SHEETS.INCOME);

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
        now                       // J: Rögzítés ideje
      ]);
    });

    if (rowsToAppend.length > 0) {
      sheet.getRange(lastRow + 1, 1, rowsToAppend.length, 10).setValues(rowsToAppend);
    }

    writeLog("INFO", "Income", "Bevétel rögzítve: " + ids.join(", "));
    return { success: true, count: incomeData.payers.length };

  } catch (error) {
    writeLog("ERROR", "Income", error.message);
    throw error;
  }
}

/* ==========================================
   3. PÉNZMOZGÁS MENTÉSE (Csoport_ID az A oszlopban)
   ========================================== */
function saveTransferData(transferData) {
  try {
    const auth = getUserAuth();
    if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(APP.SHEETS.TRANSFERS);

    const now = new Date();
    const timeStamp = Utilities.formatDate(now, APP.TIMEZONE, "yyyyMMdd");
    const seq = String(sheet.getLastRow()).padStart(3, "0");
    const transferId = "TRF-" + timeStamp + "-" + seq;

    sheet.appendRow([
      auth.csoportId,           // A: Csoport_ID
      transferData.date,        // B: Dátum
      transferData.type,        // C: Pénzmozgás típusa
      Number(transferData.amount),// D: Összeg
      transferData.comment || "",// E: Megjegyzés
      transferId,               // F: Egyedi azonosító
      auth.email,               // G: Rögzítő Email
      now                       // H: Rögzítés ideje
    ]);

    writeLog("INFO", "Transfer", "Pénzmozgás rögzítve: " + transferId);
    return { success: true, message: "Pénzmozgás sikeresen rögzítve!" };

  } catch (error) {
    writeLog("ERROR", "Transfer", error.message);
    throw error;
  }
}