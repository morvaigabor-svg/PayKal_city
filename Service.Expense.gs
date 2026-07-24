/**
 * ExpenseTracker Mentési Szolgáltatások
 * (Kiadás, Bevétel és Pénzmozgások rögzítése)
 */

/* ==========================================
   1. KIADÁS MENTÉSI SZOLGÁLTATÁS
   ========================================== */
function saveExpenseData(expense, imageUrls = [], customId = null, gpsCoords = null) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(APP.SHEETS.EXPENSES);

    if (!sheet) {
      throw new Error("Költségek munkalap nem található");
    }

    const id = customId || generateExpenseId(expense.costCenter);

    // Beküldő Google e-mail címének kiolvasása
    const userEmail = Session.getActiveUser().getEmail() || "Ismeretlen";

    sheet.appendRow([
      expense.date,          // A oszlop: Dátum
      expense.costCenter,    // B oszlop: Költséghely
      expense.amount,        // C oszlop: Összeg
      expense.paymentMethod, // D oszlop: Fizetési mód
      "",                    // E oszlop: Blokk (RichText-tel töltjük fel)
      expense.comment,       // F oszlop: Megjegyzés
      userEmail,             // G oszlop: Beküldő e-mail címe
      "",                    // H oszlop: GPS (RichText-tel töltjük fel)
      new Date(),            // I oszlop: Rögzítés ideje
      id                     // J oszlop: Egyedi azonosító
    ]);

    const lastRow = sheet.getLastRow();

    // 1. Blokk képek linkelése (E oszlop = 5. oszlop)
    if (imageUrls && imageUrls.length > 0) {
      let labels = [];
      for (let i = 0; i < imageUrls.length; i++) {
        labels.push(String(i + 1));
      }
      
      const fullText = labels.join(" | "); 
      let richTextBuilder = SpreadsheetApp.newRichTextValue().setText(fullText);

      let currentOffset = 0;
      for (let i = 0; i < imageUrls.length; i++) {
        let label = String(i + 1);
        let start = currentOffset;
        let end = start + label.length;

        richTextBuilder.setLinkUrl(start, end, imageUrls[i]);
        currentOffset = end + 3; 
      }

      sheet.getRange(lastRow, 5).setRichTextValue(richTextBuilder.build());
    }

    // 2. GPS koordináta linkelése (H oszlop = 8. oszlop)
    if (gpsCoords && gpsCoords.lat && gpsCoords.lng) {
      const mapUrl = "https://maps.google.com/?q=" + gpsCoords.lat + "," + gpsCoords.lng;
      const gpsRichText = SpreadsheetApp.newRichTextValue()
        .setText("📍 Térkép")
        .setLinkUrl(mapUrl)
        .build();

      sheet.getRange(lastRow, 8).setRichTextValue(gpsRichText);
    }

    writeLog("INFO", "Expense", "Mentés sikeres: " + id);

    return {
      success: true,
      id: id,
      message: "Mentés sikeres"
    };

  } catch (error) {
    writeLog("ERROR", "Expense", error.message);
    throw error;
  }
}

/* ==========================================
   2. BEVÉTEL MENTÉSI SZOLGÁLTATÁS
   ========================================== */
function saveIncomeData(incomeData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(APP.SHEETS.INCOME || "Bevételek");

    if (!sheet) {
      throw new Error("A Bevételek munkalap nem található!");
    }

    const userEmail = Session.getActiveUser().getEmail() || "Ismeretlen";
    const dateStr = incomeData.date;
    const purpose = incomeData.purpose;
    const paymentMethod = incomeData.paymentMethod;
    const amount = Number(incomeData.amount);
    const comment = incomeData.comment || "";
    const payers = incomeData.payers;

    if (!payers || !Array.isArray(payers) || payers.length === 0) {
      throw new Error("Legalább egy befizetőt meg kell adni!");
    }

    const now = new Date(); // Pontos rögzítési időbélyeg
    const timeStamp = Utilities.formatDate(now, APP.TIMEZONE || "Europe/Budapest", "yyyyMMdd");
    const lastRow = sheet.getLastRow();

    const rowsToAppend = [];
    const ids = [];

    payers.forEach((payer, idx) => {
      const seq = String(lastRow + idx).padStart(3, "0");
      const id = "INC-" + timeStamp + "-" + seq;
      ids.push(id);

      // Oszlopsorrend összerakása (9 oszlop):
      rowsToAppend.push([
        dateStr,         // A: Dátum
        purpose,         // B: Befizetés célja
        amount,          // C: Összeg (forintban)
        paymentMethod,   // D: Befizetés módja
        payer,           // E: Befizető neve
        comment,         // F: Megjegyzés
        id,              // G: Fizetés egyedi azonosítója
        userEmail,       // H: Email címe a rögzítő személynek
        now              // I: Rögzítés ideje (ÉÉÉÉ.MM.DD HH:MM:SS)
      ]);
    });

    if (rowsToAppend.length > 0) {
      const startRow = lastRow + 1;
      // 9 oszlop terjedelmű írás a lapra
      sheet.getRange(startRow, 1, rowsToAppend.length, 9).setValues(rowsToAppend);
    }

    if (typeof writeLog === "function") {
      writeLog("INFO", "Income", "Bevétel mentés sikeres (" + payers.length + " fő): " + ids.join(", "));
    }

    return {
      success: true,
      count: payers.length,
      message: "Sikeresen rögzítve " + payers.length + " fő része!"
    };

  } catch (error) {
    if (typeof writeLog === "function") {
      writeLog("ERROR", "Income", error.message);
    }
    throw error;
  }
}

/* ==========================================
   3. PÉNZMOZGÁS MENTÉSI SZOLGÁLTATÁS
   ========================================== */
function saveTransferData(transferData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheetName = "Pénzmozgások";
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error("A Pénzmozgások munkalap nem található!");
    }

    const userEmail = Session.getActiveUser().getEmail() || "Ismeretlen";
    const dateStr = transferData.date;
    const type = transferData.type;
    const amount = Number(transferData.amount);
    const comment = transferData.comment || "";

    if (!dateStr || !amount || amount <= 0) {
      throw new Error("Érvénytelen dátum vagy összeg!");
    }

    // Egyedi azonosító generálása és pontos rögzítési időbélyeg
    const now = new Date();
    const timeStamp = Utilities.formatDate(now, APP.TIMEZONE || "Europe/Budapest", "yyyyMMdd");
    const lastRow = sheet.getLastRow();
    const seq = String(lastRow).padStart(3, "0");
    const transferId = "TRF-" + timeStamp + "-" + seq;

    // Új sor hozzáadása a munkalaphoz (G oszlopban a pontos rögzítési idővel)
    sheet.appendRow([
      dateStr,     // A: Dátum
      type,        // B: Tranzakció típusa
      amount,      // C: Összeg (forintban)
      comment,     // D: Megjegyzés
      transferId,  // E: Tranzakció egyedi azonosítója
      userEmail,   // F: Email címe a rögzítő személynek
      now          // G: Rögzítés ideje (ÉÉÉÉ.MM.DD HH:MM:SS)
    ]);

    if (typeof writeLog === "function") {
      writeLog("INFO", "Transfer", "Pénzmozgás rögzítve: " + transferId + " (" + amount + " Ft)");
    }

    return {
      success: true,
      message: "Pénzmozgás sikeresen rögzítve! (Azonosító: " + transferId + ")"
    };

  } catch (error) {
    if (typeof writeLog === "function") {
      writeLog("ERROR", "Transfer", error.message);
    }
    throw error;
  }
}